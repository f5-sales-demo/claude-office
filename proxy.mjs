#!/usr/bin/env node
// claude-office — a local CORS-injecting reverse proxy that lets the Claude for
// Office (PowerPoint/Word/Excel) add-in talk to a self-hosted, Anthropic-compatible
// LLM gateway that only permits CORS from its own origin.
//
// The add-in's task pane runs in a WebView served from https://pivot.claude.ai.
// When it calls a gateway whose CORS policy rejects that origin, the browser
// aborts with "Load failed".
// This proxy listens locally, answers the CORS preflight, rewrites the outbound
// Origin so the gateway/WAF accepts it, and injects Access-Control-* headers on
// the way back. TLS is terminated with a publicly-trusted *.local-ip.sh cert on
// the hostname 127-0-0-1.local-ip.sh (which resolves to 127.0.0.1), so no local
// certificate trust step is needed.
//
// You must configure YOUR OWN gateway origin — this ships with no endpoint:
//   claude-office set-upstream https://your-gateway.internal.example
//
// Config resolution: env CLAUDE_OFFICE_UPSTREAM > config.json > (error).
import https from 'node:https';
import http from 'node:http';
import tls from 'node:tls';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import crypto from 'node:crypto';
import { execFileSync } from 'node:child_process';

const DIR = process.env.CLAUDE_OFFICE_DIR
  || path.join(process.env.XDG_CONFIG_HOME || path.join(os.homedir(), '.config'), 'claude-office');
const CONFIG_FILE = path.join(DIR, 'config.json');
const PUB_CERT = path.join(DIR, 'localip.pem');
const PUB_KEY = path.join(DIR, 'localip.key');
const SELF_CERT = path.join(DIR, 'localhost.pem');
const SELF_KEY = path.join(DIR, 'localhost.key');

// Public, non-sensitive constants (safe to ship):
const LOCALIP_CERT_URL = 'https://local-ip.sh/server.pem';
const LOCALIP_KEY_URL = 'https://local-ip.sh/server.key';
const DEFAULTS = {
  httpsPort: 8443,
  httpPort: 8080,
  hostname: '127-0-0-1.local-ip.sh', // resolves to 127.0.0.1, matches *.local-ip.sh cert
  allowedOriginSuffixes: ['pivot.claude.ai', 'claude.ai'], // add-in task pane (pivot.claude.ai) + OAuth (claude.ai)
};
const REFRESH_MS = 7 * 24 * 60 * 60 * 1000;
const HOP = new Set(['connection', 'keep-alive', 'proxy-authenticate', 'proxy-authorization',
  'te', 'trailer', 'transfer-encoding', 'upgrade']);

const log = (m) => console.log(`[claude-office] ${m}`);
const die = (m) => { console.error(`[claude-office] ${m}`); process.exit(1); };

function ensureDir() { fs.mkdirSync(DIR, { recursive: true }); }
function readConfigFile() { try { return JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8')); } catch { return {}; } }
function writeConfigFile(cfg) { ensureDir(); fs.writeFileSync(CONFIG_FILE, JSON.stringify(cfg, null, 2) + '\n', { mode: 0o600 }); }

function resolveConfig() {
  const file = readConfigFile();
  const cfg = { ...DEFAULTS, ...file };
  const raw = process.env.CLAUDE_OFFICE_UPSTREAM || file.upstream || '';
  cfg.upstream = raw ? normalizeOrigin(raw) : '';
  if (process.env.CLAUDE_OFFICE_HTTPS_PORT) cfg.httpsPort = Number(process.env.CLAUDE_OFFICE_HTTPS_PORT);
  if (process.env.CLAUDE_OFFICE_HTTP_PORT) cfg.httpPort = Number(process.env.CLAUDE_OFFICE_HTTP_PORT);
  if (process.env.CLAUDE_OFFICE_HOSTNAME) cfg.hostname = process.env.CLAUDE_OFFICE_HOSTNAME;
  return cfg;
}

function normalizeOrigin(input) {
  let u;
  try { u = new URL(input.includes('://') ? input : `https://${input}`); }
  catch { die(`invalid upstream URL: ${input}`); }
  if (u.pathname !== '/' && u.pathname !== '') {
    log(`note: using origin ${u.origin} (path "${u.pathname}" is supplied by the add-in URL instead)`);
  }
  return u.origin;
}

// ---- certificate provisioning ------------------------------------------------
function certDaysLeft(file) {
  try { return (new Date(new crypto.X509Certificate(fs.readFileSync(file)).validTo) - Date.now()) / 86400000; }
  catch { return -1; }
}

async function provisionPublicCert() {
  ensureDir();
  const fresh = fs.existsSync(PUB_CERT) && fs.existsSync(PUB_KEY) && certDaysLeft(PUB_CERT) > 30;
  if (fresh) return;
  try {
    const [cert, key] = await Promise.all([
      fetch(LOCALIP_CERT_URL).then((r) => { if (!r.ok) throw new Error(`cert HTTP ${r.status}`); return r.text(); }),
      fetch(LOCALIP_KEY_URL).then((r) => { if (!r.ok) throw new Error(`key HTTP ${r.status}`); return r.text(); }),
    ]);
    if (!cert.includes('BEGIN CERTIFICATE') || !key.includes('PRIVATE KEY')) throw new Error('unexpected download content');
    tls.createSecureContext({ cert, key }); // throws if key/cert mismatch
    fs.writeFileSync(PUB_CERT, cert);
    fs.writeFileSync(PUB_KEY, key, { mode: 0o600 });
    log('provisioned public *.local-ip.sh certificate');
  } catch (e) {
    log(`public cert fetch failed: ${e.message} (keeping existing / self-signed fallback)`);
  }
}

function provisionSelfSigned() {
  if (fs.existsSync(SELF_CERT) && fs.existsSync(SELF_KEY)) return;
  try {
    execFileSync('openssl', ['req', '-x509', '-newkey', 'rsa:2048', '-sha256', '-days', '800', '-nodes',
      '-keyout', SELF_KEY, '-out', SELF_CERT, '-subj', '/CN=localhost',
      '-addext', 'subjectAltName=DNS:localhost,IP:127.0.0.1',
      '-addext', 'extendedKeyUsage=serverAuth', '-addext', 'basicConstraints=critical,CA:FALSE'], { stdio: 'ignore' });
    fs.chmodSync(SELF_KEY, 0o600);
    log('generated self-signed localhost certificate (fallback)');
  } catch (e) {
    log(`self-signed generation skipped (openssl unavailable): ${e.message}`);
  }
}

function loadCtx(certF, keyF) {
  if (fs.existsSync(certF) && fs.existsSync(keyF)) {
    try { return tls.createSecureContext({ cert: fs.readFileSync(certF), key: fs.readFileSync(keyF) }); } catch { /* fall through */ }
  }
  return null;
}

// ---- request handling --------------------------------------------------------
function corsHeaders(req, cfg) {
  const origin = req.headers.origin;
  const allowed = origin && cfg.allowedOriginSuffixes.some((s) => {
    try { return new URL(origin).hostname.endsWith(s); } catch { return false; }
  });
  const h = {
    'access-control-allow-methods': 'GET, POST, OPTIONS',
    'access-control-allow-headers': req.headers['access-control-request-headers']
      || 'x-api-key, anthropic-version, anthropic-beta, content-type, authorization',
    'access-control-max-age': '600',
    vary: 'Origin',
  };
  if (allowed) { h['access-control-allow-origin'] = origin; h['access-control-allow-credentials'] = 'true'; }
  else if (origin) { h['access-control-allow-origin'] = origin; h['access-control-allow-credentials'] = 'true'; }
  return h;
}

function makeHandler(cfg) {
  const upstreamHost = new URL(cfg.upstream).hostname;
  const upstreamPort = new URL(cfg.upstream).port || 443;
  return (req, res) => {
    if (req.method === 'OPTIONS') { res.writeHead(204, corsHeaders(req, cfg)); res.end(); return; }
    const headers = { ...req.headers };
    for (const k of Object.keys(headers)) if (HOP.has(k)) delete headers[k];
    headers.host = upstreamHost;
    headers.origin = cfg.upstream; // an origin the gateway accepts
    delete headers.referer;

    const up = https.request(
      { host: upstreamHost, port: upstreamPort, method: req.method, path: req.url, headers, servername: upstreamHost },
      (upRes) => {
        const out = {};
        for (const [k, v] of Object.entries(upRes.headers)) {
          const lk = k.toLowerCase();
          if (HOP.has(lk) || lk.startsWith('access-control-')) continue;
          out[k] = v;
        }
        Object.assign(out, corsHeaders(req, cfg));
        res.writeHead(upRes.statusCode, out);
        upRes.pipe(res);
      });
    up.on('error', (err) => {
      log(`upstream error: ${err.message}`);
      if (!res.headersSent) res.writeHead(502, corsHeaders(req, cfg));
      res.end(JSON.stringify({ error: 'proxy_upstream_error', detail: err.message }));
    });
    req.pipe(up);
  };
}

// ---- commands ----------------------------------------------------------------
async function cmdRun() {
  const cfg = resolveConfig();
  if (!cfg.upstream) {
    die('no upstream gateway configured.\n' +
      '  Set it with:  claude-office set-upstream https://your-gateway.internal.example\n' +
      '  or export CLAUDE_OFFICE_UPSTREAM. This tool ships with no default endpoint.');
  }
  await provisionPublicCert();
  provisionSelfSigned();
  const publicCtx = loadCtx(PUB_CERT, PUB_KEY);
  const selfCtx = loadCtx(SELF_CERT, SELF_KEY);
  if (!publicCtx && !selfCtx) die('no TLS certificate available (public fetch failed and openssl missing).');
  log(`TLS: public *.local-ip.sh ${publicCtx ? 'loaded' : 'MISSING'}, self-signed localhost ${selfCtx ? 'loaded' : 'MISSING'}`);

  const def = publicCtx
    ? { cert: fs.readFileSync(PUB_CERT), key: fs.readFileSync(PUB_KEY) }
    : { cert: fs.readFileSync(SELF_CERT), key: fs.readFileSync(SELF_KEY) };
  const tlsOptions = {
    ...def,
    SNICallback: (name, cb) => cb(null, (name && name.endsWith('.local-ip.sh') && publicCtx) ? publicCtx : (selfCtx || publicCtx)),
  };
  const handler = makeHandler(cfg);
  const onClientError = (err, sock) => { log(`client error: ${err.message}`); try { sock.destroy(); } catch { /* noop */ } };

  const httpsServer = https.createServer(tlsOptions, handler).on('clientError', onClientError);
  const httpServer = http.createServer(handler).on('clientError', onClientError);
  httpsServer.listen(cfg.httpsPort, '127.0.0.1', () => log(`HTTPS on https://localhost:${cfg.httpsPort} -> ${cfg.upstream}`));
  httpServer.listen(cfg.httpPort, '127.0.0.1', () => {
    log(`HTTP  on http://localhost:${cfg.httpPort}  -> ${cfg.upstream}`);
    log(`add-in Gateway URL: https://${cfg.hostname}:${cfg.httpsPort}<your gateway's API path, e.g. /anthropic>`);
  });
  setInterval(() => { provisionPublicCert().catch(() => {}); }, REFRESH_MS).unref();
}

function cmdSetUpstream(arg) {
  if (!arg) die('usage: claude-office set-upstream https://your-gateway.internal.example');
  const origin = normalizeOrigin(arg);
  const cfg = readConfigFile();
  cfg.upstream = origin;
  writeConfigFile(cfg);
  log(`upstream set to ${origin}  (config: ${CONFIG_FILE})`);
  restartService();
}

// Best-effort (re)start of the Homebrew background service so a changed upstream
// takes effect immediately — the running proxy only reads the upstream at startup.
// `brew services restart` also starts it if it was stopped. Degrades gracefully
// when Homebrew isn't present (e.g. the binary is run standalone).
function restartService() {
  try {
    execFileSync('brew', ['services', 'restart', 'claude-office'], { stdio: 'ignore' });
    log('restarted the background service (brew) so the new upstream is live');
  } catch {
    log('to run it in the background:  brew services start claude-office');
  }
}

function cmdConfig() {
  const cfg = resolveConfig();
  const masked = { ...cfg };
  console.log(JSON.stringify(masked, null, 2));
  console.log(`# config file: ${CONFIG_FILE}${process.env.CLAUDE_OFFICE_UPSTREAM ? '\n# (upstream overridden by CLAUDE_OFFICE_UPSTREAM env)' : ''}`);
}

function cmdUrl() {
  const cfg = resolveConfig();
  console.log(`https://${cfg.hostname}:${cfg.httpsPort}<PATH>`);
  console.log('# append your gateway\'s API path (e.g. /anthropic) — the same path your gateway serves.');
}

async function cmdDoctor() {
  const cfg = resolveConfig();
  console.log(`upstream:        ${cfg.upstream || '(NOT SET — run: claude-office set-upstream <url>)'}`);
  console.log(`https/http port: ${cfg.httpsPort} / ${cfg.httpPort}`);
  console.log(`hostname:        ${cfg.hostname}`);
  console.log(`public cert:     ${fs.existsSync(PUB_CERT) ? `present, ~${Math.round(certDaysLeft(PUB_CERT))}d left` : 'absent'}`);
  console.log(`self-signed:     ${fs.existsSync(SELF_CERT) ? 'present' : 'absent'}`);
  if (cfg.upstream) {
    try { const r = await fetch(cfg.upstream, { method: 'HEAD' }); console.log(`upstream reachable: HTTP ${r.status}`); }
    catch (e) { console.log(`upstream reachable: NO (${e.message}) — are you on VPN?`); }
  }
}

function cmdHelp() {
  console.log(`claude-office — local CORS proxy for the Claude for Office add-in

Usage:
  claude-office set-upstream <url>   Configure your gateway origin (required, no default)
  claude-office run                  Start the proxy (used by 'brew services')
  claude-office url                  Print the add-in Gateway URL
  claude-office config              Show resolved config
  claude-office doctor              Diagnostics (cert status, upstream reachability)

Configure your own gateway; nothing is hardcoded:
  claude-office set-upstream https://your-gateway.internal.example`);
}

const [cmd, arg] = process.argv.slice(2);
switch (cmd || 'run') {
  case 'run': cmdRun(); break;
  case 'set-upstream': cmdSetUpstream(arg); break;
  case 'config': cmdConfig(); break;
  case 'url': cmdUrl(); break;
  case 'doctor': cmdDoctor(); break;
  case 'help': case '-h': case '--help': cmdHelp(); break;
  default: console.error(`unknown command: ${cmd}`); cmdHelp(); process.exit(1);
}
