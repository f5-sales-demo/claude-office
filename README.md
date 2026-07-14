# claude-office

A tiny local reverse proxy that lets the **Claude for Office** add-in (PowerPoint,
Word, Excel) talk to a **self-hosted, Anthropic-compatible LLM gateway** whose CORS
policy would otherwise reject the add-in.

## The problem it solves

The Claude for Office add-in runs inside a WebView served from `https://claude.ai`.
When you point its **Gateway** connection at an internal gateway (e.g. a LiteLLM /
Open WebUI deployment behind a WAF) that only allows CORS from its own origin, the
browser blocks the request and the add-in reports **"Could not reach gateway / Load
failed"** — even though `curl` to the same gateway works (curl doesn't enforce CORS).

`claude-office` runs on your machine and:

- answers the CORS preflight locally and injects `Access-Control-Allow-*` headers,
- rewrites the outbound `Origin` to one your gateway accepts,
- forwards everything else (streaming included) to **your** gateway,
- terminates TLS with a **publicly-trusted** certificate on `127-0-0-1.local-ip.sh`
  (a public wildcard-DNS name that resolves to `127.0.0.1`), so there is **no
  certificate to trust manually** and no `sudo`.

Nothing about any specific gateway is baked in — **you configure your own endpoint.**

## Install

```sh
brew install f5-sales-demo/tap/claude-office
```

## Configure & run

```sh
# 1. Point it at YOUR gateway origin (scheme + host only):
claude-office set-upstream https://your-gateway.internal.example

# 2. Start it as a background service (no sudo):
brew services start claude-office

# 3. Get the URL to paste into the add-in:
claude-office url
```

In the add-in's **Gateway** tab set:

- **Gateway URL:** `https://127-0-0-1.local-ip.sh:8443/<your gateway's API path>`
  (e.g. append `/anthropic` if your gateway serves the Anthropic API under `/anthropic`)
- **Token:** your gateway API key

> You must be on your VPN/network — the proxy still reaches the gateway directly; it
> only fixes the browser CORS/TLS layer.

## Commands

| Command | Description |
|---|---|
| `claude-office set-upstream <url>` | Configure your gateway origin (required; no default) |
| `claude-office run` | Start the proxy (used by `brew services`) |
| `claude-office url` | Print the add-in Gateway URL |
| `claude-office config` | Show the resolved configuration |
| `claude-office doctor` | Diagnostics: cert status and upstream reachability |

## Configuration

Resolved in this order: environment variable → config file → error.

| Setting | Env var | `config.json` key | Default |
|---|---|---|---|
| Gateway origin | `CLAUDE_OFFICE_UPSTREAM` | `upstream` | *(required)* |
| HTTPS port | `CLAUDE_OFFICE_HTTPS_PORT` | `httpsPort` | `8443` |
| HTTP port | `CLAUDE_OFFICE_HTTP_PORT` | `httpPort` | `8080` |
| Hostname | `CLAUDE_OFFICE_HOSTNAME` | `hostname` | `127-0-0-1.local-ip.sh` |

Config lives in `${XDG_CONFIG_HOME:-~/.config}/claude-office/config.json`.

## How the TLS cert works

On start, the proxy downloads the public `*.local-ip.sh` Let's Encrypt certificate
(and its published key) from `https://local-ip.sh` and serves it for the
`127-0-0-1.local-ip.sh` hostname, which resolves to `127.0.0.1`. Because that cert is
publicly trusted, the add-in's WebView accepts it with no local trust step. The cert
is refreshed automatically (weekly, and whenever it's within 30 days of expiry). A
self-signed `localhost` cert is generated as an offline fallback (served for the
`localhost` hostname via SNI).

The published `*.local-ip.sh` private key is public by design; it only terminates a
**loopback-only** listener on your own machine, so this is not a meaningful exposure.

## Security / privacy

- Ships with **no gateway endpoint, no tokens, no internal URLs** — you supply your own.
- Your gateway URL is stored only in your local `config.json` (git-ignored).
- Your API token is entered in the add-in and passed straight through; the proxy never
  stores it.

## License

MIT
