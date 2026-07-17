# claude-office

A tiny local reverse proxy that lets the **Claude for Office** add-in (PowerPoint,
Word, Excel) talk to a **self-hosted, Anthropic-compatible LLM gateway** whose CORS
policy would otherwise reject the add-in.

## The problem it solves

The Claude for Office add-in is an Office.js web add-in whose task pane runs inside a
WebView served from `https://pivot.claude.ai`. When you point its **Gateway** connection
at an internal gateway (e.g. a LiteLLM / Open WebUI deployment behind a WAF) that only
allows CORS from its own origin, the browser blocks the request and the add-in reports
**"Could not reach gateway / Load failed"** — even though `curl` to the same gateway works
(curl doesn't enforce CORS).

`claude-office` runs on your machine and:

- answers the CORS preflight locally and injects `Access-Control-Allow-*` headers,
- rewrites the outbound `Origin` to one your gateway accepts,
- forwards everything else (streaming included) to **your** gateway,
- terminates TLS with a **publicly-trusted** certificate on `127-0-0-1.local-ip.sh`
  (a public wildcard-DNS name that resolves to `127.0.0.1`), so there is **no
  certificate to trust manually** and no `sudo`.

Nothing about any specific gateway is baked in — **you configure your own endpoint.**

## The proper fix (and why this proxy exists)

Pointing the add-in at a custom endpoint is **officially supported** — the add-in has a
built-in **Gateway** connection (base URL + token, also settable via the manifest
`gateway_url` parameter). The *only* thing that breaks is CORS, and the documented,
supported fix is server-side: **your gateway — or the WAF in front of it — must return**

- `Access-Control-Allow-Origin: https://pivot.claude.ai` (or `*`) on **every** response,
  including `GET`, `POST`, `OPTIONS`, **and error responses** (setting it only on the
  preflight is not enough), and
- `Access-Control-Allow-Headers: x-api-key, authorization, content-type, anthropic-version`
  on the preflight.

In F5 terms that's a one-line edge change — a **BIG-IP iRule** or **Distributed Cloud
HTTP-LB** response-header-injection rule — and it's LLM-agnostic (nothing about the model
protocol changes).

There is **no** client-side escape hatch: CORS is enforced by the Office WebView, the
add-in origin (`https://pivot.claude.ai`) is fixed, and there is no manifest field,
environment variable, registry key, or supported WebView flag that disables it
(`<AppDomains>` only governs in-frame navigation, not `fetch`/XHR). This is confirmed by
[Microsoft's same-origin guidance for Office Add-ins][ms-cors] and Anthropic's own gateway
docs.

So this proxy is **not** a shortcut around a missing setting — it's the correct fallback
for when you **cannot** change the gateway/WAF (e.g. it's run by an outsourced IT group).
It fixes only the **CORS + TLS** layer and forwards everything else, streaming included,
to your gateway unchanged. If you *can* get one CORS header added at the edge, do that
instead and skip the proxy entirely.

[ms-cors]: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/addressing-same-origin-policy-limitations

## Install

Prebuilt macOS binaries are published to the `f5-sales-demo/homebrew-tap` Homebrew tap.
They are **Developer ID-signed and Apple-notarized**, so they install and run without
Gatekeeper "unverified developer" prompts.

```sh
brew install f5-sales-demo/tap/claude-office
```

That one command auto-adds the tap. To add it explicitly first:

```sh
brew tap f5-sales-demo/tap
brew install claude-office
```

**Requirements:** macOS on Apple Silicon (arm64) or Intel (x64). Nothing else to install —
the binary is self-contained (no Node.js runtime required).

### Upgrade

```sh
brew upgrade claude-office
```

### Uninstall

```sh
brew services stop claude-office   # only if you started it as a service
brew uninstall claude-office
```

### Verify the signature (optional)

```sh
codesign -dvv "$(brew --prefix)/bin/claude-office"
# expect: Authority=Developer ID Application: … and flags=0x10000(runtime)
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
