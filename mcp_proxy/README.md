# Retailer pipeline MCP front door

Claude's custom-connector check refuses the 302 that Apps Script web apps answer
with, so this proxy terminates MCP properly and forwards each JSON-RPC message
to the Apps Script deployment. All the tools and sheet logic stay in
`google_apps_script/retailers_sync.gs`; this only moves bytes.

    POST /mcp      MCP Streamable HTTP endpoint (this is the connector URL)
    GET  /health   liveness + whether the upstream and auth are configured

Secrets live on Fly, not in the repo:

    flyctl secrets set APPS_SCRIPT_URL="https://script.google.com/macros/s/.../exec" -a catalyst-retailers-mcp
    flyctl secrets set MCP_BEARER="<shared secret>" -a catalyst-retailers-mcp   # optional

Deploy after a change:

    flyctl deploy -a catalyst-retailers-mcp

When the Apps Script deployment URL changes (a fresh deployment rather than a
new version of the existing one), reset `APPS_SCRIPT_URL` and the connector URL
stays the same for everyone.
