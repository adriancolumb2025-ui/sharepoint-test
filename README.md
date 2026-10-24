# SharePoint + Microsoft Graph Test

A small development experiment to verify **real, delegated, read-only Microsoft Graph connectivity** for a Microsoft Entra tenant. The previous `graph-test.js` used simulated responses; this branch replaces that simulation with a real opt-in `GET https://graph.microsoft.com/v1.0/me?$select=id` request.

**Scope:** read-only `/me`, delegated `User.Read`, no writes to Graph or SharePoint. This is **not** an E5 sandbox, tenant-provisioning script, sandbox renewal mechanism, or Microsoft Developer Program eligibility test. Existing Microsoft Graph experiments remain documented in `graph-notes.md`.

## Requirements

- Node.js 22.12 or later, npm, and a modern browser.
- Access to a Microsoft Entra tenant in which an authorised app registration exists.
- An **SPA (single-page application)** redirect URI registered on that application: exactly `http://localhost:5173/redirect.html`.
- Delegated Microsoft Graph `User.Read` permission (consent according to that tenant's policies). **Do not use application permissions** for `/me`.

Do not create a new app registration or request privileged tenant-wide permissions just to use this draft demo without reviewing the existing registration first. No client secret is required or permitted in a browser app.

## Local setup (user-triggered)

1. Check an existing Entra app registration under **Authentication**. Its SPA redirect URI must match `http://localhost:5173/redirect.html` exactly. Check **API permissions** for delegated `User.Read`. The tenant's administrator should determine whether any additional consent is necessary; this demo does not perform admin consent.
2. Copy `.env.example` to `.env.local` and fill in the existing application (client) ID and tenant (directory) ID. These two IDs are not secrets, but `.env.local` is deliberately excluded from Git. **Never store a client secret, password, refresh token or access token in a `VITE_` environment variable**: Vite exposes it in browser-delivered code.
3. Run `npm install`, then `npm run dev`.
4. Open `http://localhost:5173/`, select **1. Sign in to Microsoft Entra**, and deliberately select the development account.
5. Select **2. Test Graph (read-only)**. This makes exactly one Graph GET request with minimal projection. The interface reports pass/fail status without printing the returned identifier or access token.
6. Close the tab after testing. Tokens are held by MSAL in browser session storage, not in this application's own files or logs.

### Failure diagnosis

- `Missing SPA redirect URI` / popup does not finish: check that the Entra app is registered as an SPA and the URI is **exactly** `http://localhost:5173/redirect.html` (including path and port). MSAL v5 requires a dedicated, same-origin redirect bridge.
- `401`: sign in again, check delegated scope and consent. Never paste tokens or full authentication errors into issues or pull requests.
- `403`: the signed-in account or app may lack an applicable permission, or policy may restrict access. Do not increase permissions automatically.
- Developer Program dashboard shows no sandbox: that is a **separate entitlement issue**. Successful Graph `/me` does not demonstrate that an E5 subscription is active or that activity will reinstate one.

## Checks

- `npm run test` - offline, synthetic tests of Graph request method, endpoint, credential handling, error sanitisation, manual invocation and redirect-bridge source.
- `npm run build` - bundles the app **and** the dedicated MSAL v5 redirect bridge, using Vite's multi-page configuration.
- GitHub Actions runs tests and build on the proposed branch/PR without Microsoft credentials, secrets, tenant access or Microsoft Graph network calls.

The source uses Microsoft-supported `@azure/msal-browser` v5 with authorisation-code/PKCE. It never implements token handling from scratch. The redirect bridge is bundled locally, not served from a third-party CDN. All external dependencies are explicitly versioned in `package.json`; a lockfile can be generated and reviewed in a future dependency qualification step before promoting this draft.

## Security and boundaries

- Repository is public: do not commit `.env.local`, tokens, screenshots of identity responses, credentials or tenant-private data.
- No GitHub Actions secrets, Microsoft API calls in CI, Microsoft 365 subscription operations or SharePoint writes.
- Existing `main` remains unchanged until a separately authorised review/merge.
- This proof-of-concept is not a production identity integration or a substitute for formal app/tenant security review.

## Official references

- [MSAL Browser overview](https://learn.microsoft.com/en-us/entra/msal/javascript/browser/about-msal-browser)
- [MSAL v5 redirect bridge for Vite](https://learn.microsoft.com/en-us/entra/msal/javascript/browser/redirect-bridge)
- [Microsoft Graph GET /me permissions](https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0)
