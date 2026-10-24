import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";
import { GRAPH_ME_ENDPOINT, verifyGraphMe } from "../src/graph.js";

function okResponse(body = { id: "synthetic-user-id" }) {
  return { ok: true, status: 200, json: async () => body };
}

test("uses only the Microsoft Graph /me endpoint with minimal id projection", () => {
  assert.equal(GRAPH_ME_ENDPOINT, "https://graph.microsoft.com/v1.0/me?$select=id");
});

test("one successful GET returns only success metadata, not a personal identifier", async () => {
  let invocations = 0;
  const result = await verifyGraphMe("synthetic-token", async (url, opts) => {
    invocations += 1;
    assert.equal(url, GRAPH_ME_ENDPOINT);
    assert.equal(opts.method, "GET");
    assert.equal(opts.headers.Authorization, "Bearer synthetic-token");
    assert.equal(opts.headers.Accept, "application/json");
    assert.equal(opts.cache, "no-store");
    assert.equal(opts.credentials, "omit");
    assert.equal(opts.redirect, "error");
    assert.equal(opts.body, undefined);
    return okResponse({ id: "never-display-me" });
  });
  assert.equal(invocations, 1);
  assert.deepEqual(result, { verified: true, httpStatus: 200 });
  assert.doesNotMatch(JSON.stringify(result), /never-display-me|synthetic-token/);
});

test("rejects an empty token before any network access", async () => {
  let called = false;
  await assert.rejects(() => verifyGraphMe(" ", async () => {
    called = true;
    return okResponse();
  }), /token is required/);
  assert.equal(called, false);
});

test("rejects a non-string token", async () => {
  await assert.rejects(() => verifyGraphMe({ accessToken: "bad" }), TypeError);
});

test("HTTP 401 is reported without response body or token leakage", async () => {
  const token = "very-sensitive-example-token";
  await assert.rejects(() => verifyGraphMe(token, async () => ({
    ok: false, status: 401,
    json: async () => ({ error: "sensitive-response-body" }),
  })), (err) => {
    assert.match(err.message, /HTTP 401/);
    assert.doesNotMatch(err.message, /very-sensitive|sensitive-response/);
    return true;
  });
});

test("rejects a malformed user response", async () => {
  await assert.rejects(() => verifyGraphMe("synthetic", async () => okResponse({ displayName: "do-not-retain" })), /valid \/me identifier/);
});

test("rejects an invalid id type", async () => {
  await assert.rejects(() => verifyGraphMe("synthetic", async () => okResponse({ id: 12 })), /valid \/me identifier/);
});

test("app uses delegated User.Read and never auto-executes /me on load", async () => {
  const source = await readFile(new URL("../graph-test.js", import.meta.url), "utf8");
  assert.match(source, /loginPopup\(\{\s*scopes: \["User\.Read"\]/);
  assert.match(source, /acquireTokenSilent\(\{/);
  assert.match(source, /graphButton\.addEventListener\("click"/);
  assert.doesNotMatch(source, /console\.(?:log|warn|error)/);
  assert.doesNotMatch(source, /fetch\([^\)]*(?:\/groups|\/users|\/sites)/);
});

test("redirect callback is a dedicated same-origin local module", async () => {
  const html = await readFile(new URL("../redirect.html", import.meta.url), "utf8");
  const js = await readFile(new URL("../src/redirect.js", import.meta.url), "utf8");
  assert.equal((html.match(/<script/g) || []).length, 1);
  assert.match(html, /src="\/src\/redirect\.js"/);
  assert.match(js, /broadcastResponseToMainFrame/);
  assert.doesNotMatch(html, /https?:\/\//);
});

test("no configuration secrets embedded in public source", async () => {
  const env = await readFile(new URL("../.env.example", import.meta.url), "utf8");
  assert.match(env, /VITE_ENTRA_CLIENT_ID=\n/);
  assert.match(env, /VITE_ENTRA_TENANT_ID=\n/);
  const ignored = await readFile(new URL("../.gitignore", import.meta.url), "utf8");
  assert.match(ignored, /^\.env\.\*/m);
});
