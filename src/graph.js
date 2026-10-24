// Minimal delegated, read-only request. No token or profile is logged or stored.
export const GRAPH_ME_ENDPOINT = "https://graph.microsoft.com/v1.0/me?$select=id";

export async function verifyGraphMe(accessToken, fetchImplementation = fetch) {
  if (typeof accessToken !== "string" || accessToken.trim().length === 0) {
    throw new TypeError("A delegated Microsoft Graph access token is required.");
  }

  const response = await fetchImplementation(GRAPH_ME_ENDPOINT, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
      "Cache-Control": "no-store",
    },
    cache: "no-store",
    credentials: "omit",
    redirect: "error",
  });

  if (!response.ok) {
    throw new Error(`Microsoft Graph returned HTTP ${response.status}.`);
  }

  const profile = await response.json();
  if (typeof profile?.id !== "string" || profile.id.length === 0) {
    throw new Error("Microsoft Graph did not return a valid /me identifier.");
  }

  // Never return the account identifier or the access token to callers.
  return { verified: true, httpStatus: response.status };
}
