import { getGraphClient } from "./graph.js";
import { getAccessToken } from "./auth.js";
import { jwtDecode } from "jwt-decode";

export async function readInbox(top = 10) {
  console.log("READ_INBOX: starting");

  try {
    const token = await getAccessToken();
    const claims = jwtDecode<any>(token);

    console.log("===== JWT CLAIMS =====");
    console.log({
      aud: claims.aud,
      iss: claims.iss,
      tid: claims.tid,
      scp: claims.scp,
      exp: new Date(claims.exp * 1000).toISOString(),
    });
    console.log("======================");

    const url = new URL("https://graph.microsoft.com/v1.0/me/messages");
    url.searchParams.set("$top", String(top));
    url.searchParams.set("$orderby", "receivedDateTime desc");
    url.searchParams.set(
      "$select",
      "id,subject,from,receivedDateTime,isRead,bodyPreview"
    );

    const cleanToken = token.trim();

    const headers = {
      Authorization: `Bearer ${cleanToken}`,
      Accept: "application/json",
      "Content-Type": "application/json",
      "User-Agent": "OpenClaw-Outlook-Plugin",
    };

    console.log("Request URL:", url.toString());
    console.log("Authorization prefix:", headers.Authorization.substring(0, 30) + "...");
    console.log("Token length:", token.length);
    console.log("Clean token length:", cleanToken.length);
    console.log("Token equals trimmed:", token === cleanToken);

    const response = await fetch(url.toString(), {
      method: "GET",
      headers,
      redirect: "manual",
    });

    console.log("HTTP Status:", response.status);
    console.log(
      "Response Headers:",
      Object.fromEntries(response.headers.entries())
    );
    console.log("Final URL:", response.url);
    console.log("Response redirected:", response.redirected);

    const body = await response.text();

    console.log(
      "WWW-Authenticate:",
      response.headers.get("www-authenticate")
    );
    console.log("Response Body:", body);

    if (!response.ok) {
      throw new Error(`Graph returned ${response.status}: ${body}`);
    }

    return JSON.parse(body);
  } catch (err: any) {
    console.error("========== GRAPH ERROR ==========");
    console.error(err);
    console.error("name:", err?.name);
    console.error("status:", err?.statusCode);
    console.error("code:", err?.code);
    console.error("message:", err?.message);
    console.error("body:", err?.body);
    console.error("stack:", err?.stack);

    if (err?.body) {
      try {
        console.error("parsed body:", JSON.parse(err.body));
      } catch {
        // ignore JSON parse failures
      }
    }

    throw err;
  }
}

export async function readUnreadMail(top = 10) {
  const client = await getGraphClient();

  const result = await client
    .api("/me/messages")
    .filter("isRead eq false")
    .top(top)
    .orderby("receivedDateTime DESC")
    .select("id,subject,from,receivedDateTime,bodyPreview")
    .get();

  return result.value;
}

export async function searchMail(query: string) {
  const client = await getGraphClient();

  const result = await client
    .api(`/me/messages?$search="${query}"`)
    .header("ConsistencyLevel", "eventual")
    .get();

  return result.value;
}