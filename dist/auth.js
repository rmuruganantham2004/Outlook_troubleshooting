const scopes = [
    "User.Read",
    "Mail.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.Read",
    "Calendars.ReadWrite",
    "Contacts.Read",
    "Contacts.ReadWrite",
    "offline_access",
];
import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from "url";
import { PublicClientApplication, LogLevel, } from "@azure/msal-node";
import { writeFileSync } from "fs";
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
dotenv.config({ path: path.resolve(__dirname, "../.env") });
const tenantId = process.env.AZURE_TENANT_ID;
const clientId = process.env.AZURE_CLIENT_ID;
if (!tenantId || !clientId) {
    throw new Error("Missing AZURE_TENANT_ID or AZURE_CLIENT_ID in your .env file.");
}
const msal = new PublicClientApplication({
    auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    system: {
        loggerOptions: {
            loggerCallback: () => { },
            piiLoggingEnabled: false,
            logLevel: LogLevel.Warning,
        },
    },
});
let cachedToken = null;
export async function getAccessToken() {
    if (cachedToken) {
        return cachedToken;
    }
    const accounts = await msal.getTokenCache().getAllAccounts();
    if (accounts.length > 0) {
        try {
            console.log("AUTH: attempting silent authentication");
            const silent = await msal.acquireTokenSilent({
                account: accounts[0],
                scopes,
            });
            if (silent?.accessToken) {
                console.log("AUTH: silent login succeeded");
                cachedToken = silent.accessToken;
                return cachedToken;
            }
        }
        catch {
            console.log("AUTH: silent login failed");
        }
    }
    console.log("AUTH: starting device code flow");
    const result = await msal.acquireTokenByDeviceCode({
        scopes,
        deviceCodeCallback: (response) => {
            console.log("AUTH:", response.message);
        },
    });
    if (!result?.accessToken) {
        throw new Error("Failed to acquire token.");
    }
    cachedToken = result.accessToken;
    console.log("AUTH: caching access token in memory");
    writeFileSync("token.txt", result.accessToken);
    if (result.account) {
        writeFileSync("account.json", JSON.stringify(result.account, null, 2));
    }
    console.log("AUTH: device authentication succeeded");
    return cachedToken;
}
