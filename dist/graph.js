import "isomorphic-fetch";
import { Client } from "@microsoft/microsoft-graph-client";
import { getAccessToken } from "./auth.js";
export async function getGraphClient() {
    console.log("GRAPH: creating client");
    const token = await getAccessToken();
    console.log("GRAPH: access token acquired");
    return Client.init({
        authProvider: (done) => {
            console.log("GRAPH: authProvider invoked");
            done(null, token);
        },
    });
}
