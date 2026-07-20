import { getGraphClient } from "./graph.js";
export async function readContacts(top = 20) {
    const client = await getGraphClient();
    const result = await client
        .api("/me/contacts")
        .top(top)
        .select("id,displayName,emailAddresses,mobilePhone")
        .get();
    return result.value;
}
export async function searchContacts(query) {
    const client = await getGraphClient();
    const escaped = query.replace(/'/g, "''");
    const result = await client
        .api("/me/contacts")
        .filter(`startswith(displayName,'${escaped}')`)
        .select("id,displayName,emailAddresses,mobilePhone")
        .get();
    return result.value;
}
export async function createContact(displayName, email, mobilePhone) {
    const client = await getGraphClient();
    return await client.api("/me/contacts").post({
        displayName,
        emailAddresses: [
            {
                address: email,
                name: displayName,
            },
        ],
        mobilePhone,
    });
}
