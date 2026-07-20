import { getGraphClient } from "./graph.js";
export async function sendMail(to, subject, body) {
    const client = await getGraphClient();
    await client.api("/me/sendMail").post({
        message: {
            subject,
            body: {
                contentType: "Text",
                content: body,
            },
            toRecipients: [
                {
                    emailAddress: {
                        address: to,
                    },
                },
            ],
        },
        saveToSentItems: true,
    });
    return {
        success: true,
    };
}
