

import { getGraphClient } from "./graph.js";

export async function readCalendarEvents(top = 10) {
  const client = await getGraphClient();

  const result = await client
    .api("/me/events")
    .top(top)
    .orderby("start/dateTime")
    .select("id,subject,start,end,location")
    .get();

  return result.value;
}

export async function createCalendarEvent(
  subject: string,
  start: string,
  end: string,
  timeZone = "UTC",
  location?: string,
  body?: string
) {
  const client = await getGraphClient();

  return await client.api("/me/events").post({
    subject,
    start: { dateTime: start, timeZone },
    end: { dateTime: end, timeZone },
    location: location ? { displayName: location } : undefined,
    body: body
      ? {
          contentType: "Text",
          content: body,
        }
      : undefined,
  });
}

export async function deleteCalendarEvent(id: string) {
  const client = await getGraphClient();

  await client.api(`/me/events/${id}`).delete();

  return { success: true };
}