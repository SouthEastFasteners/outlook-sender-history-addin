import { getGraphToken } from "./auth";

export interface EmailSummary {
  id: string;
  subject: string;
  receivedDateTime: string;
  bodyPreview: string;
  hasAttachments: boolean;
  isRead: boolean;
}

/**
 * Fetches the most recent messages from a specific sender address,
 * newest first, excluding the currently open message.
 */
export async function getRecentEmailsFromSender(
  senderEmail: string,
  excludeMessageId: string,
  count: number = 6
): Promise<EmailSummary[]> {
  const token = await getGraphToken();

  const filter =
    `receivedDateTime ge 1900-01-01T00:00:00Z and ` +
    `from/emailAddress/address eq '${senderEmail.replace(/'/g, "''")}'`;
  const select = "id,subject,receivedDateTime,bodyPreview,hasAttachments,isRead";
  const url =
    `https://graph.microsoft.com/v1.0/me/messages` +
    `?$filter=${encodeURIComponent(filter)}` +
    `&$orderby=receivedDateTime desc` +
    `&$top=${count + 1}` + // fetch one extra in case the open message is in the results
    `&$select=${select}`;

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!response.ok) {
    const errorBody = await response.text();
    throw new Error(`Graph request failed (${response.status}): ${errorBody}`);
  }

  const data = await response.json();
  const messages: EmailSummary[] = data.value ?? [];

  return messages.filter((m) => m.id !== excludeMessageId).slice(0, count);
}
