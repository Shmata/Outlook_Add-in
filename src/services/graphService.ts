import { getAccessToken, getAccessTokenSilent } from "../auth/msal";

export interface GraphMessage {
  id: string;
  subject?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  bodyPreview?: string;
  receivedDateTime?: string;
}

export async function getMyMessages(): Promise<GraphMessage[]> {
  let token: string;
  try {
    token = await getAccessTokenSilent();
  } catch (e) {
    // silent failed, fall back to interactive
    token = await getAccessToken();
  }

  const res = await fetch("https://graph.microsoft.com/v1.0/me/messages", {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const text = await res.text();
    if (res.status === 403) {
      let body = text;
      try {
        const parsed = JSON.parse(text);
        body = parsed.error?.message || text;
      } catch (err) {
        /* ignore parse errors */
      }
      throw new Error(`Graph API Access Denied (403): ${body}. Ensure the application and user have Mail.Read permission and have consented.`);
    }
    throw new Error(`Graph API error ${res.status}: ${text}`);
  }
  const data = await res.json();
  return data.value as GraphMessage[];
}

export async function getTheCurrentMessage(encodedId: string, token?: string): Promise<GraphMessage> {

  if(token === undefined || token === null || token === "") {
    token = await getAccessToken();
  }
  console.log(encodedId);
  const res = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${encodedId}`, {
        headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) {
        const text = await res.text();
        throw new Error(`Graph API error ${res.status}: ${text}`);
    }
    const data = await res.json();
    return data as GraphMessage;
}

// Fetch a message using an Outlook REST callback token (returned by getCallbackTokenAsync with isRest: true)
export async function getMessageWithRestToken(restId: string, token: string): Promise<GraphMessage> {
  // Outlook REST endpoint expects a REST id (convertToRestId with v2.0)
  const url = `https://outlook.office.com/api/v2.0/me/messages/${encodeURIComponent(restId)}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Outlook REST API error ${res.status}: ${text}`);
  }
  const data = await res.json();
  return data as GraphMessage;
}
