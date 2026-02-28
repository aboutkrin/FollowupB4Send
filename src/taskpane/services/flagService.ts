/**
 * Flag service: saves the draft, obtains a REST token, and PATCHes the
 * follow-up flag onto the message before sending.
 */

interface FlagDates {
  startDate: string; // ISO-8601 date-time
  dueDate: string;
}

/** Save the current compose item and return the draft item ID. */
function saveDraft(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item!.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error?.message ?? "saveAsync failed"));
      }
    });
  });
}

/** Get a REST callback token. */
function getRestToken(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync(
      { isRest: true },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(result.error?.message ?? "getCallbackTokenAsync failed"));
        }
      }
    );
  });
}

/** Convert an EWS item ID to a REST-compatible ID. */
function toRestId(ewsId: string): string {
  return Office.context.mailbox.convertToRestId(
    ewsId,
    Office.MailboxEnums.RestVersion.v2_0
  );
}

/** Get the REST API base URL for the current mailbox. */
function getRestUrl(): string {
  return Office.context.mailbox.restUrl;
}

/**
 * PATCH the follow-up flag on the draft via Outlook REST v2.0.
 */
async function patchFlagRest(
  restId: string,
  token: string,
  dates: FlagDates
): Promise<void> {
  const url = `${getRestUrl()}/v2.0/me/messages/${encodeURIComponent(restId)}`;
  const body = {
    Flag: {
      FlagStatus: "Flagged",
      StartDateTime: {
        DateTime: dates.startDate,
        TimeZone: "UTC",
      },
      DueDateTime: {
        DateTime: dates.dueDate,
        TimeZone: "UTC",
      },
    },
  };

  const response = await fetch(url, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`REST PATCH failed (${response.status}): ${text}`);
  }
}

/**
 * Fallback: set the flag via EWS UpdateItem when REST is unavailable.
 */
function patchFlagEws(ewsId: string, dates: FlagDates): Promise<void> {
  return new Promise((resolve, reject) => {
    const soapRequest = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013"/>
  </soap:Header>
  <soap:Body>
    <m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite">
      <m:ItemChanges>
        <t:ItemChange>
          <t:ItemId Id="${ewsId}"/>
          <t:Updates>
            <t:SetItemField>
              <t:FieldURI FieldURI="item:Flag"/>
              <t:Message>
                <t:Flag>
                  <t:FlagStatus>Flagged</t:FlagStatus>
                  <t:StartDate>${dates.startDate}</t:StartDate>
                  <t:DueDate>${dates.dueDate}</t:DueDate>
                </t:Flag>
              </t:Message>
            </t:SetItemField>
          </t:Updates>
        </t:ItemChange>
      </m:ItemChanges>
    </m:UpdateItem>
  </soap:Body>
</soap:Envelope>`;

    Office.context.mailbox.makeEwsRequestAsync(soapRequest, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error?.message ?? "EWS UpdateItem failed"));
      }
    });
  });
}

/**
 * High-level: save draft, set the follow-up flag, then send.
 * Tries REST first, falls back to EWS.
 */
export async function setFlagAndSend(dates: FlagDates): Promise<void> {
  // 1. Save the draft to get an item ID
  const ewsId = await saveDraft();

  // 2. Try REST, fall back to EWS
  try {
    const token = await getRestToken();
    const restId = toRestId(ewsId);
    await patchFlagRest(restId, token, dates);
  } catch {
    await patchFlagEws(ewsId, dates);
  }

  // 3. Send the message
  await sendMessage();
}

/** Send the current compose item. Falls back to user instruction if API unavailable. */
function sendMessage(): Promise<void> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item!;
    if (typeof (item as any).sendAsync === "function") {
      (item as any).sendAsync((result: Office.AsyncResult<void>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message ?? "sendAsync failed"));
        }
      });
    } else {
      // sendAsync not available (requires Mailbox 1.15).
      // The flag is already set on the draft. Reject so the UI can tell
      // the user to click Send manually.
      reject(new Error("SEND_UNAVAILABLE"));
    }
  });
}

/** Send without setting a flag. */
export async function sendWithoutFlag(): Promise<void> {
  await sendMessage();
}
