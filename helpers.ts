// helper functions & parameters

/* global console fetch */

const Host = "https://fg.server.lavel.io";

export interface DialogInfo {
  show: boolean;
  text: string;
}

export interface FieldInfo {
  current: string;
  state: "none" | "error" | "warning" | "success" | undefined;
}

// send Request to API
export async function sendRequest(
  endpoint: string,
  apiKey: string,
  subject: string,
  sender: string,
  body: string,
  user_input: string,
  user_email: string
): Promise<any> {
  var url = Host + endpoint;
  console.log(` sendRequest to: ${url}`);

  const headers = {
    "Content-Type": "application/json",
    authorization: apiKey,
  };

  const data = {
    messages: [
      {
        subject: subject,
        body: body,
        sender: sender,
      },
    ],
    user_input: user_input,
    target_language: "german",
    user_email: user_email,
  };

  console.log("data = " + JSON.stringify(data, null, 2));

  var options = {
    method: "POST",
    headers: headers,
    body: JSON.stringify(data),
    muteHttpExceptions: true,
  };

  // HTTP request
  var response = await fetch(url, options);

  if (![200, 400, 401, 404, 502].includes(response.status)) {
    console.log(`Wrong response code: ${response.status}`);
    return { detail: `Wrong response code: ${response.status}` };
  }

  try {
    let data: any = await response.json();
    console.log("data: " + JSON.stringify(data, null, 2));
    return data;
  } catch (error) {
    console.log(`Invalid JSON: ${error}`);
    return { detail: `Invalid JSON: ${error}` };
  }
}
