/* global Office console */

// Write text to the cursor point in the compose surface.
export async function insertText(text: string) {
  try {
    Office.context.mailbox.item?.body.setSelectedDataAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      (asyncResult: Office.AsyncResult<void>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          throw asyncResult.error.message;
        }
      }
    );
  } catch (error) {
    console.log("Error: " + error);
  }
}

// get Mail Details
export async function getMailDetails(
  callback: (data: { subject: string; sender: string; body: string; user_email: string }) => void
) {
  console.log("getMailDetails");

  Office.context.mailbox.item.subject.getAsync((subjectResult) => {
    if (subjectResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to get subject:", subjectResult.error.message);
      return;
    }

    const subject = subjectResult.value;

    Office.context.mailbox.item.from.getAsync((fromResult) => {
      if (fromResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to get sender:", fromResult.error.message);
        return;
      }

      const sender = fromResult.value.emailAddress;
      const user_email = sender;

      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
        if (bodyResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to get body:", bodyResult.error.message);
          return;
        }

        const body = bodyResult.value;

        callback({ subject, sender, body, user_email });
      });
    });
  });
}
