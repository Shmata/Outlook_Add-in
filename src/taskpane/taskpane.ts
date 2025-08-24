/* global Office console */

export async function insertText(text: string) {
  // Write text to the cursor point in the compose surface.
  // TEMPORARY: force a breakpoint when a debugger is attached. Remove after testing.
  // eslint-disable-next-line no-debugger
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
