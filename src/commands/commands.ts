/* global Office console */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  // let dialog: Office.Dialog;

  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/answer_form.html",
    { height: 50, width: 50, displayInIframe: false },
    (asyncResult: Office.AsyncResult<Office.Dialog>) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(asyncResult.error.message);
      } else {
        // dialog =
        asyncResult.value;
        // dialog.addEventHandler(
        //   Office.EventType.DialogMessageReceived,
        //   (arg: Office.DialogMessageReceivedEventArgs) => messageHandler(arg) // Тип события
        // );
      }
    }
  );

  event.completed();
}

Office.actions.associate("action", action);
