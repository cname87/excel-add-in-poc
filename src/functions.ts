/**
 * This contains all functions invoked by a UI-less command.
 */

/**
 * Reads data from current document selection and displays a notification
 *
 * @param event The `Event` object is passed as a parameter to add-in functions invoked by UI-less command buttons. The object allows the add-in to identify which button was clicked and to signal the host that it has completed its processing.
 */
const writeText = (event: Office.AddinCommands.Event): void => {
  console.log('writeText running');
  Office.context.document.setSelectedDataAsync(
    'ExecuteFunction Works with Prod Office.js Button ID=' + event.source.id,
    (asyncResult) => {
      const error = asyncResult.error;
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(`Error: ${error}`);
        /* Show error. Upcoming displayDialog API will help here. */
      } else {
        console.log('Text written');
        /* Show success. Upcoming displayDialog API will help here. */
      }
    },
  );
  /* Required: Call event.completed to let the platform know you are done processing */
  event.completed();
};

export { writeText };
