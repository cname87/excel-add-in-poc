/* eslint-disable @typescript-eslint/no-non-null-assertion */
/**
 * This contains all functions invoked by a UI-less command.
 */

const _logError = (error: any): void => {
  console.log(`Error: ${JSON.stringify(error)}`);
  if (error instanceof OfficeExtension.Error) {
    console.log(`Debug info: ${JSON.stringify(error.debugInfo)}`);
  }
};

/**
 *  Writes supplied text to the current selection
 *  @param text Text to be written
 */
const writeTextToSelected = (text = ''): void => {
  Office.context.document.setSelectedDataAsync(text, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      _logError(asyncResult.error);
      /* Show error. Upcoming displayDialog API will help here. */
    } else {
      console.log('Text written');
      /* Show success. Upcoming displayDialog API will help here. */
    }
  });
};

/**
 *  Gets the address of the current selection and writes it to the current selection
 */
const getAddress = async (): Promise<string | void> => {
  const address = await Excel.run(async (context) => {
    const selectedRange = context.workbook.getSelectedRange();
    selectedRange.load('address');
    await context.sync();
    console.log('The selected range is: ' + selectedRange.address);
    return selectedRange.address;
  }).catch(_logError);
  return address;
};

/* Writes text to a 'test' tag in the pane */
const _writeToPane = (text = '') => {
  document.getElementById('test')!.textContent = text;
  console.log(
    `Pane text is now: ${document.getElementById('test')!.textContent!}`,
  );
};

/**
 * Reads address of the current document selection and writes it to the current document selection.
 * @param event Office.AddinCommands.Event
 */
const getAndWriteAddress = async (
  event: Office.AddinCommands.Event,
): Promise<void> => {
  console.log('getAndWriteAddress running');
  const address = await getAddress();
  writeTextToSelected(`The selected range is: ${address}`);
  event.completed();
};

/**
 * Increments a tag on the Taskpane
 * @param event Office.AddinCommands.Event
 */
let _count = 0;
const incrementPane = (event: Office.AddinCommands.Event): void => {
  console.log('incrementPane running');
  _writeToPane(`Count: ${_count++}`);
  event.completed();
};

/**
 * Reads data from current document selection and writes it to the current selection.
 *
 * @param event The `Event` object is passed as a parameter to add-in functions invoked by UI-less command buttons. The object allows the add-in to identify which button was clicked and to signal the host that it has completed its processing.
 */

const writeNote = (event: Office.AddinCommands.Event): void => {
  console.log('writeNote running');
  writeTextToSelected(
    `ExecuteFunction Works with Office.js Button ID = ${event.source.id}`,
  );
  /* Required: Call event.completed to let the platform know you are done processing */
  event.completed();
};

const formatRange = (event: Office.AddinCommands.Event): void => {
  Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItemOrNullObject('Sample');

    await context.sync();
    if (sheet.isNullObject) {
      sheet = context.workbook.worksheets.add('Sample');
    }
    /* Set `sheet` to be the second worksheet in the workbook */
    sheet.position = 1;
    /* Format range */
    const range = sheet.getRange('B2:E2');
    range.format.autofitColumns();
    range.set({
      format: {
        fill: {
          color: '#4472C4',
        },
        font: {
          name: 'Verdana',
          color: 'white',
        },
      },
    });
  }).catch(_logError);
  event.completed();
};

export {
  formatRange,
  getAddress,
  getAndWriteAddress,
  incrementPane,
  writeNote,
  writeTextToSelected,
};
