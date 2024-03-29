import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';
import { environment } from './environments/environment';

import {
  formatRange,
  getAndWriteAddress,
  incrementPane,
  writeNote,
} from './functions';

if (environment.production) {
  enableProdMode();
}

Office.initialize = (reason) => {
  /* If you need to initialize something you can do so here */
  console.log('Office is initialized');

  if (reason === Office.InitializationReason.Inserted) {
    console.log(`The add-in was just inserted into the document`);
  }
  if (reason === Office.InitializationReason.DocumentOpened) {
    console.log(`The add-in is already part of the document that was opened`);
  }

  if (Office.context.requirements.isSetSupported('ExcelApi', '1.12')) {
    console.log('ExcelAp1 v1.12 is supported');
  } else {
    console.log('ExcelApi v1.12 is not supported');
  }

  /* Bootstrap the app */
  platformBrowserDynamic()
    .bootstrapModule(AppModule)
    .catch((error) => console.error(error));
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    /* Do Excel-specific initialization (for example, make add-in task pane's appearance compatible with Excel "green") */
  }
  if (info.platform === Office.PlatformType.PC) {
    /* Make platform-specific changes e.g. minor layout changes in the task pane */
  }
  console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
});

/* Functions available to a UI-less command */
window.writeNote = writeNote;
window.getAndWriteAddress = getAndWriteAddress;
window.incrementPane = incrementPane;
window.formatRange = formatRange;
