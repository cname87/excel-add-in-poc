/* eslint-disable no-var */
import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';
import { environment } from './environments/environment';

if (environment.production) {
  enableProdMode();
}

Office.initialize = (reason) => {
  /* If you need to initialize something you can do so here */
  console.log('Office is initialized');
  console.log(`Reason is ${reason.toString()}`);

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

/* Reads data from current document selection and displays a notification */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
window.writeText = (event: any): void => {
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
