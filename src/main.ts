/// <reference path="../node_modules/@types/office-js/index.d.ts" />

import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';
import { environment } from './environments/environment';

if (environment.production) {
  enableProdMode();
}

Office.onReady((info) => {
  console.log(info);
  if (info.host === Office.HostType.Excel) {
      // Do Excel-specific initialization (for example, make add-in task pane's
      // appearance compatible with Excel "green").
  }
  if (info.platform === Office.PlatformType.PC) {
      // Make minor layout changes in the task pane.
  }
  console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
});


Office.initialize = (reason) => {
  //If you need to initialize something you can do so here. 
  console.log('Office is initialized');
  console.log(`Reason in ${reason}`)

  if (Office.context.requirements.isSetSupported('ExcelApi', '1.12'))
{
  console.log('ExcelAp1 v1.12 supported');
}
else
{
  console.log('ExcelApi v1.12 supported');
}
  // Bootstrap the app
  platformBrowserDynamic()
      .bootstrapModule(AppModule)
      .catch(error => console.error(error));
};