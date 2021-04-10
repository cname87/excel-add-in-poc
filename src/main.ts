/// <reference path="../node_modules/@types/office-js/index.d.ts" />

import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';
import { environment } from './environments/environment';

if (environment.production) {
  enableProdMode();
}
// declare const Office: any;
Office.initialize = reason => {
  // document.getElementById('sideload-msg').style.display = 'none';
  console.log('office is initialized');
  // Bootstrap the app
  platformBrowserDynamic()
      .bootstrapModule(AppModule)
      .catch(error => console.error(error));
};