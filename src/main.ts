/// <reference path="../node_modules/@types/office-js/index.d.ts" />

import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';
import { environment } from './environments/environment';

if (environment.production) {
  enableProdMode();
}

Office.initialize = (_reason: any) => {
  //If you need to initialize something you can do so here. 
  console.log('ReleasePlanner addin - Office is initialized');
  // Bootstrap the app
  platformBrowserDynamic()
      .bootstrapModule(AppModule)
      .catch(error => console.error(error));
};