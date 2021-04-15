import { Component } from '@angular/core';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
})
export class AppComponent {
  title = 'ex2-angular';

  onClick = (_event: MouseEvent): void => {
    this.writeText();
  };

  /* Reads data from current document selection and displays a notification */
  writeText = (): void => {
    Office.context.document.setSelectedDataAsync('Data here', (asyncResult) => {
      const error = asyncResult.error;
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(`Error: ${error}`);
        /* Show error. Upcoming displayDialog API will help here. */
      } else {
        console.log('Text written');
        /* Show success. Upcoming displayDialog API will help here. */
      }
    });
  };
}
