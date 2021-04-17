import { Component } from '@angular/core';
import { getAddress, writeTextToSelected } from '../functions';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
})
export class AppComponent {
  title = 'Taskpane';

  writeTextButton = (_event: MouseEvent): void => {
    writeTextToSelected('Note written from Taskpane');
  };

  getAndWriteAddressButton = async (_event: MouseEvent): Promise<void> => {
    const address = await getAddress();
    if (address) {
      writeTextToSelected(`The selected range is: ${address}`);
    }
  };
}
