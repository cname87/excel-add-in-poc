/* Declare all functions in main.ts */
interface Window {
  writeNote: (event: Office.AddinCommands.Event) => void;
  getAndWriteAddress: (event: Office.AddinCommands.Event) => void;
  incrementPane: (event: Office.AddinCommands.Event) => void;
  formatRange: (event: Office.AddinCommands.Event) => void;
}
