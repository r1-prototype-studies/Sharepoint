import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
} from "@microsoft/sp-listview-extensibility";
import { Dialog } from "@microsoft/sp-dialog";

import * as strings from "CustomCommandSetCommandSetStrings";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICustomCommandSetCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = "CustomCommandSetCommandSet";

export default class CustomCommandSetCommandSet extends BaseListViewCommandSet<ICustomCommandSetCommandSetProperties> {
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized CustomCommandSetCommandSet");
    return Promise.resolve();
  }

  public onListViewUpdated(
    event: IListViewCommandSetListViewUpdatedParameters
  ): void {
    const compareOneCommand: Command = this.tryGetCommand("COMMAND_1");
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case "COMMAND_1":
        //Dialog.alert(`${this.properties.sampleTextOne}`);

        let title: string = event.selectedRows[0].getValueByName("Title");
        let status: string = event.selectedRows[0].getValueByName("Status");
        Dialog.alert(
          `Project Name: ${title} - Current status: ${status}% done`
        );
        break;
      case "COMMAND_2":
        Dialog.alert(`${this.properties.sampleTextTwo}`);
        break;
      default:
        throw new Error("Unknown command");
    }
  }
}
