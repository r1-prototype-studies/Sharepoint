import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
} from "@microsoft/sp-listview-extensibility";
import { Dialog } from "@microsoft/sp-dialog";

import * as strings from "CustomCommandSetCommandSetStrings";
import pnp from "sp-pnp-js";

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

    const compareThreeCommand: Command = this.tryGetCommand("COMMAND_3");
    if (compareThreeCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareThreeCommand.visible = event.selectedRows.length > 1;
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
      case "COMMAND_3":
        Dialog.prompt(`Project Status Remarks`).then((value: string) => {
          this.UpdateRemarks(event.selectedRows, value);
        });
        break;
      default:
        throw new Error("Unknown command");
    }
  }

  private UpdateRemarks(
    selectedRows: readonly import("@microsoft/sp-listview-extensibility").RowAccessor[],
    value: string
  ) {
    let batch = pnp.sp.createBatch();
    selectedRows.forEach((selectedRow) => {
      pnp.sp.web.lists
        .getByTitle("ProjectsStatus")
        .items.getById(selectedRow.getValueByName("ID"))
        .inBatch(batch)
        .update({ Remarks: value })
        .then((res) => {});
    });

    batch.execute().then((res) => {
      location.reload();
    });
  }
}
