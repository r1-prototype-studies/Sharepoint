import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./CruddemoWebPart.module.scss";
import * as strings from "CruddemoWebPartStrings";

import {
  ISPHttpClientOptions,
  SPHttpClient,
  SPHttpClientResponse,
} from "@microsoft/sp-http";
import { ISoftwareListItem } from "./ISoftwareListItem";

export interface ICruddemoWebPartProps {
  description: string;
}

export default class CruddemoWebPart extends BaseClientSideWebPart<ICruddemoWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div>
     
      <div>
        <table border='5' bgcolor='aqua'>

          <tr>
            <td>Please Enter Software ID </td>
            <td><input type='text' id='txtID' />
            <td><input type='submit' id='btnRead' value='Read Details' />
            </td>
          </tr>
          
          <tr>
            <td>Software Title</td>
            <td><input type='text' id='txtSoftwareTitle' />
          </tr>

          <tr>
            <td>Software Name</td>
            <td><input type='text' id='txtSoftwareName' />
          </tr>

          <tr>
            <td>Software Vendor</td>
            <td>
            <select id="ddlSoftwareVendor">
              <option value="Microsoft">Microsoft</option>
              <option value="Sun">Sun</option>
              <option value="Oracle">Oracle</option>
              <option value="Google">Google</option>
            </select>  
            </td>
          
          </tr>

          <tr>
            <td>Software Version</td>
            <td><input type='text' id='txtSoftwareVersion' />
          </tr>

          <tr>
            <td>Software Description</td>
            <td><textarea rows='5' cols='40' id='txtSoftwareDescription'> </textarea> </td>
          </tr>

          <tr>
            <td colspan='2' align='center'>
            <input type='submit'  value='Insert Item' id='btnSubmit' />
            <input type='submit'  value='Update' id='btnUpdate' />
            <input type='submit'  value='Delete' id='btnDelete' />      
            <input type='submit'  value='Show All Records' id='btnReadAll' />      
            </td>
          </tr>
        </table>
      </div>
      <div id="divStatus"/>
    </div>`;

    this._bindEvents();
  }

  private _bindEvents(): void {
    this.domElement
      .querySelector("#btnSubmit")
      .addEventListener("click", () => {
        this.addListItem();
      });

    this.domElement.querySelector("#btnRead").addEventListener("click", () => {
      this.readListItem();
    });

    this.domElement
      .querySelector("#btnUpdate")
      .addEventListener("click", () => {
        this.updateListItem();
      });

    this.domElement
      .querySelector("#btnDelete")
      .addEventListener("click", () => {
        this.deleteListItem();
      });

    this.domElement
      .querySelector("#btnReadAll")
      .addEventListener("click", () => {
        this.ReadAllItems();
      });
  }

  private ReadAllItems(): void {
    this._getListItems().then((listItems) => {
      let html: string =
        '<table border=1 width=100% style="border-collapse: collapse;">';
      html +=
        "<th>Title</th> <th>Vendor</th><th>Description</th><th>Name</th><th>Version</th>";

      listItems.forEach((listItem) => {
        html += `<tr>            
      <td>${listItem.Title}</td>
      <td>${listItem.SoftwareVendor}</td>
      <td>${listItem.SoftwareDescription}</td>
      <td>${listItem.SoftwareName}</td>
      <td>${listItem.SoftwareVersion}</td>      
      </tr>`;
      });
      html += "</table>";
      const listContainer: Element =
        this.domElement.querySelector("#divStatus");

      listContainer.innerHTML = html;
    });
  }

  private _getListItems(): Promise<ISoftwareListItem[]> {
    const siteUrl: string =
      this.context.pageContext.site.absoluteUrl +
      "/_api/web/lists/getbytitle('SampleList')/items";

    return this.context.spHttpClient
      .get(siteUrl, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      })
      .then((json) => {
        return json.value;
      }) as Promise<ISoftwareListItem[]>;
  }

  private deleteListItem() {
    let id: string = document.getElementById("txtID")["value"];

    const siteUrl: string =
      this.context.pageContext.site.absoluteUrl +
      "/_api/web/lists/getbytitle('SampleList')/items(" +
      id +
      ")";

    const headers: any = {
      "X-HTTP-Method": "DELETE",
      "IF-MATCH": "*",
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      headers: headers,
    };

    this.context.spHttpClient
      .post(siteUrl, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        const statusMessage: Element =
          this.domElement.querySelector("#divStatus");
        if (response.status === 204) {
          statusMessage.innerHTML = "List Item has been deleted successfully.";
        } else {
          statusMessage.innerHTML =
            "An error has occurred: " +
            response.status +
            " - " +
            response.statusText +
            " ";
          response.json().then((res) => {
            statusMessage.innerHTML += JSON.stringify(res);
          });
        }
      });
  }

  private updateListItem(): void {
    var softwaretitle = document.getElementById("txtSoftwareTitle")["value"];
    var softwarename = document.getElementById("txtSoftwareName")["value"];
    var softwareversion =
      document.getElementById("txtSoftwareVersion")["value"];
    var softwarevendor = document.getElementById("ddlSoftwareVendor")["value"];
    var softwareDescription = document.getElementById("txtSoftwareDescription")[
      "value"
    ];

    let id: string = document.getElementById("txtID")["value"];

    const siteUrl: string =
      this.context.pageContext.site.absoluteUrl +
      "/_api/web/lists/getbytitle('SampleList')/items(" +
      id +
      ")";

    alert(siteUrl);
    const itemBody: any = {
      Title: softwaretitle,
      SoftwareVendor: softwarevendor,
      SoftwareDescription: softwareDescription,
      SoftwareName: softwarename,
      SoftwareVersion: softwareversion,
    };

    const headers: any = {
      "X-HTTP_Method": "MERGE",
      "IF-MATCH": "*",
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(itemBody),
      headers: headers,
    };

    //alert(JSON.stringify(itemBody));
    this.context.spHttpClient
      .post(siteUrl, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        const statusMessage: Element =
          this.domElement.querySelector("#divStatus");
        if (response.status === 204) {
          statusMessage.innerHTML = "List Item has been updated successfully.";
        } else {
          statusMessage.innerHTML =
            "An error has occurred: " +
            response.status +
            " - " +
            response.statusText +
            " ";
          response.json().then((res) => {
            statusMessage.innerHTML += JSON.stringify(res);
          });
        }
      });
  }

  private readListItem(): void {
    let id: string = document.getElementById("txtID")["value"];
    this._getListItemByID(id)
      .then((listItem) => {
        document.getElementById("txtSoftwareTitle")["value"] = listItem.Title;
        document.getElementById("ddlSoftwareVendor")["value"] =
          listItem.SoftwareVendor;
        document.getElementById("txtSoftwareDescription")["value"] =
          listItem.SoftwareDescription;
        document.getElementById("txtSoftwareName")["value"] =
          listItem.SoftwareName;
        document.getElementById("txtSoftwareVersion")["value"] =
          listItem.SoftwareVersion;
      })
      .catch((error) => {
        const message: Element = this.domElement.querySelector(
          "#spListCreateItemUpdate"
        );
        message.innerHTML = "Read: could not fetch details.." + error.message;
      });
  }
  private _getListItemByID(id: string): Promise<ISoftwareListItem> {
    const siteUrl: string =
      this.context.pageContext.site.absoluteUrl +
      "/_api/web/lists/getbytitle('SampleList')/items?$filter=Id eq " +
      id;

    return this.context.spHttpClient
      .get(siteUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((listItems: any) => {
        const untypedItem: any = listItems.value[0];
        const listItem: ISoftwareListItem = untypedItem as ISoftwareListItem;
        return listItem;
      }) as Promise<ISoftwareListItem>;
  }

  private addListItem(): void {
    var softwaretitle = document.getElementById("txtSoftwareTitle")["value"];
    var softwarename = document.getElementById("txtSoftwareName")["value"];
    var softwareversion =
      document.getElementById("txtSoftwareVersion")["value"];
    var softwarevendor = document.getElementById("ddlSoftwareVendor")["value"];
    var softwareDescription = document.getElementById("txtSoftwareDescription")[
      "value"
    ];

    const siteUrl: string =
      this.context.pageContext.site.absoluteUrl +
      "/_api/web/lists/getbytitle('SampleList')/items";

    const itemBody: any = {
      Title: softwaretitle,
      SoftwareVendor: softwarevendor,
      SoftwareDescription: softwareDescription,
      SoftwareName: softwarename,
      SoftwareVersion: softwareversion,
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(itemBody),
    };

    //alert(JSON.stringify(itemBody));
    this.context.spHttpClient
      .post(siteUrl, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        const statusMessage: Element =
          this.domElement.querySelector("#divStatus");
        if (response.status === 201) {
          statusMessage.innerHTML = "List Item has been created successfully.";
          this.clear();
        } else {
          statusMessage.innerHTML =
            "An error has occurred: " +
            response.status +
            " - " +
            response.statusText +
            " ";
          response.json().then((res) => {
            statusMessage.innerHTML += JSON.stringify(res);
          });
        }
      });
  }

  private clear(): void {
    document.getElementById("txtSoftwareTitle")["value"] = "";
    document.getElementById("txtSoftwareName")["value"] = "Microsoft";
    document.getElementById("txtSoftwareVersion")["value"] = "";
    document.getElementById("ddlSoftwareVendor")["value"] = "";
    document.getElementById("txtSoftwareDescription")["value"] = "";
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams
      return this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentTeams
        : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost
      ? strings.AppLocalEnvironmentSharePoint
      : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    this.domElement.style.setProperty("--bodyText", semanticColors.bodyText);
    this.domElement.style.setProperty("--link", semanticColors.link);
    this.domElement.style.setProperty(
      "--linkHovered",
      semanticColors.linkHovered
    );
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
