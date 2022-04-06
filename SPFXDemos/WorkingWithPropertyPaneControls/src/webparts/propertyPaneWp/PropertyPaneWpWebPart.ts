import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./PropertyPaneWpWebPart.module.scss";
import * as strings from "PropertyPaneWpWebPartStrings";

export interface IPropertyPaneWpWebPartProps {
  description: string;

  productName: string;
  productDescription: string;
  productCost: number;
  quantity: number;
  billAmount: number;
  discount: number;
  netBillAmount: number;

  currentTime: Date;
  isCertified: boolean;
  rating: number;
  processorType: string;
  invoiceFileType: string;
}

export default class PropertyPaneWpWebPart extends BaseClientSideWebPart<IPropertyPaneWpWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    this.properties.productName = "Mouse";
    this.properties.productDescription = "mouse product description";
    this.properties.productCost = 50;
    this.properties.quantity = 12;
    this.properties.isCertified = false;
    this.properties.rating = 1;
    return super.onInit();
  }

  protected get disableReactivePropertyChanges(): boolean {
    return false;
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.propertyPaneWp} ${
      !!this.context.sdks.microsoftTeams ? styles.teams : ""
    }">
      <div class="${styles.propertyPaneWp}">
        <table>
          <tr>
            <td>Product Name</td>
            <td>${this.properties.productName}</td>
          </tr>
          <tr>
            <td>Description</td>
            <td>${this.properties.productDescription}</td>
          </tr>
          <tr>
            <td>Product Cost</td>
            <td>${this.properties.productCost}</td>
          </tr>
          <tr>
            <td>Product Quantity</td>
            <td>${this.properties.quantity}</td>
          </tr>
          <tr>
            <td>Bill Amount</td>
            <td>${(this.properties.billAmount =
              this.properties.productCost * this.properties.quantity)}</td>
          </tr>
          <tr>
            <td>Discount</td>
            <td>${(this.properties.discount =
              this.properties.billAmount * 0.1)}</td>
          </tr>
          <tr>
            <td>Net Nill Amount</td>
            <td>${(this.properties.netBillAmount =
              this.properties.billAmount - this.properties.discount)}</td>
          </tr>
          <tr>
            <td>Is Certified?</td>
            <td>${this.properties.isCertified}</td>
          </tr>
          <tr>
            <td>Rating</td>
            <td>${this.properties.rating}</td>
          </tr>
          <tr>
            <td>Processor Type</td>
            <td>${this.properties.processorType}</td>
          </tr>
          <tr>
            <td>Invoice file type</td>
            <td>${this.properties.invoiceFileType}</td>
          </tr>
        </table>
      </div>
    </section>`;
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

  // protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  //   return {
  //     pages: [
  //       {
  //         header: {
  //           description: "Propery Values"
  //         },
  //         groups: [
  //           {
  //             groupName: "Grouping",
  //             groupFields: [
  //               PropertyPaneTextField('description', {
  //                 label: strings.DescriptionFieldLabel
  //               })
  //             ]
  //           },
  //           {
  //             groupName: "Section",
  //             groupFields: [
  //               PropertyPaneTextField('description', {
  //                 label: "Test"
  //               })
  //             ]
  //           }
  //         ]
  //       }
  //     ]
  //   };
  //}

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: "Product details",
              groupFields: [
                PropertyPaneTextField("productName", {
                  label: "Product Name",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter product name",
                  description: "Name property field",
                }),
                PropertyPaneTextField("productDescription", {
                  label: "Product Description",
                  multiline: true,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter product description",
                  description: "Name property field",
                }),
                PropertyPaneTextField("productCost", {
                  label: "Product Cost",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter product cost",
                  description: "Name property field",
                }),
                PropertyPaneTextField("quantity", {
                  label: "Quantity",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter quantity",
                  description: "Name property field",
                }),
                PropertyPaneToggle("isCertified", {
                  label: "Is it certified?",
                  key: "isCertified",
                  onText: "ISI certified!",
                  offText: "Not an ISI certified Product",
                }),
                PropertyPaneSlider("rating", {
                  label: "Select your rating",
                  min: 0.5,
                  max: 10,
                  step: 0.5,
                  showValue: true,
                  value: 1,
                }),
                PropertyPaneChoiceGroup("processorType", {
                  label: "Choices",
                  options: [
                    { key: "I5", text: "Intel I5" },
                    { key: "I7", text: "Intel I7", checked: true },
                    { key: "I9", text: "Intel I9" },
                  ],
                }),
                PropertyPaneChoiceGroup("invoiceFileType", {
                  label: "Select Invoice File Type",
                  options: [
                    {
                      key: "MSWord",
                      text: "MS Word",
                      imageSrc:
                        "https://upload.wikimedia.org/wikipedia/commons/thumb/f/fd/Microsoft_Office_Word_%282019%E2%80%93present%29.svg/1200px-Microsoft_Office_Word_%282019%E2%80%93present%29.svg.png",
                      selectedImageSrc:
                        "https://upload.wikimedia.org/wikipedia/commons/thumb/f/fd/Microsoft_Office_Word_%282019%E2%80%93present%29.svg/1200px-Microsoft_Office_Word_%282019%E2%80%93present%29.svg.png",
                      imageSize: { width: 32, height: 32 },
                    },
                    {
                      key: "MSExcel",
                      text: "MS Excel",
                      checked: true,
                      imageSrc:
                        "https://upload.wikimedia.org/wikipedia/commons/thumb/3/34/Microsoft_Office_Excel_%282019%E2%80%93present%29.svg/640px-Microsoft_Office_Excel_%282019%E2%80%93present%29.svg.png",
                      selectedImageSrc:
                        "https://upload.wikimedia.org/wikipedia/commons/thumb/3/34/Microsoft_Office_Excel_%282019%E2%80%93present%29.svg/640px-Microsoft_Office_Excel_%282019%E2%80%93present%29.svg.png",
                      imageSize: { width: 32, height: 32 },
                    },
                    {
                      key: "MSPowerPoint",
                      text: "MS PowerPoint",
                      imageSrc:
                        "https://e7.pngegg.com/pngimages/742/145/png-clipart-powerpoint-logo-microsoft-powerpoint-computer-icons-ppt-presentation-microsoft-powerpoint-network-icon-angle-text.png",
                      selectedImageSrc:
                        "https://e7.pngegg.com/pngimages/742/145/png-clipart-powerpoint-logo-microsoft-powerpoint-computer-icons-ppt-presentation-microsoft-powerpoint-network-icon-angle-text.png",
                      imageSize: { width: 32, height: 32 },
                    },
                  ],
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
