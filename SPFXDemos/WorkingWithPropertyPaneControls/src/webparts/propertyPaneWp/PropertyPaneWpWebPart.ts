import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PropertyPaneWpWebPart.module.scss';
import * as strings from 'PropertyPaneWpWebPartStrings';

export interface IPropertyPaneWpWebPartProps {
  description: string;

  productName: string;
  productDescription: string;
  productCost: number;
  quantity: number;
  billAmount: number;
  discount: number;
  netBillAmount: number;

}

export default class PropertyPaneWpWebPart extends BaseClientSideWebPart<IPropertyPaneWpWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    this.properties.productName = "Mouse";
    this.properties.productDescription = "mouse product description";
    this.properties.productCost = 50;
    this.properties.quantity = 12;
    return super.onInit();
  }

  protected get disableReactivePropertyChanges(): boolean {
      return true;
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.propertyPaneWp} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
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
            <td>${this.properties.billAmount = this.properties.productCost * this.properties.quantity}</td>
          </tr>
          <tr>
            <td>Discount</td>
            <td>${this.properties.discount = this.properties.billAmount * 0.1}</td>
          </tr>
          <tr>
            <td>Net Nill Amount</td>
            <td>${this.properties.netBillAmount = this.properties.billAmount - this.properties.discount}</td>
          </tr>
        </table>
      </div>
    </section>`;
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
    return{
      pages: [
        {
          groups: [
            {
              groupName: "Product details",
              groupFields: [
                PropertyPaneTextField('productName',{
                  label: "Product Name",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter product name","description": "Name property field"
                }),
                PropertyPaneTextField('productDescription',{
                  label: "Product Description",
                  multiline: true,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter product description","description": "Name property field"
                }),
                PropertyPaneTextField('productCost',{
                  label: "Product Cost",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter product cost","description": "Name property field"
                }),
                PropertyPaneTextField('quantity',{
                  label: "Quantity",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter quantity","description": "Name property field"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
