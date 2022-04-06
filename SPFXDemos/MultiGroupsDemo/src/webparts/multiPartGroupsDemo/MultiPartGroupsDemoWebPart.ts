import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./MultiPartGroupsDemoWebPart.module.scss";
import * as strings from "MultiPartGroupsDemoWebPartStrings";

export interface IMultiPartGroupsDemoWebPartProps {
  description: string;
  productName: string;
  isCertified: boolean;
}

export default class MultiPartGroupsDemoWebPart extends BaseClientSideWebPart<IMultiPartGroupsDemoWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.multiPartGroupsDemo} ${
      !!this.context.sdks.microsoftTeams ? styles.teams : ""
    }">
      <div class="${styles.welcome}">
        <img alt="" src="${
          this._isDarkTheme
            ? require("./assets/welcome-dark.png")
            : require("./assets/welcome-light.png")
        }" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(
          this.context.pageContext.user.displayName
        )}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(
          this.properties.description
        )}</strong></div>
      </div>
      <div>Product Name: <strong>${escape(
        this.properties.productName
      )}</strong></div>
    </div>
    <div>Is Certified: <strong>${this.properties.isCertified}</strong></div>
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: "First Group",
              groupFields: [
                PropertyPaneTextField("productName", {
                  label: "Product Name",
                }),
              ],
            },
            {
              groupName: "Second Group",
              groupFields: [
                PropertyPaneToggle("isCertified", {
                  label: "Is Certified?",
                }),
              ],
            },
          ],
          displayGroupsAsAccordion: true,
        },
      ],
    };
  }
}
