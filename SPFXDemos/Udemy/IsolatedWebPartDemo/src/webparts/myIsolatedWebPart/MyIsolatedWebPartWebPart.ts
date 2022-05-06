import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MyIsolatedWebPartWebPart.module.scss';
import * as strings from 'MyIsolatedWebPartWebPartStrings';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IMyIsolatedWebPartWebPartProps {
  description: string;
}

export default class MyIsolatedWebPartWebPart extends BaseClientSideWebPart <IMyIsolatedWebPartWebPartProps> {

  public render(): void {

    this.context.msGraphClientFactory
    .getClient()
    .then((graphclient: MSGraphClient): void => {        
      graphclient
        .api('/me')
        .get((error, user:MicrosoftGraph.User, rawResponse?: any) => {

    this.domElement.innerHTML = `
          <div>
          <p class="${ styles.description }">Display Name: ${user.displayName}</p>
          <p class="${ styles.description }">Given Name: ${user.givenName}</p>
          <p class="${ styles.description }">Surname: ${user.surname}</p>
          <p class="${ styles.description }">Email ID: ${user.mail}</p>
          <p class="${ styles.description }">Mobile Phone: ${user.mobilePhone}</p>   
          </div>`;
        });
      });
     
  }

  protected get dataVersion(): Version {
  return Version.parse('1.0');
}

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  return {
    pages: [
      {
        header: {
          description: strings.PropertyPaneDescription
        },
        groups: [
          {
            groupName: strings.BasicGroupName,
            groupFields: [
              PropertyPaneTextField('description', {
                label: strings.DescriptionFieldLabel
              })
            ]
          }
        ]
      }
    ]
  };
}
}
