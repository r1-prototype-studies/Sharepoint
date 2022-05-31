import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MicrosoftTeamsTabWebPart.module.scss';
import * as strings from 'MicrosoftTeamsTabWebPartStrings';

import * as microsoftteams from '@microsoft/teams-js';

export interface IMicrosoftTeamsTabWebPartProps {
  description: string;
}


//URL for the manifest.json file
//https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/creating-team-manifest-manually-for-webpart

export default class MicrosoftTeamsTabWebPart extends BaseClientSideWebPart <IMicrosoftTeamsTabWebPartProps> {


  private _teamsContext:microsoftTeams.Context;

  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();

    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }


  public render(): void {


    let greetingsTitle: string = '';
    let siteOrTabTitle: string = '';
  
    if (this._teamsContext) {
      
      greetingsTitle = "Welcome to Microsoft Teams!";
      siteOrTabTitle = "Team Name is : " + this._teamsContext.teamName;
    }
    else
    {
      
      greetingsTitle = "Welcome to SharePoint!";
      siteOrTabTitle = "SharePoint site: " + this.context.pageContext.web.title;
    }

    this.domElement.innerHTML = `
      <div class="${ styles.microsoftTeamsTab }">
    <div class="${ styles.container }">
      <div class="${ styles.row }">
        <div class="${ styles.column }">
          <span class="${ styles.title }">Welcome to SharePoint!</span>
  <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
    <p class="${ styles.description }">${escape(this.properties.description)}</p>

    <p class="${ styles.title }">${greetingsTitle}</p>
    <p class="${ styles.description }">${siteOrTabTitle}</p>


      <a href="https://aka.ms/spfx" class="${ styles.button }">
        <span class="${ styles.label }">Learn more</span>
          </a>
          </div>
          </div>
          </div>
          </div>`;
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
