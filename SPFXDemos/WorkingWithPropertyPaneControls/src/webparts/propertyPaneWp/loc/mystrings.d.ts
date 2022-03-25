declare interface IPropertyPaneWpWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'PropertyPaneWpWebPartStrings' {
  const strings: IPropertyPaneWpWebPartStrings;
  export = strings;
}
