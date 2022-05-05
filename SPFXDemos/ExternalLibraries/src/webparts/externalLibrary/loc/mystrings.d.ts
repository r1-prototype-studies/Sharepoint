declare interface IExternalLibraryWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'ExternalLibraryWebPartStrings' {
  const strings: IExternalLibraryWebPartStrings;
  export = strings;
}
