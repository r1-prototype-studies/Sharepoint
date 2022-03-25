declare interface IGetListOfListDemoWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'GetListOfListDemoWebPartStrings' {
  const strings: IGetListOfListDemoWebPartStrings;
  export = strings;
}
