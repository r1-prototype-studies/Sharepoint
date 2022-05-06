declare interface IShowAllUsersWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'ShowAllUsersWebPartStrings' {
  const strings: IShowAllUsersWebPartStrings;
  export = strings;
}
