declare interface IShowAllUsersWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  SearchFor: string;
  SearchForValidationErrorMessage: string;
}

declare module "ShowAllUsersWebPartStrings" {
  const strings: IShowAllUsersWebPartStrings;
  export = strings;
}
