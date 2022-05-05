import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IAnonymousApiProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  apiUrl: string;
  userID: string;
  context: WebPartContext;
}
