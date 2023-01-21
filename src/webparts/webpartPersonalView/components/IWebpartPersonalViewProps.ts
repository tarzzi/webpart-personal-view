import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IWebpartPersonalViewProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context : WebPartContext;
  greetingPrefix: string;
  greetingSuffix: string;
  greetingShowUser: boolean;
  subGreeting: string;
  showGreeting: boolean;
  mailRetrieveCount: number;
  eventRetrieveCount: number;
  fileRetrieveCount: number;
}
