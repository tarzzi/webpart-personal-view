declare interface IWebpartPersonalViewWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  RetrievableItemsGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  GreetingFieldLabel: string;
  SubGreetingFieldLabel: string;
  ShowGreetingFieldLabel: string;
  GreetingFieldPrefix: string;
  GreetingFieldSuffix: string;
  GreetingFieldShowUser: string;
  MailRetrieveCountFieldLabel: string;
  EventRetrieveCountFieldLabel: string;
  FileRetrieveCountFieldLabel: string;
}

declare module 'WebpartPersonalViewWebPartStrings' {
  const strings: IWebpartPersonalViewWebPartStrings;
  export = strings;
}
