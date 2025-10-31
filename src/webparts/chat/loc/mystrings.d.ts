declare interface IChatWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  UseSiteThemeLabel: string;
  PrimaryButtonColorLabel: string;
  PrimaryButtonTextColorLabel: string;
  SurfaceBorderColorLabel: string;
  MessageBorderColorLabel: string;
  SelfMessageBackgroundColorLabel: string;
  MentionBackgroundColorLabel: string;
  MentionTextColorLabel: string;
  ColorHelpText: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'ChatWebPartStrings' {
  const strings: IChatWebPartStrings;
  export = strings;
}
