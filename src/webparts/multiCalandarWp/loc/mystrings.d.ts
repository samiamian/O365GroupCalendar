declare interface IMultiCalandarWpWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  PropertyPaneSite: string;
  PropertyPaneColor: string;

  DescriptionFieldLabel: string;
  SiteFieldLabel: string;
  ColorFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'MultiCalandarWpWebPartStrings' {
  const strings: IMultiCalandarWpWebPartStrings;
  export = strings;
}
