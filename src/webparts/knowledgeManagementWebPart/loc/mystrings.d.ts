declare interface IKnowledgeManagementWebPartWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  InvalidDataSourceInstance: string;
  SearchBox: {
    DefaultPlacerHolder: string;
  }
  PropertyPanel: {
    DisplayDescription: string;
    HideShowSearchBar: string;
    HideShowRefinementPanel: string;
    ToggleResultYes: string;
    ToggleResultNo: string;
    ThemeLayout: {
      Theme: string;
      ThemeDescription: string;
      ListView: string;
      GridView: string;
    }
  }
}

declare module 'KnowledgeManagementWebPartWebPartStrings' {
  const strings: IKnowledgeManagementWebPartWebPartStrings;
  export = strings;
}
