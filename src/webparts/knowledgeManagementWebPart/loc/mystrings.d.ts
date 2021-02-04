declare interface IKnowledgeManagementWebPartWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  SearchBox: {
    DefaultPlacerHolder: "Enter your search..."
  }
}

declare module 'KnowledgeManagementWebPartWebPartStrings' {
  const strings: IKnowledgeManagementWebPartWebPartStrings;
  export = strings;
}
