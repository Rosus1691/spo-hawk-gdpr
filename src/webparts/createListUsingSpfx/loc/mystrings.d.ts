declare interface ICreateListUsingSpfxWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'CreateListUsingSpfxWebPartStrings' {
  const strings: ICreateListUsingSpfxWebPartStrings;
  export = strings;
}
