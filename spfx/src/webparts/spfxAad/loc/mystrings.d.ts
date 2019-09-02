declare interface ISpfxAadWebPartStrings {
  PropertyPaneDescription: string;
  GroupName: string;
  ClientIDFieldLabel: string;
  APIUrlFieldLabel: string;
}

declare module 'SpfxAadWebPartStrings' {
  const strings: ISpfxAadWebPartStrings;
  export = strings;
}
