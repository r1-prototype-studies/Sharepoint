declare interface ICrudWithReactWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  lblID: string;
  lblSoftwareTitle: string;
  lblSoftwareName: string;
  lblSoftwareDescription: string;
  lblSoftwareVendor: string;
  lblSoftwareVersion: string;
}

declare module 'CrudWithReactWebPartStrings' {
  const strings: ICrudWithReactWebPartStrings;
  export = strings;
}
