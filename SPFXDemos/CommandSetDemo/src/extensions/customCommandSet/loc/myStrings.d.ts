declare interface ICustomCommandSetCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'CustomCommandSetCommandSetStrings' {
  const strings: ICustomCommandSetCommandSetStrings;
  export = strings;
}
