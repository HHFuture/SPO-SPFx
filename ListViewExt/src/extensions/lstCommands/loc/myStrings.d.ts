declare interface ILstCommandsCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'LstCommandsCommandSetStrings' {
  const strings: ILstCommandsCommandSetStrings;
  export = strings;
}
