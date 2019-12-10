declare interface ITeamCreatorWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  TeamNameLabel: string;
  TeamNickNameLabel: string;
  TeamDescriptionLabel: string;
  Owners: string;
  Members: string;
  CreateChannel: string;
  ChannelName: string;
  ChannelDescription: string;
  AddTab: string;
  AddTabToGeneral: string;
  TabName: string;
  App: string;
  Welcome: string;
  Create: string;
  Clear: string;
  CreatingGroup: string;
  CreatingTeam: string;
  CreatingChannel: string;
  InstallingApp: string;
  CreatingTab: string;
  Error: string;
  Success: string;
  StartOver: string;
  Again:string;
  OpenTeams: string;
  
  //  ボタン
  Department:string;
  Project:string;
  NewEmployee:string;
  Training:string;

  //  テンプレート一覧
  DepartmentDesc: string;
  ProjectDesc: string;
  NewEmployeeDesc: string;
  TrainingDesc: string;
}

declare module 'TeamCreatorWebPartStrings' {
  const strings: ITeamCreatorWebPartStrings;
  export = strings;
}
