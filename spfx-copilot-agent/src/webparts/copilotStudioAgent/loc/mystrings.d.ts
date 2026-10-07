declare interface ICopilotStudioAgentWebPartStrings {
  PropertyPaneDescription: string;
  AgentGroupName: string;
  AuthGroupName: string;
  AppearanceGroupName: string;
  TitleFieldLabel: string;
  ModeFieldLabel: string;
  ModeSdk: string;
  ModeEmbed: string;
  EmbedUrlFieldLabel: string;
  EmbedUrlFieldDescription: string;
  EmbedUrlInvalid: string;
  ConnectionStringFieldLabel: string;
  ConnectionStringFieldDescription: string;
  ConnectionStringInvalid: string;
  OrLabel: string;
  EnvironmentIdFieldLabel: string;
  AgentIdentifierFieldLabel: string;
  AppClientIdFieldLabel: string;
  AppClientIdFieldDescription: string;
  TenantIdFieldLabel: string;
  TenantIdFieldDescription: string;
  GuidInvalid: string;
  ShowHeaderFieldLabel: string;
  FluentThemeFieldLabel: string;
  HeightFieldLabel: string;
}

declare module 'CopilotStudioAgentWebPartStrings' {
  const strings: ICopilotStudioAgentWebPartStrings;
  export = strings;
}
