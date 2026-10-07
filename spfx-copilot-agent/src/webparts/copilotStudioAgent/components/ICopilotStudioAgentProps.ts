export interface ICopilotStudioAgentProps {
  title: string;
  /** 'sdk' = Microsoft 365 Agents SDK (default); 'embed' = Copilot Studio hosted web chat iframe. */
  mode: string;
  connectionString: string;
  embedUrl: string;
  environmentId: string;
  agentIdentifier: string;
  appClientId: string;
  tenantId: string;
  height: number;
  useFluentTheme: boolean;
  showHeader: boolean;
  userLoginName: string;
  userDisplayName: string;
  locale: string;
  isEditMode: boolean;
  isDarkTheme: boolean;
  accentColor?: string;
  onConfigure: () => void;
}
