import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { AadHttpClient } from '@microsoft/sp-http';

export interface IShowPageContextIssueProps {
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
  onConfigure: () => void;
  title: string;
  context: WebPartContext;
  aadHttpClient: AadHttpClient;
  useFakeData: boolean;
  sourceList: string;
}
