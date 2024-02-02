import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICalendarSummaryWebPartProps {
  apiKey: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}
