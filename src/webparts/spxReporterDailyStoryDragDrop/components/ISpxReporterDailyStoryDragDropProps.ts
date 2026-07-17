import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AppMode } from "../SpxReporterDailyStoryDragDropWebPart";

export interface ISpxReporterDailyStoryDragDropProps {
  description: string;
  isDarkTheme: boolean;
  context: WebPartContext;
  appMode: AppMode;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
