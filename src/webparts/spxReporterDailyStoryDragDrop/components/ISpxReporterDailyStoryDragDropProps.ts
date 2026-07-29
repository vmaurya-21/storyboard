import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AppMode } from "../SpxReporterDailyStoryDragDropWebPart";

/**
 * Props the web part passes to its React root.
 *
 * Everything here is supplied by SharePoint at render time — either from the
 * web part's property pane or from its host context.
 */
export interface ISpxReporterDailyStoryDragDropProps {
  /** Free-text description configured in the property pane. */
  description: string;
  /** Whether the host page is using a dark theme. */
  isDarkTheme: boolean;
  /** SharePoint context, used to authenticate every list call. */
  context: WebPartContext;
  /** Whether the web part renders the editor or the read-only board. */
  appMode: AppMode;
  /** Whether the web part is hosted inside Microsoft Teams. */
  hasTeamsContext: boolean;
  /** Display name of the signed-in user. */
  userDisplayName: string;
}
