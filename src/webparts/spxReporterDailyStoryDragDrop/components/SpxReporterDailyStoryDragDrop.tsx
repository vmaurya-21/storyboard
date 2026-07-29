import * as React from 'react';
import type { ISpxReporterDailyStoryDragDropProps } from './ISpxReporterDailyStoryDragDropProps';
import StoryDragDrop from './StoryDragDrop';

/**
 * React root of the web part.
 *
 * A thin adapter: it forwards the SharePoint context to {@link StoryDragDrop},
 * which owns all board state and data access. The remaining props exist for
 * the SPFx scaffold and are unused here.
 */
export default class SpxReporterDailyStoryDragDrop extends React.Component<ISpxReporterDailyStoryDragDropProps> {
  /**
   * Renders the storyboard.
   *
   * @returns The board element.
   */
  public render(): React.ReactElement<ISpxReporterDailyStoryDragDropProps> {
    return (
      <StoryDragDrop context={this.props.context} />
    );
  }
}
