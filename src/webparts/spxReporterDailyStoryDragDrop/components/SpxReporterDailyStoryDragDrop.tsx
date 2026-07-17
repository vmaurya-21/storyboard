import * as React from 'react';
import type { ISpxReporterDailyStoryDragDropProps } from './ISpxReporterDailyStoryDragDropProps';
import StoryDragDrop from './StoryDragDrop';

export default class SpxReporterDailyStoryDragDrop extends React.Component<ISpxReporterDailyStoryDragDropProps> {
  public render(): React.ReactElement<ISpxReporterDailyStoryDragDropProps> {
    return (
      <StoryDragDrop context={this.props.context} />
    );
  }
}
