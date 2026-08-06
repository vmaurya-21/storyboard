import * as React from 'react';
import type { ISpxConnectStoryDragDropProps } from './ISpxConnectStoryDragDropProps';
import StoryDragDrop from './StoryDragDrop';

/**
 * React root of the web part.
 *
 * A thin adapter: it forwards the SharePoint context to {@link StoryDragDrop},
 * which owns all board state and data access. The remaining props exist for
 * the SPFx scaffold and are unused here.
 */
export default class SpxConnectStoryDragDrop extends React.Component<ISpxConnectStoryDragDropProps> {
  /**
   * Renders the storyboard.
   *
   * @returns The board element.
   */
  public render(): React.ReactElement<ISpxConnectStoryDragDropProps> {
    return (
      <StoryDragDrop
        context={this.props.context}
        availableStoriesListName={this.props.availableStoriesListName}
        scheduledStoriesListName={this.props.scheduledStoriesListName}
        publishedStoriesListName={this.props.publishedStoriesListName}
      />
    );
  }
}
