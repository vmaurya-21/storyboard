import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SpxConnectStoryDragDropWebPartStrings';
import SpxConnectStoryDragDrop from './components/SpxConnectStoryDragDrop';
import { ISpxConnectStoryDragDropProps } from './components/ISpxConnectStoryDragDropProps';

/** Host the web part is running in. */
export enum AppMode {
  /** Hosted on a classic or modern SharePoint page. */
  Spfx = 'SPFX',
  /** Hosted inside a Microsoft Teams tab. */
  Teams = 'TEAMS',
  /** Running outside SharePoint, e.g. in local workbench or tests. */
  Standalone = 'STANDALONE'
}

/** Persisted property-pane settings for the web part. */
export interface ISpxConnectStoryDragDropWebPartProps {
  /** Free-text description configured by the page author. */
  description: string;
  /** SharePoint list title for available stories. */
  availableStoriesListName: string;
  /** SharePoint list title for scheduled stories. */
  scheduledStoriesListName: string;
  /** SharePoint list title for published stories. */
  publishedStoriesListName: string;
}

/**
 * SPFx entry point for the Connect storyboard.
 *
 * Owns the SharePoint lifecycle — mounting React, tracking the host theme and
 * unmounting on dispose — and delegates all board behaviour to
 * {@link SpxConnectStoryDragDrop}.
 */
export default class SpxConnectStoryDragDropWebPart extends BaseClientSideWebPart<ISpxConnectStoryDragDropWebPartProps> {

  /** Whether the host page is currently using an inverted (dark) theme. */
  private _isDarkTheme: boolean = false;

  /**
   * Mounts the React tree into the web part's DOM element.
   *
   * Called by SharePoint on first render and again whenever a property-pane
   * value changes.
   */
  public render(): void {
    const element: React.ReactElement<ISpxConnectStoryDragDropProps> = React.createElement(
      SpxConnectStoryDragDrop,
      {
        description: this.properties.description,
        availableStoriesListName: this.properties.availableStoriesListName,
        scheduledStoriesListName: this.properties.scheduledStoriesListName,
        publishedStoriesListName: this.properties.publishedStoriesListName,
        isDarkTheme: this._isDarkTheme,
        context: this.context,
        appMode: AppMode.Spfx,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  /**
   * Initialises the web part before its first render.
   *
   * Nothing is needed up front: the services take their SharePoint context
   * from props on first use.
   *
   * @returns A promise that resolves once initialisation is complete.
   */
  protected onInit(): Promise<void> {
    this.properties.description = this.properties.description || 'Connect Story board';
    this.properties.availableStoriesListName = this.properties.availableStoriesListName || 'AvailableStories';
    this.properties.scheduledStoriesListName = this.properties.scheduledStoriesListName || 'ScheduledStories';
    this.properties.publishedStoriesListName = this.properties.publishedStoriesListName || 'PublishedStories';
    return Promise.resolve();
  }

  /**
   * Tracks the host page's theme.
   *
   * Records whether the theme is inverted and republishes the semantic colours
   * the stylesheets consume as CSS custom properties.
   *
   * @param currentTheme - The new theme, or `undefined` when SharePoint has
   * none to report — in which case the previous values are left in place.
   */
  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  /** Unmounts the React tree so its effects and timers are torn down. */
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /**
   * Schema version of {@link ISpxConnectStoryDragDropWebPartProps}.
   *
   * @returns The version SharePoint uses to migrate persisted properties.
   */
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * Describes the property pane.
   *
   * @returns A single page with one group holding the description field.
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('availableStoriesListName', {
                  label: strings.AvailableStoriesListNameFieldLabel
                }),
                PropertyPaneTextField('scheduledStoriesListName', {
                  label: strings.ScheduledStoriesListNameFieldLabel
                }),
                PropertyPaneTextField('publishedStoriesListName', {
                  label: strings.PublishedStoriesListNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
