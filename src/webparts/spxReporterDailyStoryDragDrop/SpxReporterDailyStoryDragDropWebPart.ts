import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SpxReporterDailyStoryDragDropWebPartStrings';
import SpxReporterDailyStoryDragDrop from './components/SpxReporterDailyStoryDragDrop';
import { ISpxReporterDailyStoryDragDropProps } from './components/ISpxReporterDailyStoryDragDropProps';

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
export interface ISpxReporterDailyStoryDragDropWebPartProps {
  /** Free-text description configured by the page author. */
  description: string;
}

/**
 * SPFx entry point for the Reporter Daily storyboard.
 *
 * Owns the SharePoint lifecycle — mounting React, tracking the host theme and
 * unmounting on dispose — and delegates all board behaviour to
 * {@link SpxReporterDailyStoryDragDrop}.
 */
export default class SpxReporterDailyStoryDragDropWebPart extends BaseClientSideWebPart<ISpxReporterDailyStoryDragDropWebPartProps> {

  /** Whether the host page is currently using an inverted (dark) theme. */
  private _isDarkTheme: boolean = false;

  /**
   * Mounts the React tree into the web part's DOM element.
   *
   * Called by SharePoint on first render and again whenever a property-pane
   * value changes.
   */
  public render(): void {
    const element: React.ReactElement<ISpxReporterDailyStoryDragDropProps> = React.createElement(
      SpxReporterDailyStoryDragDrop,
      {
        description: this.properties.description,
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
   * Schema version of {@link ISpxReporterDailyStoryDragDropWebPartProps}.
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
