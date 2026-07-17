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

export enum AppMode {
  Spfx = 'SPFX',
  Teams = 'TEAMS',
  Standalone = 'STANDALONE'
}

export interface ISpxReporterDailyStoryDragDropWebPartProps {
  description: string;
}

export default class SpxReporterDailyStoryDragDropWebPart extends BaseClientSideWebPart<ISpxReporterDailyStoryDragDropWebPartProps> {

  private _isDarkTheme: boolean = false;

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

  protected onInit(): Promise<void> {
    return Promise.resolve();
  }



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

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

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
