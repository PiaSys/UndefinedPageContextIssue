import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ShowPageContextIssueWebPartStrings';
import ShowPageContextIssue from './components/ShowPageContextIssue';
import { IShowPageContextIssueProps } from './components/IShowPageContextIssueProps';

import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import { AadHttpClient } from '@microsoft/sp-http';

export interface IShowPageContextIssueWebPartProps {
  title: string;
  useFakeData: boolean;
  sourceList: string;
}

export default class ShowPageContextIssueWebPart extends BaseClientSideWebPart<IShowPageContextIssueWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _aadHttpClient: AadHttpClient = null;

  protected async onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    this._aadHttpClient = await this.context.aadHttpClientFactory.getClient("https://graph.microsoft.com");

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IShowPageContextIssueProps> = React.createElement(
      ShowPageContextIssue,
      {
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        title: this.properties.title,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        onConfigure: () => {
          this.context.propertyPane.open();
        },
        context: this.context,
        aadHttpClient: this._aadHttpClient,
        useFakeData: this.properties.useFakeData,
        sourceList: this.properties.sourceList
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

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
                PropertyFieldListPicker('sourceList', {
                  label: strings.SourceListPropertyPaneLabel,
                  selectedList: this.properties.sourceList,
                  includeHidden: false,
                  multiSelect: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'sourceList'
                }),
                PropertyPaneToggle('useFakeData', {
                  key: 'useFakeData',
                  label: strings.UseFakeDataPropertyPaneLabel,
                  onAriaLabel: strings.UseFakeDataPropertyPaneAriaLabel,
                  checked: this.properties.useFakeData
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
