import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLabel
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxInTeamsWebPartStrings';
import SpfxInTeams from './components/SpfxInTeams';
import { ISpfxInTeamsProps } from './components/ISpfxInTeamsProps';

export interface ISpfxInTeamsWebPartProps {
  subTitle: string;
}

export default class SpfxInTeamsWebPart extends BaseClientSideWebPart<ISpfxInTeamsWebPartProps> {

  private _teamsContext: microsoftTeams.Context;

  public onInit(): Promise<void> {
    let retVal: Promise<void> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }

  public render(): void {
    let title: string = '';
    let siteTabTitle: string = '';

    if (this._teamsContext) {
      // We have teams context for the web part
      title = 'Welcome to Teams!';
      siteTabTitle = 'We are in the context of following Team: ' + this._teamsContext.teamName;
    } else {
      // We are rendered in normal SharePoint context
      title = 'Welcome to SharePoint!';
      siteTabTitle = 'We are in the context of following site: ' + this.context.pageContext.web.title;
    }

    const element: React.ReactElement<ISpfxInTeamsProps> = React.createElement(
      SpfxInTeams,
      {
        title,
        subTitle: this.properties.subTitle,
        siteTabTitle
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
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
                PropertyPaneTextField('subTitle', {
                  label: strings.SubTitleFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
