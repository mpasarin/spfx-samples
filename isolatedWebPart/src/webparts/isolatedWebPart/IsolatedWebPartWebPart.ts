import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import styles from './IsolatedWebPartWebPart.module.scss';
import * as strings from 'IsolatedWebPartWebPartStrings';

export interface IIsolatedWebPartWebPartProps {
  description: string;
}

export default class IsolatedWebPartWebPart extends BaseClientSideWebPart<IIsolatedWebPartWebPartProps> {
  private _user: MicrosoftGraph.User;

  public render(): void {

    if (!this._user) {
    this.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        // get information about the current user from the Microsoft Graph
        client
          .api('/me')
          .get((error, user: MicrosoftGraph.User) => {
            this._user = user;
            this.render();
          });
      });
    }

    this.domElement.innerHTML = `
      <div class="${ styles.isolatedWebPart}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle}">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.subTitle}">Who am I? ${this._user ? this._user.displayName : 'Who knows'}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button}">
                <span class="${ styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
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
