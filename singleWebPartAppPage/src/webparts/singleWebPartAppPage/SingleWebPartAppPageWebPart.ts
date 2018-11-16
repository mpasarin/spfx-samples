import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  WebPartFormFactor
} from '@microsoft/sp-webpart-base';

import * as strings from 'SingleWebPartAppPageWebPartStrings';
import SingleWebPartAppPage from './components/SingleWebPartAppPage';
import { ISingleWebPartAppPageProps } from './components/ISingleWebPartAppPageProps';

export interface ISingleWebPartAppPageWebPartProps {
  description: string;
}

export default class SingleWebPartAppPageWebPart extends BaseClientSideWebPart<ISingleWebPartAppPageWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISingleWebPartAppPageProps> = React.createElement(
      SingleWebPartAppPage,
      {
        description: this.properties.description,
        isAppPage: this.context.formFactor === WebPartFormFactor.FullSize,
        scriptText: this.getScript()
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

  private getScript(): string {
    const siteUrl: string = this.context.pageContext.web.absoluteUrl;
    const queryStringIndex: number | undefined =
      document.location.href.indexOf('?') >= 0 ? document.location.href.indexOf('?') : undefined;
    const pageUrl: string = document.location.href.substring(siteUrl.length + 1, queryStringIndex);

return `fetch('${siteUrl}/_api/contextinfo', {
  method: 'POST',
  headers: {
    accept: 'application/json;odata=nometadata'
  }
})
.then(function (response) {
  return response.json();
})
.then(function (ctx) {
  return fetch("${siteUrl}/_api/web/getfilebyurl('${pageUrl}')/ListItemAllFields", {
    method: 'POST',
    headers: {
      accept: 'application/json;odata=nometadata',
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': '*',
      'X-RequestDigest': ctx.FormDigestValue,
      'content-type': 'application/json;odata=nometadata',
    },
    body: JSON.stringify({
      PageLayoutType: "${
        this.context.formFactor !== WebPartFormFactor.FullSize ? 'SingleWebPartAppPage' : 'Article'
      }"
    })
  })
})
.then(function(res) {
  console.log(res.ok ? 'DONE' : 'Error: ' + res.statusText);
});`;
  }
}
