import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import DynamicDataSource from './components/DynamicDataSource';
import { IDynamicDataSourceProps } from './components/IDynamicDataSourceProps';
import DynamicDataSourceController from './DynamicDataSourceController';

export interface IDynamicDataSourceWebPartProps {
  color: string;
}

export default class DynamicDataSourceWebPart extends BaseClientSideWebPart<IDynamicDataSourceWebPartProps> {
  private _dynamicDataSourceController: DynamicDataSourceController;

  public onInit(): Promise<void> {
    this._dynamicDataSourceController = new DynamicDataSourceController(this.properties);
    this.context.dynamicDataSourceManager.initializeSource(this._dynamicDataSourceController);
    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IDynamicDataSourceProps > = React.createElement(
      DynamicDataSource,
      {
        properties: this.properties,
        dynamicDataSourceManager: this.context.dynamicDataSourceManager
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
}
