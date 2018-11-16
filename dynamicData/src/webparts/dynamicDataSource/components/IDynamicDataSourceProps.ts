import { DynamicDataSourceManager } from '@microsoft/sp-component-base';
import { IDynamicDataSourceWebPartProps } from '../DynamicDataSourceWebPart';

export interface IDynamicDataSourceProps {
  properties: IDynamicDataSourceWebPartProps;
  dynamicDataSourceManager: DynamicDataSourceManager;
}
