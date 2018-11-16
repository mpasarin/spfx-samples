import {
  IDynamicDataAnnotatedPropertyValue, IDynamicDataCallables, IDynamicDataPropertyDefinition
} from '@microsoft/sp-dynamic-data';
import { IDynamicDataSourceWebPartProps } from './DynamicDataSourceWebPart';

export default class DynamicDataSourceController implements IDynamicDataCallables {
  private _props: IDynamicDataSourceWebPartProps;

  constructor(props: IDynamicDataSourceWebPartProps) {
    this._props = props;
  }

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      {
        id: 'color',
        title: 'Selected color'
      }
    ];
  }

  // tslint:disable-next-line:no-any
  public getPropertyValue(propertyId: string): any {
    switch (propertyId) {
      case 'color':
        return this._props.color;
      default:
        throw new Error('Unsupported property id');
    }
  }

  public getAnnotatedPropertyValue?(propertyId: string): IDynamicDataAnnotatedPropertyValue {
    switch (propertyId) {
      case 'color':
        return {
          sampleValue: '#ff0000'
        };
      default:
        throw new Error('Unsupported property id');
    }
  }
}