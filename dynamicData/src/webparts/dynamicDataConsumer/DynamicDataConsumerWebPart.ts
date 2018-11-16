import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDynamicField,
  IWebPartPropertiesMetadata,
  IPropertyPaneConditionalGroup
} from '@microsoft/sp-webpart-base';

import DynamicDataConsumer from './components/DynamicDataConsumer';
import { IDynamicDataConsumerProps } from './components/IDynamicDataConsumerProps';
import { DynamicProperty } from '@microsoft/sp-component-base';

export interface IDynamicDataConsumerWebPartProps {
  description: string;
  // New type DynamicProperty<T>
  // If you use this type and properties metadata, SPFx takes care of everything
  bgColor: DynamicProperty<string>;
}

export default class DynamicDataConsumerWebPart extends BaseClientSideWebPart<IDynamicDataConsumerWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDynamicDataConsumerProps> = React.createElement(
      DynamicDataConsumer,
      {
        description: this.properties.description,
        // tryGetValue will get the current value for the dynamic property
        bgColor: this.properties.bgColor.tryGetValue()
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

  // You need to declare your properties with the appropriate metadata
  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      'bgColor': {
        dynamicPropertyType: 'string'
      }
    };
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            // Use primary and secondary groups to allows users to select both by value and reference
            {
              primaryGroup: {
                groupName: 'Dynamic data by value',
                groupFields: [
                  PropertyPaneTextField('bgColor', {
                    label: 'Background color'
                  })
                ]
              },
              secondaryGroup: {
                groupName: 'Dynamic data by reference',
                groupFields: [
                  PropertyPaneDynamicField('bgColor', {
                    label: 'Background color'
                  })
                ]
              },
              showSecondaryGroup: !!this.properties.bgColor.tryGetSource()
            } as IPropertyPaneConditionalGroup
          ]
        }
      ]
    };
  }
}
