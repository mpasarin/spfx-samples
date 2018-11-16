import { ColorPicker } from 'office-ui-fabric-react/lib/ColorPicker';
import * as React from 'react';
import styles from './DynamicDataSource.module.scss';
import { IDynamicDataSourceProps } from './IDynamicDataSourceProps';

export default class DynamicDataSource extends React.Component<IDynamicDataSourceProps, {}> {
  public render(): React.ReactElement<IDynamicDataSourceProps> {
    return (
      <div>
        <h1>Sample dynamic data source</h1>
        <p>Choose a color to be set as dynamic data.</p>
        <ColorPicker
          color={this.props.properties.color}
          onColorChanged={this._onColorChanged.bind(this)}
        />
      </div>
    );
  }

  private _onColorChanged(color: string) {
    this.props.properties.color = color;
    this.props.dynamicDataSourceManager.notifyPropertyChanged('color');
  }
}
