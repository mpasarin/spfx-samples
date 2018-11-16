import * as React from 'react';
import styles from './SpfxInTeams.module.scss';
import { ISpfxInTeamsProps } from './ISpfxInTeamsProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SpfxInTeams extends React.Component<ISpfxInTeamsProps, {}> {
  public render(): React.ReactElement<ISpfxInTeamsProps> {
    return (
      <div className={ styles.spfxInTeams }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>{this.props.title}</span>
              <p className={ styles.subTitle }>{this.props.subTitle}</p>
              <p className={ styles.description }>{this.props.siteTabTitle}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
