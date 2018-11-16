import * as React from 'react';
import styles from './DynamicDataConsumer.module.scss';
import { IDynamicDataConsumerProps } from './IDynamicDataConsumerProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class DynamicDataConsumer extends React.Component<IDynamicDataConsumerProps, {}> {
  public render(): React.ReactElement<IDynamicDataConsumerProps> {
    return (
      <div className={styles.dynamicDataConsumer}>
        <div className={styles.container}>
          {/* The background is defined as an inline style */}
          <div className={styles.row} style={{backgroundColor: this.props.bgColor}}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
