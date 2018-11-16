import * as React from 'react';
import styles from './SingleWebPartAppPage.module.scss';
import { ISingleWebPartAppPageProps } from './ISingleWebPartAppPageProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SingleWebPartAppPage extends React.Component<ISingleWebPartAppPageProps, {}> {
  public render(): React.ReactElement<ISingleWebPartAppPageProps> {
    return (
      <div className={styles.singleWebPartAppPage}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>
                Welcome to SharePoint {this.props.isAppPage ? 'App Pages' : 'Web Parts'}!
              </span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>
        <p>To change the layout, run the following script:</p>
        <pre className={styles.codeBlock}>{this.props.scriptText}</pre>
        <p>
          Get more information at: <a href='https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/single-part-app-pages'>
            https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/single-part-app-pages
          </a>
        </p>
      </div>
    );
  }
}
