import { Log } from '@microsoft/sp-core-library';
import * as React from 'react';

import styles from './FecDemo.module.scss';

export interface IFecDemoProps {
  text: string;
}

const LOG_SOURCE: string = 'FecDemo';

export default class FecDemo extends React.Component<IFecDemoProps, {}> {
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: FecDemo mounted');
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: FecDemo unmounted');
  }

  public render(): React.ReactElement<{}> {

    const myStyles = {
      color: 'blue',
      width: `${ this.props.text }px`,
      background: 'green',
    }

    return (

      <div className={styles.FecDemo}>
        <div className={styles.cell}>
          <div style={myStyles}>
            {this.props.text}
          </div>
        </div>
      </div>

    );
  }
}
