import * as React from 'react';
import styles from './MgtReact.module.scss';
import { IMgtReactProps } from './IMgtReactProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Person } from '@microsoft/mgt-react';

export default class MgtReact extends React.Component<IMgtReactProps, {}> {
  public render(): React.ReactElement<IMgtReactProps> {
    return (
      <div className={ styles.mgtReact }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
            
              <span className={ styles.title }>  React Webpart</span>
              <Person personQuery="me" />
            </div>
          </div>
        </div>
      </div>
    );
  }
}
