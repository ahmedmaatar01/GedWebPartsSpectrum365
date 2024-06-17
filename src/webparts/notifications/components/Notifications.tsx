import * as React from 'react';
import styles from './Notifications.module.scss';
import type { INotificationsProps } from './INotificationsProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class Notifications extends React.Component<INotificationsProps, {}> {
  public render(): React.ReactElement<INotificationsProps> {
    const {
      description,
      environmentMessage,
      hasTeamsContext,
    } = this.props;

    return (
      <section className={`${styles.notifications} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <h2>Notifications</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>

      </section>
    );
  }
}
