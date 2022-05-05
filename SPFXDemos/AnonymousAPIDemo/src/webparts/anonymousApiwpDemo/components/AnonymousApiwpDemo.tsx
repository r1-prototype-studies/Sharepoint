import * as React from "react";
import styles from "./AnonymousApiwpDemo.module.scss";
import { IAnonymousApiwpDemoProps } from "./IAnonymousApiwpDemoProps";
import { escape } from "@microsoft/sp-lodash-subset";

export default class AnonymousApiwpDemo extends React.Component<
  IAnonymousApiwpDemoProps,
  {}
> {
  public render(): React.ReactElement<IAnonymousApiwpDemoProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <div
        className={`${styles.anonymousApiwpDemo} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <span>User Details:</span>

        <div>
          <strong>ID: </strong>
          {this.props.id}
        </div>
        <br />
        <div>
          <strong>User Name: </strong>
          {this.props.username}
        </div>
        <br />
        <div>
          <strong>Name: </strong>
          {this.props.name}
        </div>
        <br />
        <div>
          <strong>Address: </strong>
          {this.props.address}
        </div>
        <br />
        <div>
          <strong>Email: </strong>
          {this.props.email}
        </div>
        <br />
        <div>
          <strong>Phone: </strong>
          {this.props.phone}
        </div>
        <br />
        <div>
          <strong>WebSite: </strong>
          {this.props.website}
        </div>
        <br />
        <div>
          <strong>Company: </strong>
          {this.props.company}
        </div>
        <br />
      </div>
    );
  }
}
