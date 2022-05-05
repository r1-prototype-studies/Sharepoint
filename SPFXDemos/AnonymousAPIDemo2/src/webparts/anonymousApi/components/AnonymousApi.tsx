import * as React from "react";
import styles from "./AnonymousApi.module.scss";
import { IAnonymousApiProps } from "./IAnonymousApiProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { IAnonymousState } from "./IAnonymousState";

import { HttpClient, HttpClientResponse } from "@microsoft/sp-http";
export default class AnonymousApi extends React.Component<
  IAnonymousApiProps,
  IAnonymousState
> {
  public constructor(props: IAnonymousApiProps, state: IAnonymousState) {
    super(props);

    this.state = {
      id: null,
      name: null,
      username: null,
      email: null,
      address: null,
      phone: null,
      website: null,
      company: null,
    };
  }

  public getUserDetails(): Promise<any> {
    return this.props.context.httpClient
      .get(
        `${this.props.apiUrl}/${this.props.userID}`,
        HttpClient.configurations.v1
      )
      .then((response: HttpClientResponse) => {
        return response.json();
      })
      .then((jsonResponse) => {
        return jsonResponse;
      }) as Promise<any>;
  }

  public invokeApiandSetDatainState() {
    this.getUserDetails().then((userDetails) => {
      this.setState({
        id: userDetails.id,
        name: userDetails.name,
        username: userDetails.username,
        email: userDetails.email,
        address: `Street: ${userDetails.address.street}  Suite: ${userDetails.address.suite}  City: ${userDetails.address.city}`,
        phone: userDetails.phone,
        website: userDetails.website,
        company: userDetails.company.name,
      });
    });
  }

  public componentDidMount() {
    this.invokeApiandSetDatainState();
  }

  public componentDidUpdate(
    prevProps: IAnonymousApiProps,
    prevState: IAnonymousState,
    prevcontext: any
  ) {
    this.invokeApiandSetDatainState();
  }

  public render(): React.ReactElement<IAnonymousApiProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <div
        className={`${styles.anonymousApi} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <span>User Details:</span>

        <div>
          <strong>ID: </strong>
          {this.state.id}
        </div>
        <br />
        <div>
          <strong>User Name: </strong>
          {this.state.username}
        </div>
        <br />
        <div>
          <strong>Name: </strong>
          {this.state.name}
        </div>
        <br />
        <div>
          <strong>Address: </strong>
          {this.state.address}
        </div>
        <br />
        <div>
          <strong>Email: </strong>
          {this.state.email}
        </div>
        <br />
        <div>
          <strong>Phone: </strong>
          {this.state.phone}
        </div>
        <br />
        <div>
          <strong>WebSite: </strong>
          {this.state.website}
        </div>
        <br />
        <div>
          <strong>Company: </strong>
          {this.state.company}
        </div>
        <br />
      </div>
    );
  }
}
