import * as React from "react";
import styles from "./ShowAllUsers.module.scss";
import { IShowAllUsersProps } from "./IShowAllUsersProps";
import { escape } from "@microsoft/sp-lodash-subset";

import { IUser } from "./IUser";
import { IShowAllUsersState } from "./IShowAllUsersState";

import { MSGraphClient } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

import {
  TextField,
  //autobind,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
} from "office-ui-fabric-react";

import * as strings from "ShowAllUsersWebPartStrings";
export default class ShowAllUsers extends React.Component<
  IShowAllUsersProps,
  {}
> {
  constructor(props: IShowAllUsersProps, state: IShowAllUsersState) {
    super(props);

    this.state = {
      users: [],
      searchFor: "Aroan",
    };
  }

  public render(): React.ReactElement<IShowAllUsersProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <div className={styles.showAllUsers}>
        <TextField
          label={strings.SearchFor}
          required={true}
          value={this.state.searchFor}
          onChanged={this._onSearchForChanged}
          onGetErrorMessage={this._getSearchForErrorMessage}
        />

        <p className={styles.title}>
          <PrimaryButton text="Search" title="Search" onClick={this._search} />
        </p>

        {this.state.users != null && this.state.users.length > 0 ? (
          <p className={styles.row}>
            <DetailsList
              items={this.state.users}
              columns={_usersListColumns}
              setKey="set"
              checkboxVisibility={CheckboxVisibility.onHover}
              selectionMode={SelectionMode.single}
              layoutMode={DetailsListLayoutMode.fixedColumns}
              compact={true}
            />
          </p>
        ) : null}
      </div>
    );
  }
}
