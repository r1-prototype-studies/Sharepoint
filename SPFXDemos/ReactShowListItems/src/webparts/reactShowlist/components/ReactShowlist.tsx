import * as React from "react";
import styles from "./ReactShowlist.module.scss";
import { IReactShowlistProps } from "./IReactShowlistProps";
import { escape } from "@microsoft/sp-lodash-subset";

import * as jquery from "jquery";

export interface IReactShowlistItemsWPState {
  listItems: [
    {
      Title: "";
      ID: "";
      SoftwareName: "";
    }
  ];
}

export default class ReactShowlist extends React.Component<
  IReactShowlistProps,
  IReactShowlistItemsWPState,
  {}
> {
  static siteurl: string = "";
  public constructor(
    props: IReactShowlistProps,
    state: IReactShowlistItemsWPState
  ) {
    super(props);
    this.state = {
      listItems: [
        {
          Title: "",
          ID: "",
          SoftwareName: "",
        },
      ],
    };
    ReactShowlist.siteurl = this.props.websiteUrl;
  }

  public componentDidMount() {
    let reactcontexthandler = this;

    jquery.ajax({
      url: `${ReactShowlist.siteurl}/_api/web/lists/getbytitle('SampleList')/items`,
      type: "GET",
      headers: { Accept: "application/json; odata=verbose" },
      success: function (resultdata: any) {
        reactcontexthandler.setState({ listItems: resultdata.d.results });
      },
      error: function (jqXHR, textStatus, errorThrown) {},
    });
  }

  public render(): React.ReactElement<IReactShowlistProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <div className={styles.reactShowlist}>
        <table className={styles.row}>
          {this.state.listItems.map(function (listitem, listitemkey) {
            let fullurl: string = `${ReactShowlist.siteurl}/lists/SampleList/DispForm.aspx?ID=${listitem.ID}`;
            return (
              <tr>
                <td className={styles.label}>{listitem.ID}</td>
                <td>
                  <a className={styles.label} href={fullurl}>
                    {listitem.Title}
                  </a>
                </td>

                <td className={styles.label}>{listitem.SoftwareName}</td>
              </tr>
            );
          })}
        </table>

        <ol>
          {this.state.listItems.map(function (listitem, listitemkey) {
            let fullurl: string = `${ReactShowlist.siteurl}/lists/SampleList/DispForm.aspx?ID=${listitem.ID}`;
            return (
              <li>
                <a className={styles.label} href={fullurl}>
                  <span>{listitem.ID}</span>,<span>{listitem.Title}</span>,
                  <span>{listitem.SoftwareName}</span>
                </a>
              </li>
            );
          })}
        </ol>
      </div>
    );
  }
}
