import * as React from "react";
import styles from "./ReactLifeCycleWebPart.module.scss";
import { IReactLifeCycleWebPartProps } from "./IReactLifeCycleWebPartProps";
import { escape } from "@microsoft/sp-lodash-subset";

export interface IReactLifeCycleWebPartState {
  stageTitle: string;
}
export default class ReactLifeCycleWebPart extends React.Component<
  IReactLifeCycleWebPartProps,
  IReactLifeCycleWebPartState,
  {}
> {
  public constructor(
    props: IReactLifeCycleWebPartProps,
    state: IReactLifeCycleWebPartState
  ) {
    super(props);
    this.state = {
      stageTitle: "component Constructor has been called",
    };
    this.updateState = this.updateState.bind(this);

    console.log("Stage Title from constructor : " + this.state.stageTitle);
  }

  public componentWillMount() {
    console.log("component will mount has been called");
  }

  public componentDidMount() {
    console.log(
      "Stage Title from componentDidMount : " + this.state.stageTitle
    );
    this.setState({
      stageTitle: "componentDidMount has been called",
    });
  }

  public updateState() {
    this.setState({
      stageTitle: "updateState has been called",
    });
  }

  public render(): React.ReactElement<IReactLifeCycleWebPartProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <div>
        <h1>ReactJS component's Lifecycle</h1>
        <h3>{this.state.stageTitle}</h3>
        <button onClick={this.updateState}>
          Click here to Update State Data!
        </button>
      </div>
    );
  }

  public componentWillUnmount() {
    console.log("component will unmount has been called");
  }
}
