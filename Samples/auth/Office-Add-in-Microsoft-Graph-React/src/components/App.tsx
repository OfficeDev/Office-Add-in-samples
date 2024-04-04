import * as React from "react";
import { Spinner, SpinnerType } from "office-ui-fabric-react";
import Header from "./Header";
import { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import StartPageBody from "./StartPageBody";
import GetDataPageBody from "./GetDataPageBody";
import SuccessPageBody from "./SuccessPageBody";
import OfficeAddinMessageBar from "./OfficeAddinMessageBar";
import { getGraphData } from "../../utilities/microsoft-graph-helpers";
import {
  writeFileNamesToWorksheet,
  logoutFromO365,
  signInO365,
} from "../../utilities/office-apis-helpers";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  authStatus?: string;
  fileFetch?: string;
  headerMessage?: string;
  errorMessage?: string;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      authStatus: "notLoggedIn",
      fileFetch: "notFetched",
      headerMessage: "Welcome",
      errorMessage: "",
    };

    // Bind the methods that we want to pass to, and call in, a separate
    // module to this component. And rename setState to boundSetState
    // so code that passes boundSetState is more self-documenting.
    this.boundSetState = this.setState.bind(this);
    this.setToken = this.setToken.bind(this);
    this.setUserName = this.setUserName.bind(this);
    this.displayError = this.displayError.bind(this);
    this.login = this.login.bind(this);
  }

  /*
        Properties
    */

  // The access token is not part of state because React is all about the
  // UI and the token is not used to affect the UI in any way.
  accessToken: string;
  userName: string;

  listItems: HeroListItem[] = [
    {
      icon: "PlugConnected",
      primaryText: "Connects to OneDrive for Business.",
    },
    {
      icon: "ExcelDocument",
      primaryText:
        "Gets the names of the first three workbooks in OneDrive for Business.",
    },
    {
      icon: "AddNotes",
      primaryText: "Adds the names to the current document.",
    },
  ];

  /*
        Methods
    */

  boundSetState: () => {};

  setToken = (accesstoken: string) => {
    this.accessToken = accesstoken;
  };

  setUserName = (userName: string) => {
    this.userName = userName;
  };

  displayError = (error: string) => {
    this.setState({ errorMessage: error, fileFetch: "notFetched" });
  };

  // Runs when the user clicks the X to close the message bar where
  // the error appears.
  errorDismissed = () => {
    this.setState({ errorMessage: "" });

    // If the error occured during a "in process" phase (logging in or getting files),
    // the action didn't complete, so return the UI to the preceding state/view.
    this.setState((prevState) => {
      if (prevState.authStatus === "loginInProcess") {
        return { authStatus: "notLoggedIn" };
      } else if (prevState.fileFetch === "fetchInProcess") {
        return { fileFetch: "notFetched" };
      }
      return null;
    });
  };

  login = async () => {
    await signInO365(
      this.boundSetState,
      this.setToken,
      this.setUserName,
      this.displayError
    );
  };

  logout = async () => {
    await logoutFromO365(
      this.boundSetState,
      this.setUserName,
      this.userName,
      this.displayError
    );
  };

  getFileNames = async () => {
    this.setState({ fileFetch: "fetchInProcess" });
    try {
      let response = await getGraphData(
        // Get the `name` property of the first 3 Excel workbooks in the user's OneDrive.
        "https://graph.microsoft.com/v1.0/me/drive/root/microsoft.graph.search(q = '.xlsx')?$select=name&top=3",
        this.accessToken
      );
      await writeFileNamesToWorksheet(response, this.displayError);
      this.setState({ fileFetch: "fetched", headerMessage: "Success" });
    } catch (requestError) {
      // This error must be
      // from the Axios request in getGraphData, not the Office.js in
      // writeFileNamesToWorksheet
      this.displayError(requestError);
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo="assets/Onedrive_Charts_icon_80x80px.png"
          message="Please sideload your add-in to see app body."
        />
      );
    }

    // Set the body of the page based on where the user is in the workflow.
    let body;

    if (this.state.authStatus === "notLoggedIn") {
      body = <StartPageBody login={this.login} listItems={this.listItems} />;
    } else if (this.state.authStatus === "loginInProcess") {
      body = (
        <Spinner
          className="spinner"
          type={SpinnerType.large}
          label="Please sign-in on the pop-up window."
        />
      );
    } else {
      if (this.state.fileFetch === "notFetched") {
        body = (
          <GetDataPageBody
            getFileNames={this.getFileNames}
            logout={this.logout}
          />
        );
      } else if (this.state.fileFetch === "fetchInProcess") {
        body = (
          <Spinner
            className="spinner"
            type={SpinnerType.large}
            label="We are getting the data for you."
          />
        );
      } else {
        body = (
          <SuccessPageBody
            getFileNames={this.getFileNames}
            logout={this.logout}
          />
        );
      }
    }

    return (
      <div>
        {this.state.errorMessage ? (
          <OfficeAddinMessageBar
            onDismiss={this.errorDismissed}
            message={this.state.errorMessage + " "}
          />
        ) : null}

        <div className="ms-welcome">
          <Header
            logo="assets/Onedrive_Charts_icon_80x80px.png"
            title={this.props.title}
            message={this.state.headerMessage}
          />
          {body}
        </div>
      </div>
    );
  }
}
