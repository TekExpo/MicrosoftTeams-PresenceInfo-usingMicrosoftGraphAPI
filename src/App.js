import React, { Component } from "react";
import hello from "hellojs";
import GraphSdkHelper from "./helpers/GraphSdkHelper";
import { CommandBar } from "office-ui-fabric-react/lib/CommandBar";
import PeoplePickerExample from "./component/PeoplePicker";
import DetailsListExample from "./component/DetailsList";
import SearchExample from "./component/Search";
import PresenceExample from "./component/Presence";
import { applicationId, redirectUri } from "./helpers/config";

window.hello = hello;

export default class App extends Component {
  constructor(props) {
    super(props);

    // Initialize the auth network.
    hello.init({
      aad: {
        name: "Azure Active Directory",
        oauth: {
          version: 2,
          auth: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
        },
        form: false
      }
    });

    // Initialize the Graph SDK helper and save it in the window object.
    this.sdkHelper = new GraphSdkHelper({ login: this.login.bind(this) });
    window.sdkHelper = this.sdkHelper;

    // Set the isAuthenticated prop and the (empty) Fabric example selection.
    this.state = {
      isAuthenticated: !!hello("aad").getAuthResponse(),
      example: ""
    };
  }

  // Get the user's display name.
  componentWillMount() {
    if (this.state.isAuthenticated) {
      this.sdkHelper.getMe((err, me) => {
        if (!err) {
          this.setState({
            displayName: `Hello ${me.displayName}!`
          });
        }
      });
    }
  }

  // Sign the user into Azure AD. HelloJS stores token info in localStorage.hello.
  login() {
    // Initialize the auth request.
    hello.init(
      {
        aad: applicationId
      },
      {
        redirect_uri: redirectUri,
        scope: "user.readbasic.all+mail.send+files.read"
      }
    );

    hello.login("aad", {
      display: "page",
      state: "abcd"
    });
  }

  // Sign the user out of the session.
  logout() {
    hello("aad").logout();
    this.setState({
      isAuthenticated: false,
      example: "",
      displayName: ""
    });
  }

  render() {
    return (
      <div>
        <div>
          {
            // Show the command bar with the Sign in or Sign out button.
            <CommandBar
              items={[
                {
                  key: "component-example-menu",
                  name: "Choose component",
                  disabled: !this.state.isAuthenticated,
                  ariaLabel: "Choose a component example to render in the page",
                  items: [
                  /*  {
                      key: "people-picker-example",
                      name: "People Picker",
                      onClick: () => {
                        this.setState({ example: "people-picker-example" });
                      }
                    },
                    {
                      key: "details-list-example",
                      name: "Details List",
                      onClick: () => {
                        this.setState({ example: "details-list-example" });
                      }
                    },
                    {
                      key: "search-example",
                      name: "Search",
                      onClick: () => {
                        this.setState({ example: "search-example" });
                      }
                    }, */
                    {
                      key: "presence-example",
                      name: "Presence",
                      onClick: () => {
                        this.setState({ example: "presence-example" });
                      }
                    }
                  ]
                }
              ]}
              farItems={[
                {
                  key: "display-name",
                  name: this.state.displayName
                },
                {
                  key: "log-in-out=button",
                  name: this.state.isAuthenticated ? "Sign out" : "Sign in",
                  onClick: this.state.isAuthenticated
                    ? this.logout.bind(this)
                    : this.login.bind(this)
                }
              ]}
            />
          }
        </div>
        <div className="ms-font-m">
          <div>
            <h2>Demo - Using Microsoft Graph</h2>
            {(!this.state.isAuthenticated || this.state.example === "") && (
              <div>
                <p>Use Microsoft Graph data in React components.</p>
              </div>
            )}
          </div>
          <br />
          {// Show the selected fabric component example.
          this.state.isAuthenticated && (
            <div>
              {this.state.example === "people-picker-example" && (
                <PeoplePickerExample />
              )}
              {this.state.example === "details-list-example" && (
                <DetailsListExample />
              )}
              {this.state.example === "search-example" && <SearchExample />}
              {this.state.example === "presence-example" && <PresenceExample />}
            </div>
          )}
          <br />
        </div>
      </div>
    );
  }
}
