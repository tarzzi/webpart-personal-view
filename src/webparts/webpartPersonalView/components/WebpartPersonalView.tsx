import * as React from "react";
import styles from "./WebpartPersonalView.module.scss";
import { IWebpartPersonalViewProps } from "./IWebpartPersonalViewProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { SPHttpClientResponse } from "@microsoft/sp-http";

import MailItem from "./mail/MailItem";

interface IWebpartPersonalViewState {
  mail: any;
  calendar: any;
  todoLists: any;
}

export default class WebpartPersonalView extends React.Component<
  IWebpartPersonalViewProps,
  IWebpartPersonalViewState
> {
  constructor(props: IWebpartPersonalViewProps) {
    super(props);
    this.state = {
      mail: null,
      calendar: null,
      todoLists: null,
    };
  }

  private _getMail = (): void => {
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3) => {
        client
          .api("/me/mailFolders/inbox/messages")
          .version("v1.0")
          .get((err: SPHttpClientResponse, res: SPHttpClientResponse) => {
            if (err) {
              console.log(err);
              return;
            }
            this.setState({ mail: res });
            console.log(res);
          });
      })
      .catch((err: SPHttpClientResponse) => {
        console.log(err);
      });
  };

  
  private _getCalendar = () : void => {
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3) => {
        client
          .api("/me/calendar/events")
          .version("v1.0")
          .get((err: SPHttpClientResponse, res: SPHttpClientResponse) => {
            if (err) {
              console.log(err);
              return;
            }
            console.log(res);
          });
      })
      .catch((err: SPHttpClientResponse) => {
        console.log(err);
      });
  };

  private _getTodoLists = () : void => {
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3) => {
        client
          .api("/me/todo/lists")
          .version("v1.0")
          .get((err: SPHttpClientResponse, res: SPHttpClientResponse) => {
            if (err) {
              console.log(err);
              return;
            }
            console.log(res);
          });
      })
      .catch((err: SPHttpClientResponse) => {
        console.log(err);
      });
  };

  public render(): React.ReactElement<IWebpartPersonalViewProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <>
        <section
          className={`${styles.webpartPersonalView} ${
            hasTeamsContext ? styles.teams : ""
          }`}
        >
          <div>
            <h1>Hello {userDisplayName}!</h1>
            <p>Is dark theme? {isDarkTheme ? "Yes" : "No"}</p>
            <p>Desc: {escape(description)}</p>
            <p>Env: {escape(environmentMessage)}</p>
          </div>
        </section>
        
        <section
          className={`${styles.webpartPersonalView} ${
            hasTeamsContext ? styles.teams : ""
          }`}
        >

          <button onClick={this._getMail}>Get mail</button>
           
          {this.state.mail &&
            this.state.mail.value.map((mailItem: any) => {
              return <MailItem key={mailItem.id} mailItem={mailItem} />;
            })
            }
        </section>

        <section
          className={`${styles.webpartPersonalView} ${
            hasTeamsContext ? styles.teams : ""
          }`}
        >
          <button onClick={this._getCalendar}>Get calendar</button>
        </section>
        <section
          className={`${styles.webpartPersonalView} ${
            hasTeamsContext ? styles.teams : ""
          }`}
        >
          <button onClick={this._getTodoLists}>Get Todo lists</button>
        </section>
      </>
    );
  }
}
