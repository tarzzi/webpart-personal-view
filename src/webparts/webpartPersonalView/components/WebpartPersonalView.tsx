import * as React from "react";
import styles from "./WebpartPersonalView.module.scss";
import { IWebpartPersonalViewProps } from "./IWebpartPersonalViewProps";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { SPHttpClientResponse } from "@microsoft/sp-http";
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Text } from "@fluentui/react/lib/Text";

// import icons for mail, calendar, files and todo
import { initializeIcons } from "@fluentui/react/lib/Icons";
initializeIcons();

import MailItem from "./mail/MailItem";
import CalendarItem from "./calendar/CalendarItem";
import DriveItem from "./drive/DriveItem";
import TodoItem from "./todo/TodoItem";

interface IWebpartPersonalViewState {
  mail: any;
  calendar: any;
  todoLists: any;
  files: any;
  activeView: string;
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
      files: null,
      activeView: "mail",
    };
  }

  private _changeActiveView = (view: string): void => {
    this.setState({ activeView: view });
  };

  public componentDidMount(): void {
    this._getMail();
    this._getCalendar();
    this._getRecentFiles();
  }

  public componentDidUpdate(prevProps: IWebpartPersonalViewProps): void {
    if (prevProps.mailRetrieveCount !== this.props.mailRetrieveCount &&  prevProps.mailRetrieveCount !== undefined) {
      this._getMail();
    }
    if (prevProps.eventRetrieveCount !== this.props.eventRetrieveCount &&  prevProps.eventRetrieveCount !== undefined) {
      this._getCalendar();
    }
    if (prevProps.fileRetrieveCount !== this.props.fileRetrieveCount &&  prevProps.fileRetrieveCount !== undefined) {
      this._getRecentFiles();
    }
  }


  private _getMail = (): void => {
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3) => {
        client
          .api(`/me/mailFolders/inbox/messages?$top=${this.props.mailRetrieveCount}`)
          .version("v1.0")
          .get((err: SPHttpClientResponse, res: SPHttpClientResponse) => {
            if (err) {
              console.log(err);
              return;
            }
            this.setState({ mail: res });
          });
      })
      .catch((err: SPHttpClientResponse) => {
        console.log(err);
      });
  };

  private _getCalendar = (): void => {
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3) => {
        client
          .api(`/me/calendar/events?$top=${this.props.eventRetrieveCount}`)
          .version("v1.0")
          .get((err: SPHttpClientResponse, res: SPHttpClientResponse) => {
            if (err) {
              console.log(err);
              return;
            }
            console.log(res)
            this.setState({ calendar: res });
          });
      })
      .catch((err: SPHttpClientResponse) => {
        console.log(err);
      });
  };

  private _getRecentFiles = (): void => {
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3) => {
        client
          .api(`/me/insights/used?$filter=resourceVisualization/type ne 'spsite' and NOT (resourceVisualization/type eq 'Web')&$orderby=lastUsed/lastAccessedDateTime desc&$top=${this.props.fileRetrieveCount}`)
          .version("v1.0")
          .get((err: SPHttpClientResponse, res: SPHttpClientResponse) => {
            if (err) {
              console.log(err);
              return;
            }
            this.setState({ files: res });
            console.log(res);
          });
      })
      .catch((err: SPHttpClientResponse) => {
        console.log(err);
      });
  };
 /*  private _getTodoLists = (): void => {
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
            this.setState({ todoLists: res });
            console.log(res);
          });
      })
      .catch((err: SPHttpClientResponse) => {
        console.log(err);
      });
  }; */

  public render(): React.ReactElement<IWebpartPersonalViewProps> {
    const {
      hasTeamsContext,
      userDisplayName,
      greetingPrefix,
      greetingSuffix,
      greetingShowUser,
      subGreeting,
      showGreeting
    } = this.props;

    const mailIconProps = { iconName: "Mail" };
    const calendarIconProps = { iconName: "Calendar" };
    const filesIconProps = { iconName: "OneDriveLogo" };
    const todoIconProps = { iconName: "ToDoLogoOutline" };


    return (
      <>
        <section
          className={`${styles.webpartPersonalView} ${
            hasTeamsContext ? styles.teams : ""
          }`}
        >
          {showGreeting && (
          <div className={styles.welcome}>
            <Text variant="xLarge">{greetingPrefix} {greetingShowUser ? userDisplayName : ""}{greetingSuffix}</Text><br />
            <Text variant="medium">{subGreeting}</Text>
          </div>
          )}
          <div className={styles.menu_grid}>
            <DefaultButton iconProps={mailIconProps} primary={this.state.activeView === "mail" ? true : false} text="Mail" onClick={() => this._changeActiveView("mail")}  />
            <DefaultButton iconProps={calendarIconProps} primary={this.state.activeView === "calendar" ? true : false} text="Calendar" onClick={() => this._changeActiveView("calendar")} />
            <DefaultButton iconProps={filesIconProps} primary={this.state.activeView === "files" ? true : false} text="Files" onClick={() => this._changeActiveView("files")} />
            <DefaultButton iconProps={todoIconProps} disabled={true} primary={this.state.activeView === "todo" ? true : false} text="Todo" onClick={() => this._changeActiveView("todo")} />
          </div>
        </section>

        <section
          className={`${styles.webpartPersonalView} ${ hasTeamsContext ? styles.teams : ""}`}
        >
          {this.state.mail && this.state.activeView === "mail" &&
            this.state.mail.value.map((mailItem: any) => {
              return <MailItem key={mailItem.id} mailItem={mailItem} />;
            })}
        </section>
        <section
          className={`${styles.webpartPersonalView} ${styles.calendar}  ${
            hasTeamsContext ? styles.teams : ""
          }`}
        >
          {this.state.calendar && this.state.activeView === "calendar" &&
            this.state.calendar.value.map((calendarItem: any) => {
              return (
                <CalendarItem key={calendarItem.id} event={calendarItem} />
              );
            })}
        </section>
        <section
          className={`${styles.webpartPersonalView} ${styles.filesection} ${
            hasTeamsContext ? styles.teams : ""
          }`}
        >
          {this.state.files && this.state.activeView === "files" &&
            this.state.files.value.map((fileItem: any) => {
              return <DriveItem key={fileItem.id} file={fileItem} />;
            })}
            {!this.state.files && this.state.activeView === "files" &&
            <div className={styles.container}>
              <Text variant="medium">No files found</Text>
            </div>
            }
        </section>
        <section
          className={`${styles.webpartPersonalView} ${
            hasTeamsContext ? styles.teams : ""
          }`}
        >
          {this.state.todoLists && this.state.activeView === "todo" &&
            this.state.todoLists.value.map((todoItem: any) => {
              return <TodoItem key={todoItem.id} task={todoItem} />;
            })}
            {!this.state.todoLists && this.state.activeView === "todo" &&
            <div className={styles.container}>
            <Text variant="medium">No tasks found</Text>
            </div>}
            
        </section>
      </>
    );
  }
}
