import {
  Field,
  FieldUserSelectionMode,
  IItemAddResult,
  sp,
  Web,
} from "@pnp/sp/presets/all";

import DocViewer, { DocViewerRenderers } from "react-doc-viewer";

import "@pnp/sp/webs";
import "@pnp/sp/files/web";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { UrlQueryParameterCollection } from "@microsoft/sp-core-library";

import * as React from "react";
import styles from "./PdfViewer.module.scss";
import { IPdfViewerProps } from "./IPdfViewerProps";
import { IPdfViewerReactState } from "./IPdfViewerReactState";

import "./docValidator.css";

import { mergeStyles } from "office-ui-fabric-react/lib/Styling";
import {
  SuggestionsHeaderFooterItem,
  ThemeProvider,
} from "office-ui-fabric-react";
import { PopupWindowPosition } from "@microsoft/sp-property-pane";
import { MiscFunctions, SPLogging } from "../../../services";

interface viewerDoc {
  uri: string;
}

export default class PdfViewer extends React.Component<
  IPdfViewerProps,
  IPdfViewerReactState
> {
  public base64data;
  constructor(props: IPdfViewerProps) {
    super(props);

    sp.setup({
      spfxContext: props.ctx,
    });

    let itemID: number = this.getItemID();
    let doclib: string = this.getDocLib();
    let taskID: number = this.getTaskID();
    let taskList: string = this.getTaskList();

    let bValidated: boolean = true;
    // alert(this.props.validationMode);
    this.state = {
      currentUserEmail: props.ctx.pageContext.user.email.toLowerCase(),
      validated: bValidated,
      docUrl: "",
      showValidation: true,
      loading: true,
      validDocument: true,
      documentLibrary: doclib,
      taskList: taskList,
      docItemID: itemID,
      taskItemID: taskID,
    };
  }

  public async componentDidMount() {
    if (this.isValid()) {
      await this.GetDocumentUrl();
      await this.TaskCompleted().then((result) => {
        debugger;
        let bValidated: boolean = result == "DocOnly";
        let bDocValid: boolean = result != "Not Assigned";
        this.setState({
          validated: bValidated,
          loading: false,
          validDocument: bDocValid,
        });
      });
    }
  }

  public GetHeaderItems = (): JSX.Element[] => {
    debugger;
    if (this.state.validated) {
      return [<div></div>];
    }
    let headerItems: JSX.Element[] = [];
    let icnName = "Accept";
    if (this.props.validationIcon != undefined) {
      icnName = this.props.validationIcon;
    }
    // if (!this.state.validDocument && this.state.validDocument) {
    headerItems.push(
      <Icon
        iconName={icnName}
        className={styles.signIcon}
        title={this.props.validationText}
        onClick={this.ValidateDocument}
      />
    );
    // }
    return headerItems;
    // tslint:disable-next-line
  };

  public render(): React.ReactElement<IPdfViewerProps> {
    if (MiscFunctions.IsInternetExplorer()) {
      return (
        <h1>
          This application does not support Internet Explorer. Please cut and
          paste the URL in to Edge, Chrome, or FireFox.
        </h1>
      );
    }
    if (!this.isValid()) {
      return (
        <h3 title="The Proper Parameters have not be specified on the Url">
          Unable to Load Document.
        </h3>
      );
    }

    let headerItems: JSX.Element[] = this.GetHeaderItems();
    let bodyItems: JSX.Element[] = [];
    let pdfContainer: JSX.Element[] = [];

    debugger;
    if (!MiscFunctions.IsEmpty(this.state.docUrl)) {
      pdfContainer.push(<div className={styles.navHeader}>{headerItems}</div>);
    }
    if (this.state.loading || MiscFunctions.IsEmpty(this.state.docUrl)) {
      return <h1>Loading....</h1>;
    }

    const docs: viewerDoc[] = [{ uri: this.state.docUrl }];
    debugger;

    let taskCompleteMessage: string = "Task is Complete!";
    if (!MiscFunctions.IsEmpty(this.props.taskCompleteMessage)) {
      taskCompleteMessage = this.props.taskCompleteMessage;
    }
    return (
      <div className={styles.pdfViewer}>
        <div>
          {this.state.validated && <h2>{taskCompleteMessage}</h2>}
          {!this.state.validated && (
            <div>
              <div
                className={styles.headerText}
                dangerouslySetInnerHTML={{ __html: this.props.headerMessage }}
              />
              <div className={styles.navHeader}>{headerItems}</div>
            </div>
          )}
          <DocViewer
            pluginRenderers={DocViewerRenderers}
            documents={docs}
            config={{
              header: {
                disableHeader: true,
                disableFileName: true,
                retainURLParams: false,
              },
            }}
          />
          <div className={styles.navHeader}>{headerItems}</div>
        </div>
        {/*     });
         */}
      </div>
    );
  }

  private getDocLib(): string {
    let queryParms = new URLSearchParams(window.location.href);

    let docUrl: string = queryParms.get("doclib");
    console.log("DocLib: " + docUrl);
    return docUrl;
  }

  private getItemID(): number {
    let queryParms = new URLSearchParams(window.location.href);
    let itemID: number = parseInt(queryParms.get("ItemID"));
    console.log("ItemID: " + itemID);
    return itemID;
  }
  private isValid(): boolean {
    if (
      this.state.docItemID > 0 &&
      this.state.taskItemID > 0 &&
      !MiscFunctions.IsEmpty(this.state.documentLibrary) &&
      !MiscFunctions.IsEmpty(this.state.taskList)
    ) {
      return true;
    } else {
      return false;
    }
  }

  private getTaskList(): string {
    let queryParms = new URLSearchParams(window.location.href);
    let docUrl: string = queryParms.get("TaskList");
    return docUrl;
  }

  private getTaskID(): number {
    let queryParms = new URLSearchParams(window.location.href);
    let itemID: number = parseInt(queryParms.get("TaskID"));
    return itemID;
  }

  private ValidateDocument = async () => {
    debugger;
    try {
      debugger;
      let result: boolean = await this.updateTask(
        this.state.taskList,
        this.state.taskItemID,
        {
          rrReadingStatus: "Complete",
          rrDateCompleted: new Date(),
        }
      );
      if (result == true) {
        this.setState({ validated: true });
        if (this.props.taskCompleteMessage != undefined) {
          alert(this.props.taskCompleteMessage);
        } else {
          alert("Reading Task Completed");
        }
        let redirect: string = this.props.ctx.pageContext.web.absoluteUrl;
        if (
          this.props.redirectUrl != undefined &&
          this.props.redirectUrl != ""
        ) {
          redirect = this.props.redirectUrl;
        }
        window.location.href = redirect;
      } else {
        SPLogging.LogError(
          "Validate Document",
          "TaskItem:" +
            this.state.taskItemID +
            " - TaskList:" +
            this.state.taskList
        );
      }
    } catch (error) {
      SPLogging.LogError("Validate Document", error.message);
    }

    // tslint:disable-next-line
  };

  private updateTask(
    listName: string,
    itemID: number,
    updates: any
  ): Promise<boolean> {
    let retVal: boolean = true;
    try {
      debugger;
      let oWeb = Web(this.props.ctx.pageContext.web.absoluteUrl);
      oWeb.lists.getByTitle(listName).items.getById(itemID).update(updates);
    } catch (error) {
      debugger;
      retVal = false;
      SPLogging.LogError("updateTask", error.message);
    }
    return new Promise<boolean>(
      (resolve: (retVal: boolean) => void, reject: (error: Error) => void) => {
        resolve(retVal);
      }
    );
  }

  private GetDocumentUrl = async () => {
    let oWeb = Web(this.props.ctx.pageContext.web.absoluteUrl);
    oWeb.lists
      .getByTitle(this.state.documentLibrary)
      .items.getById(this.state.docItemID)
      .select("FileRef")
      .get()
      .then((result) => {
        this.setState({ docUrl: result.FileRef });
      });
  };

  private TaskCompleted = async (): Promise<string> => {
    let retVal: string = "Not Validated";

    debugger;
    try {
      if (!isNaN(this.state.taskItemID)) {
        let oWeb = Web(this.props.ctx.pageContext.web.absoluteUrl);
        let lstTasks = await oWeb.lists.getByTitle(this.state.taskList);
        let taskItem = await lstTasks.items
          .getById(this.state.taskItemID)
          .select("rrReadingStatus,rrRequiredReadingUserEmail")
          .get()
          .then(async (result) => {
            debugger;
            retVal = result.dvValidationStatus;
            if (result.rrRequiredReadingUserEmail) {
              let userEmail: string = result.rrRequiredReadingUserEmail;
              userEmail = userEmail.toLocaleLowerCase();
              if (result.rrReadingStatus == "Not Started") {
                await this.updateTask(
                  this.state.taskList,
                  this.state.taskItemID,
                  {
                    rrReadingStatus: "In Progress",
                    rrDateRead: new Date(),
                  }
                );
              }
              if (userEmail != this.state.currentUserEmail) {
                retVal = "Not Assigned";
              } else {
                if (result.rrReadingStatus == "Complete") {
                  retVal = "DocOnly";
                }
              }
            }
          });
      }
    } catch (error) {
      SPLogging.LogError("TaskCompleted", error.message);
    }
    return new Promise<string>(
      (resolve: (retVal: string) => void, reject: (error: Error) => void) => {
        resolve(retVal);
      }
    );
    // tslint:disable-next-line
  };
}
