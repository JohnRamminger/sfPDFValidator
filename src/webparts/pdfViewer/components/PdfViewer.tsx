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
import { escape, fromPairs, unescape } from "@microsoft/sp-lodash-subset";
import { IPdfViewerReactState } from "./IPdfViewerReactState";

import "./docValidator.css";

import {
  PdfViewerComponent,
  Toolbar,
  Magnification,
  Navigation,
  LinkAnnotation,
  BookmarkView,
  ThumbnailView,
  Print,
  TextSelection,
  Annotation,
  TextSearch,
  Inject,
} from "@syncfusion/ej2-react-pdfviewer";
import { mergeStyles } from "office-ui-fabric-react/lib/Styling";
import {
  SuggestionsHeaderFooterItem,
  ThemeProvider,
} from "office-ui-fabric-react";
import { PopupWindowPosition } from "@microsoft/sp-property-pane";

interface viewerDoc {
  uri: string;
}

export default class PdfViewer extends React.Component<
  IPdfViewerProps,
  IPdfViewerReactState
> {
  public viewer: PdfViewerComponent;
  public base64data;
  constructor(props: IPdfViewerProps) {
    super(props);

    sp.setup({
      spfxContext: props.ctx,
    });

    let documentUrl: string = this.getUrlDoc();
    console.log(documentUrl);

    let npEnabled: boolean = false;
    if (!documentUrl || documentUrl == "") {
      npEnabled = true;
    }
    npEnabled = true;

    let bValidated: boolean = true;
    // alert(this.props.validationMode);
    this.state = {
      currentPage: 1,
      pageCount: 0,
      pagesRead: [1],
      validated: bValidated,
      docUrl: documentUrl,
      navPaneEnabled: npEnabled,
      showValidation: true,
      loading: true,
      validDocument: true,
    };
  }

  public componentDidMount() {
    this.TaskCompleted().then((result) => {
      debugger;
      let bValidated: boolean = result == "Complete";
      let bDocValid: boolean = result != "Not Assigned";
      this.setState({
        validated: bValidated,
        loading: false,
        validDocument: bDocValid,
      });
    });
  }

  public GetHeaderItems = (): JSX.Element[] => {
    let headerItems: JSX.Element[] = [];
    let icnName = "";
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
  };

  public render(): React.ReactElement<IPdfViewerProps> {
    let navObj = { css: { nextPageBtn: "nextPage" } };
    let nextDisabled = this.state.currentPage == this.state.pageCount;
    let prevDisabled = this.state.currentPage < 2;

    if (this.state.loading) {
      return <h1>Loading....</h1>;
    }

    let headerItems: JSX.Element[] = this.GetHeaderItems();
    let bodyItems: JSX.Element[] = [];
    let pdfContainer: JSX.Element[] = [];

    if (this.state.docUrl) {
      pdfContainer.push();
      pdfContainer.push();
      pdfContainer.push(<div className={styles.navHeader}>{headerItems}</div>);
    }
    const docs: viewerDoc[] = [{ uri: this.getUrlDoc() }];
    debugger;

    if (this.state.validated) {
      return (
        <h1 className={styles.textHeader}>{this.props.taskCompleteMessage}</h1>
      );
    }

    return (
      <div className={styles.pdfViewer}>
        <div>
          <div
            className={styles.headerText}
            dangerouslySetInnerHTML={{ __html: this.props.headerMessage }}
          />
          <div className={styles.navHeader}>{headerItems}</div>
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

  private getUrlDoc(): string {
    let queryParms = new UrlQueryParameterCollection(window.location.href);
    let docUrl: string = queryParms.getValue("docUrl");
    console.log("DocUrl: " + docUrl);
    return docUrl;
  }

  private getTaskList(): string {
    let queryParms = new UrlQueryParameterCollection(window.location.href);
    let docUrl: string = queryParms.getValue("TaskList");
    return docUrl;
  }

  private getTaskID(): string {
    let queryParms = new UrlQueryParameterCollection(window.location.href);
    let docUrl: string = queryParms.getValue("TaskID");
    return docUrl;
  }

  private ValidateDocument = async () => {
    debugger;
    try {
      let oWeb = Web(this.props.ctx.pageContext.web.absoluteUrl);
      let taskList = this.getTaskList();
      let lstTasks = await oWeb.lists.getByTitle(taskList);
      debugger;
      let itemID: number = parseInt(this.getTaskID());
      debugger;
      let taskItem = await lstTasks.items.getById(itemID);
      taskItem.update({
        rrReadingStatus: "Complete",
        rrDateCompleted: new Date(),
      });
      this.setState({ validated: true });
      if (this.props.taskCompleteMessage != undefined) {
        alert(this.props.taskCompleteMessage);
      } else {
        alert("Reading Task Completed");
      }
      let redirect: string = this.props.ctx.pageContext.web.absoluteUrl;
      if (this.props.redirectUrl != undefined && this.props.redirectUrl != "") {
        redirect = this.props.redirectUrl;
      }
      window.location.href = redirect;
    } catch (error) {}

    // tslint:disable-next-line
  };

  private TaskCompleted = async (): Promise<string> => {
    let retVal: string = "DocOnly";
    let itemID: number = parseInt(this.getTaskID());
    debugger;
    if (!isNaN(itemID)) {
      let oWeb = Web(this.props.ctx.pageContext.web.absoluteUrl);
      let lstTasks = await oWeb.lists.getByTitle(this.getTaskList());

      let taskItem = await lstTasks.items
        .getById(itemID)
        .select("rrReadingStatus,rrRequiredReadingUserEmail")
        .get()
        .then((result) => {
          retVal = result.dvValidationStatus;
          if (result.rrRequiredReadingUserEmail) {
            let userEmail: string = result.rrRequiredReadingUserEmail;
            userEmail = userEmail.toLocaleLowerCase();
            if (userEmail != this.props.ctx.pageContext.user.email) {
              retVal = "Not Assigned";
            }
          }
        });
    }
    return new Promise<string>(
      (resolve: (retVal: string) => void, reject: (error: Error) => void) => {
        resolve(retVal);
      }
    );

    // tslint:disable-next-line
  };
}
