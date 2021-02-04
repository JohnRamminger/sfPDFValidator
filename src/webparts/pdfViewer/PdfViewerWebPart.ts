import * as React from "react";
// import "@pnp/polyfill-ie11";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  IPropertyPaneDropdownCalloutProps,
  PropertyPaneDropdown,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "PdfViewerWebPartStrings";
import PdfViewer from "./components/PdfViewer";
import { IPdfViewerProps } from "./components/IPdfViewerProps";
import { MiscFunctions } from "../../services";

export interface IPdfViewerWebPartProps {
  validationIcon: string;
  validationText: string;
  taskCompleteMessage: string;
  headerMessage: string;
  redirectUrl: string;
}

export default class PdfViewerWebPart extends BaseClientSideWebPart<IPdfViewerWebPartProps> {
  public render(): void {
    if (MiscFunctions.IsInternetExplorer()) {
      this.domElement.innerHTML = "<h1>IE Not Suppored</h1>";
      return;
    }
    console.log("Got by the IE Check");
    const element: React.ReactElement<IPdfViewerProps> = React.createElement(
      PdfViewer,
      {
        ctx: this.context,
        validationIcon: this.properties.validationIcon,
        validationText: this.properties.validationText,
        taskCompleteMessage: this.properties.taskCompleteMessage,
        headerMessage: this.properties.headerMessage,
        redirectUrl: this.properties.redirectUrl,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("validationText", {
                  label: "Validation Text",
                }),
                PropertyPaneTextField("validationIcon", {
                  label: "Validation Icon",
                }),
                PropertyPaneTextField("taskCompleteMessage", {
                  label: "Task Complete Message",
                }),
                PropertyPaneTextField("headerMessage", {
                  label: "Header Message",
                }),
                PropertyPaneTextField("redirectUrl", {
                  label: "Url to redirect to when complete",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
