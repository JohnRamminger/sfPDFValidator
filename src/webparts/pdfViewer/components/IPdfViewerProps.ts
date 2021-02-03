import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPdfViewerProps {
  ctx: WebPartContext;
  validationIcon: string;
  validationText: string;
  taskCompleteMessage: string;
  headerMessage: string;
  redirectUrl: string;
}
