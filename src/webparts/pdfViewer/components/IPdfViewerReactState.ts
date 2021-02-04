export interface IPdfViewerReactState {
  validated: boolean;
  loading: boolean;
  validDocument: boolean;
  docUrl: string;
  showValidation: boolean;
  documentLibrary: string;
  taskList: string;
  docItemID: number;
  taskItemID: number;
  currentUserEmail: string;
}
