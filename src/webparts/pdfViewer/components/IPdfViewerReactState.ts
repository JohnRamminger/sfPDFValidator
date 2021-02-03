export interface IPdfViewerReactState {
  currentPage: number;
  pageCount: number;
  validated: boolean;
  loading: boolean;
  validDocument: boolean;
  pagesRead: number[];
  navPaneEnabled: boolean;
  docUrl: string;
  showValidation: boolean;
}
