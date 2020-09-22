import * as React from 'react';
import styles from './PdfViewer.module.scss';
import { IPdfViewerProps } from './IPdfViewerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IPdfViewerReactState } from './IPdfViewerReactState';
import '../../../../node_modules/@syncfusion/ej2-base/styles/material.css';
import '../../../../node_modules/@syncfusion/ej2-buttons/styles/material.css';
import '../../../../node_modules/@syncfusion/ej2-dropdowns/styles/material.css';
import '../../../../node_modules/@syncfusion/ej2-inputs/styles/material.css';
import '../../../../node_modules/@syncfusion/ej2-navigations/styles/material.css';
import '../../../../node_modules/@syncfusion/ej2-popups/styles/material.css';
import '../../../../node_modules/@syncfusion/ej2-splitbuttons/styles/material.css';
import '../../../../node_modules/@syncfusion/ej2-react-pdfviewer/styles/material.css';
import '../../../../node_modules/@syncfusion/ej2-notifications/styles/material.css';

import { ToolbarComponent, ItemsDirective, ItemDirective, ClickEventArgs } from '@syncfusion/ej2-react-navigations';

import { PdfViewerComponent, Toolbar, Magnification, Navigation, LinkAnnotation, BookmarkView, ThumbnailView, Print, TextSelection, Annotation, TextSearch, Inject } from '@syncfusion/ej2-react-pdfviewer';

export default class PdfViewer extends React.Component<IPdfViewerProps, IPdfViewerReactState> {

  constructor(props: IPdfViewerProps) {
    super(props);
    this.state = { currentPage: 0, pageCount: 0 };
  }

  public render(): React.ReactElement<IPdfViewerProps> {
    return (
      <PdfViewerComponent documentLoad={this.docLoaded} pageChange={this.pageChange} id="container" documentPath="PDF_Succinctly.pdf" serviceUrl="https://ej2services.syncfusion.com/production/web-services/api/pdfviewer" style={{ 'height': '1000px' }}>
        <Inject services={[Toolbar, Magnification, Navigation, Annotation, LinkAnnotation, BookmarkView, ThumbnailView, Print, TextSelection, TextSearch]} />
      </PdfViewerComponent>
    );
  }

  private docLoaded = (args) => {
    let pd: string = args.pageData;
    pd = pd.substring(pd.indexOf(':') + 1, pd.indexOf(','));
    let IPageCount: number = parseInt(pd);
    this.setState({ pageCount: IPageCount, currentPage: 1 });
  }

  private pageChange = (args) => {
    this.setState({ currentPage: args.currentPageNumber });

  }

}
