import * as React from 'react';
import styles from './followDocumentGrid.module.scss';

// Used to render list grid
import { FocusZone } from 'office-ui-fabric-react/lib/FocusZone';
import { List } from 'office-ui-fabric-react/lib/List';
import { IRectangle, ISize } from 'office-ui-fabric-react/lib/Utilities';

import { Person } from '@microsoft/mgt-react/dist/es6/spfx';
import { ViewType, PersonCardInteraction } from '@microsoft/mgt-spfx';

import {
  DocumentCard,
  DocumentCardImage,
  DocumentCardActions,
  ImageFit
} from 'office-ui-fabric-react';

import { IfollowDocumentGridProps, IfollowDocumentGridState } from './followDocumentGrid.types';

const ROWS_PER_PAGE: number = +styles.rowsPerPage;

export class FollowDocumentGrid extends React.Component<IfollowDocumentGridProps, IfollowDocumentGridState> {
  private _columnWidth: number;
  private _rowHeight: number;
  private _isCompact: boolean;

  public render(): React.ReactElement<IfollowDocumentGridProps> {

    const sortedItems = this.props.items ? this.props.items.sort((a,b) => a.Title.localeCompare(b.Title)) : [];

    return (
      <div role="group" aria-label={this.props.ariaLabel} >
      <FocusZone className={styles.documentCardCntnr}>
        
        {/* <List
          role="presentation"
          className={styles.followDocumentGrid}
          items={this.props.items}
          getItemCountForPage={this._getItemCountForPage}
          getPageHeight={this._getPageHeight}
          onRenderCell={this._onRenderCell}
          {...this.props.listProps}
        /> */}

        {sortedItems && sortedItems.map(item => {

          const docExt = item.Title.substring(item.Title.lastIndexOf('.')+1).toLowerCase();
          let docIcon = "";
          switch (docExt){
            case "aspx" :
              docIcon = 'spo';
              break;
            case "doc" :
              docIcon = 'docx';
              break;
            case "xls" :
              docIcon = 'xlsx';
              break;
            case "ppt" :
              docIcon = 'pptx';
              break;
            default:
              docIcon = 'spo';
              break;
          }

          return(
            <DocumentCard className={styles.documentCard} onClickHref={item.Url}>
              <div className={styles.docName}>
                <img width='18' height='18' src={`https://static2.sharepointonline.com/files/fabric/assets/item-types/16/${docIcon}.svg`} />
                <span>{item.Title.substring(0, item.Title.lastIndexOf('.'))}</span>
              </div>
              <DocumentCardImage height={150} imageFit={ImageFit.cover} imageSrc={item.Thumbnail} />
              <div className={styles.docDetails}>
                <Person 
                  showPresence 
                  personCardInteraction={ PersonCardInteraction.hover} 
                  personQuery={item.lastModifiedBy.user.email} 
                  avatarSize={'small'}
                  view={ViewType.oneline} />
                  <div className={styles.docModifiedDate}>{item.lastModifiedDate}</div>
              </div>
              <DocumentCardActions actions={item.documentCardActions} />
            </DocumentCard>
          );

        })}

      </FocusZone>
      </div>
    );
  }

  private _getItemCountForPage = (itemIndex: number, surfaceRect: IRectangle): number => {
    return ROWS_PER_PAGE;
  }

  private _getPageHeight = (): number => {
    return this._rowHeight * ROWS_PER_PAGE;
  }

  private _onRenderCell = (item: any, index: number | undefined): JSX.Element => {
    const isCompact: boolean = this._isCompact;
    const finalSize: ISize = { width: this._columnWidth, height: this._rowHeight };
    return (
      <div
        style={{
          width: "200px",
          marginRight: "20px"
        }}
      >
          {this.props.onRenderGridItem(item, finalSize, isCompact)}
      </div>
    );
  }
}