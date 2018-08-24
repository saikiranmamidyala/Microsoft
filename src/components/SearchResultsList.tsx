import * as React from 'react'

import Loading from './Loading'

import {
  IColumn,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
  ConstrainMode
} from 'office-ui-fabric-react/lib/DetailsList'

import {
  ContextualMenu,
  ContextualMenuItemType,
  IContextualMenuItem
} from 'office-ui-fabric-react/lib/ContextualMenu'

import { Link } from 'office-ui-fabric-react/lib/Link'
import { Icon } from 'office-ui-fabric-react/lib/Icon'

import {
  IFSObject,
  redirectToSitePage,
  uploadFile,
  renameFile,
  siteDomain,
  siteUrl
} from '../shared/SharePoint'

import {
  humanFileSize,
  getFileExt,
  getOfficeUiFabricFileIconUrl,
  sortBy,
  sortMomentsBy
} from '../shared/util'

import {
  Button,
  ActionButton,
} from 'office-ui-fabric-react/lib/Button';

import styles from './SearchResultsList.module.scss'
import { autobind } from '@uifabric/utilities';
import { FileAddResult, ItemUpdateResult} from 'sp-pnp-js/lib/sharepoint';
import { FolderUpdateResult } from 'sp-pnp-js/lib/sharepoint/folders'
import { find, debounce } from 'lodash';
import { css } from 'office-ui-fabric-react/lib/Utilities';

export interface ISearchResultsProps {
  items: IFSObject[]
  rootFolder: string
  onDownloadFileClick?: (item: IFSObject) => void
  onOpenFileInBrowserClick?: (item: IFSObject) => void
}

export interface ISearchResultsState {
  sort: ISortableList
  rowContextMenuMouseEvent: MouseEvent
  currentFolder: string
  columns: IColumn[]
  activeItem: IFSObject
  loading: boolean
}

export interface ISortableList {
  index: number
  asc: boolean
}

export default class SearchResultsDocuments extends React.Component<ISearchResultsProps, ISearchResultsState> {
  constructor(props: ISearchResultsProps) {
    super(props)

    this.state = {
      loading: false,
      activeItem: null,
      sort: null, /* {
        index: 2,
        asc: true
      }, */
      rowContextMenuMouseEvent: null,
      currentFolder: props.rootFolder,
      columns: [
         {
          key: "column0",
          fieldName: "type",
          name: "Type",
          minWidth: 42,
          maxWidth: 42,
          isResizable: false,
          onRender: this.renderTypeIcon.bind(this)          
        }, {
          key: "column1",
          fieldName: "name",
          name: "Name",
          minWidth: 150,
          maxWidth: 500,
          isResizable: true,
          onColumnClick: () => this.sortColumn(1),
          onRender: this.renderName.bind(this)
        }, {
          key: "column2",
          fieldName: "modified",
          name: "Modified Date",
          minWidth: 100,
          maxWidth: 100,
          isResizable: false,
          onColumnClick: () => this.sortColumn(2),
          onRender: item => item.modified.format("YYYY-MM-DD"),
        }, {
          key: "column3",
          fieldName: "author",
          name: "Author",
          minWidth: 150,
          maxWidth: 150,
          isResizable: true,
          onColumnClick: () => this.sortColumn(3),
          onRender: item => item.author
        }, {
          key: "column4",
          fieldName: "size",
          name: "File Size",
          minWidth: 50,
          maxWidth: 50,
          isResizable: false,
          onColumnClick: () => this.sortColumn(4),
          onRender: item => item.size ? humanFileSize(item.size) : null
        }
      ]
    }

    // Initial sort UI
    const { columns, sort } = this.state
    if (sort) {
      const column = columns[sort.index]
      column.isSorted = true
      column.isSortedDescending = !sort.asc
    }
  }

  public componentDidMount() {
    window.addEventListener("resize", debounce(() => {
      this.forceUpdate()
    }, 250))
  }

  private sortColumn(index: number) {
    const { columns } = this.state
    let prevSort = this.state.sort
    // Default sorting order when column is first clicked
    const nextSort: ISortableList = { index, asc: true }
    
    if (!prevSort) {
      prevSort = {
        index: -1,
        asc: true
      }
    }

    // Toggle sorting order if column is clicked again
    if (prevSort.index === nextSort.index) {
      nextSort.asc = !prevSort.asc
      const col = columns[nextSort.index]
      // Toggle column arrow UI
      col.isSortedDescending = !nextSort.asc
    } else {
      // We get here when the user clicks a column header
      // that is different than the currently sorted column
      
      // Remove arrow UI from currently sorted column
      const prevCol = columns[prevSort.index]
      if (prevCol) {
        delete prevCol.isSorted
        delete prevCol.isSortedDescending
      }
      // Add arrow UI for the newly sorted column
      const nextCol = columns[nextSort.index]
      nextCol.isSorted = true
      nextCol.isSortedDescending = !nextSort.asc
    }

    // Re-render
    this.setState({
      sort: nextSort,
      columns
    })
  }

  private detailsListContainerRef(div: HTMLDivElement) {
    if (div) {
      const windowHeight = window.innerHeight
      const divTop = div.getBoundingClientRect().top
      const height = windowHeight - divTop
      div.style.minHeight = `${height}px`
    }
  }

  @autobind
  private renderTypeIcon(item: IFSObject) {
    if (item.type === "file") {
      const extIconUrl = getOfficeUiFabricFileIconUrl("." + item.fileExtension)
      if (extIconUrl) {
        return (
          <img className={styles.fileIcon}
            src={extIconUrl}
            alt={item.fileExtension} />
        )
      }
    }
    return null
  }

  private renderName(item: IFSObject) {
    if (item.type === "file") {
      const ext = item.fileExtension 
      let prefix = ""
      let target = ""

      if (ext === "pptx") {
        prefix = "ms-powerpoint:ofe%7Cu%7C"
      } else if (ext === "docx") {
        prefix = "ms-word:ofe%7Cu%7C"
      } else if (ext === "xslx") {
        prefix = "ms-excel:ofe%7Cu%7C"
      } else {
        target = "_blank"
      }

      const href = prefix + item.serverRelativeUrl

      return (
        <Link href={href} target={target}>
          {item.name}
        </Link>
      )
    }
    return null
  }

  @autobind
  private onDownloadFileClick() {
    if (this.props.onDownloadFileClick) {
      this.props.onDownloadFileClick(this.state.activeItem)
    }
  }

  @autobind
  private onOpenFileInBrowserClick() {
    if (this.props.onOpenFileInBrowserClick) {
      this.props.onOpenFileInBrowserClick(this.state.activeItem)
    }
  }

  public render() {
    const {
      sort,
      columns,
      currentFolder,
      rowContextMenuMouseEvent,
    } = this.state
    
    let items = this.props.items.slice()

    if (sort) {
      const { fieldName } = columns[sort.index]

      // Sorting
      if (fieldName === "modified") {
        sortMomentsBy(items, fieldName, sort.asc ? "asc" : "desc")
      } else {
        sortBy(items, fieldName, sort.asc ? "asc" : "desc")
      }
    }

    // Rendering
    return (
      <div>        
        <div className={styles.DetailsList}
          ref={this.detailsListContainerRef.bind(this)}
          onContextMenu={(evt) => {
            evt.preventDefault()
          }
        }>
          <DetailsList
            columns={this.state.columns}
            items={items}
            layoutMode={DetailsListLayoutMode.justified}
            constrainMode={ConstrainMode.unconstrained}
            checkboxVisibility={CheckboxVisibility.hidden}
            selectionMode={SelectionMode.none}
            onItemContextMenu={(item, i, evt: MouseEvent) => {
              this.setState({ 
                rowContextMenuMouseEvent: evt,
                activeItem: item
              })
            }
            } />
          {rowContextMenuMouseEvent && 
          (this.state.activeItem.type == 'file') && (
            <ContextualMenu
              items={[
                { key: "1", name: "Download" },
                { key: "2", name: "Open in Browser" },
              ]}
              target={rowContextMenuMouseEvent}
              onItemClick={(evt, item: IContextualMenuItem) => {
                switch (item.key) {
                  case "1": this.onDownloadFileClick(); break
                  case "2": this.onOpenFileInBrowserClick(); break
                  default: break
                }
              }}
              onDismiss={() => {
                this.setState({ rowContextMenuMouseEvent: null })
              }
            }/>
          )}
        </div>
      </div>
    )
  }  
}
