//dropZone
import Dropzone from 'react-dropzone'
import {
  IDragDropEvents,
  IDragDropContext
} from 'office-ui-fabric-react/lib/utilities/dragdrop/interfaces';
export let _userDragFileByEvent:any=[];
export let _currentFolderByEvent="";
//used to access file from ExecDocuments
import execDocu, { _userDragFile, _currentFolder } from './ExecDocuments'
//end
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

import {
  Breadcrumb,
  IBreadcrumbItem
} from 'office-ui-fabric-react/lib/Breadcrumb'

import { Link } from 'office-ui-fabric-react/lib/Link'
import { Icon } from 'office-ui-fabric-react/lib/Icon'

import {
  IFSObject,
  redirectToSitePage,
  uploadFile,
  renameFile,
  siteDomain,
  siteUrl,
  shareFile,
  addFile,
  cdnAssetsBaseUrl
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

import styles from './EventDocuments.module.scss'
import { autobind } from '@uifabric/utilities';
import { FileAddResult, ItemUpdateResult} from 'sp-pnp-js/lib/sharepoint';
import { FolderUpdateResult } from 'sp-pnp-js/lib/sharepoint/folders'
import { find, debounce } from 'lodash';
import { css } from 'office-ui-fabric-react/lib/Utilities';

///////////////////////| Breadcrumbs |///////////////////////

interface IBreadCrumbsProps {
  items: IBreadcrumbItem[]
  canShare: boolean
  onClick: (item: IBreadcrumbItem, i: number) => void
  onShare: (item: IBreadcrumbItem) => void
}

class Breadcrumbs extends React.PureComponent<IBreadCrumbsProps> {
  public render() {
    const { items, onClick, canShare, onShare } = this.props

    return (
      <div className={styles.documentsBreadcrumb}>
        <Breadcrumb
          items={items.map((item, i) => {
            const breadcrumb: IBreadcrumbItem = {
              ...item,
              key: "crumb" + (i + 1),
              isCurrentItem: i === items.length - 1
            }
            // Breadcrumbs at the end can't be clicked
            if (i !== items.length - 1) {
              breadcrumb.onClick = () => onClick(item, i)
            }
            return breadcrumb
          })}
          onRenderItem={(item, defaultRender) => {
            if (canShare && (item as any).isSharedFolder) {
              return (
                <div className={styles.shareBreadcrumb}>
                  {defaultRender(item)}
                  <ActionButton text="Collaborate"
                    iconProps={{ iconName: "People" }}
                    className={styles.shareBreadcrumbBtn}
                    onClick={() => onShare(item)}
                    />
                </div>
              )
            } else {
              return defaultRender(item)
            }
          }}          
          />
      </div>
    )
  }
}

///////////////////////| EventDocuments |///////////////////////

export interface IEventDocumentsProps {
  items: IFSObject[]
  canShare: boolean
  rootFolder: string
  onNewFolderClick?: (currentFolder: string) => void
  onUploadFileFromComputerClick?: (currentFolder: string) => void
  onSyncToOneDriveClick?: (currentFolder: string) => void
  onDownloadFileClick?: (item: IFSObject) => void
  onOpenFileInBrowserClick?: (item: IFSObject) => void
  onRenameFileClick?: (item: IFSObject) => void
  onDuplicateFileClick?: (item: IFSObject) => void
  onDeleteFileClick?: (item: IFSObject) => void
  onShareIconClick?: (item: IFSObject) => void
  onShareEventClick?: (item: IFSObject) => void
  onFileUpload?: (currentFolder: string) => void  
}

export interface IEventDocumentsState {
  sort: ISortableList
  rowContextMenuMouseEvent: MouseEvent
  backgroundContextMenuMouseEvent: MouseEvent
  breadcrumbs: any[]
  currentFolder: string
  columns: IColumn[]
  activeItem: IFSObject
  loading: boolean
  //dropZone
  listStyle:boolean
}

export interface ISortableList {
  index: number
  asc: boolean
}

export default class EventDocuments extends React.Component<IEventDocumentsProps, IEventDocumentsState> {
  constructor(props: IEventDocumentsProps) {
    super(props)

    this.state = {
      //dropZone
      listStyle:false,
      //end
      loading: false,
      activeItem: null,
      sort: {
        index: 1,
        asc: true
      },
      rowContextMenuMouseEvent: null,
      backgroundContextMenuMouseEvent: null,
      breadcrumbs: [
        { text: "Shared Documents", serverRelativeUrl: props.rootFolder, isSharedFolder: true }
      ],
      currentFolder: props.rootFolder,
      columns: [
        {
          key: "column0",
          fieldName: "type",
          name: "Type",
          minWidth: 32,
          maxWidth: 32,
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
        }
      ]
    }

    // Initial sort UI
    const { columns, sort } = this.state
    const column = columns[sort.index]
    column.isSorted = true
    column.isSortedDescending = !sort.asc
  }

  public componentDidMount() {
    window.addEventListener("resize", debounce(() => {
      this.forceUpdate()
    }, 250))
  }

  private sortColumn(index: number) {
    const { sort: prevSort, columns } = this.state
    // Default sorting order when column is first clicked
    const nextSort: ISortableList = { index, asc: true }
    
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
      delete prevCol.isSorted
      delete prevCol.isSortedDescending
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

  private onFolderClick(item) {
    const { breadcrumbs } = this.state
    
    this.setState({
      breadcrumbs: [
        ...breadcrumbs,
        {
          text: item.name,
          serverRelativeUrl: item.serverRelativeUrl,
          isSharedFolder: this.isSharedFolderLevel()
        }
      ],
      currentFolder: item.serverRelativeUrl
    })
  }

  @autobind
  private isSharedFolderLevel(): boolean {
    return this.state.currentFolder.split("/").length === 5
  }

  @autobind
  private renderTypeIcon(item: IFSObject) {
    if (item.type === "file") {
      const extIconUrl = getOfficeUiFabricFileIconUrl(item.name)
      if (extIconUrl) {
        const ext = getFileExt(item.name)
        return (
          <img className={styles.fileIcon}
            src={extIconUrl}
            alt={ext} />
        )
      }
    }
    if (item.type === "folder") {
      return (
        <Icon className={styles.folderIcon}
          iconName="FabricFolderFill" />
      )
    }
    return null
  }

  private renderName(item: IFSObject) {
    if (item.type === "folder") {
      return (
        <Link onClick={() => this.onFolderClick(item)}>
          {item.name}
        </Link>
      )
    }
    if (item.type === "file") {
      const ext = getFileExt(item.name)
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

      const href = prefix + siteDomain + item.serverRelativeUrl
      return (
        <Link href={href} target={target}>
          {item.name}
        </Link>
      )
    }
    return null
  }

  @autobind
  private onNewFolderClick() {
    if (this.props.onNewFolderClick) {
      this.props.onNewFolderClick(this.state.currentFolder)
    }
  }

  @autobind
  private onUploadFileFromComputerClick() {
    if (this.props.onUploadFileFromComputerClick) {
      this.props.onUploadFileFromComputerClick(this.state.currentFolder)
    }
  }

  @autobind
  private onSyncToOneDriveClick() {
    if (this.props.onSyncToOneDriveClick) {
      this.props.onSyncToOneDriveClick(this.state.currentFolder);
    }
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

  @autobind
  private onRenameFileClick() {
    if (this.props.onRenameFileClick) {
      this.props.onRenameFileClick(this.state.activeItem)
    }
  }

  @autobind
  private onDuplicateFileClick() {
    if (this.props.onDuplicateFileClick) {
      this.props.onDuplicateFileClick(this.state.activeItem)
    }
  }

  @autobind
  private onDeleteFileClick() {
    if (this.props.onDeleteFileClick) {
      this.props.onDeleteFileClick(this.state.activeItem)
    }
  }

  @autobind
  private onShareIconClick(item: IFSObject) {
    if (this.props.onShareIconClick) {
      this.props.onShareIconClick(item)
    }
  }

  @autobind
  private onShareEventClick(breadcrumb: IBreadcrumbItem) {
    if (this.props.onShareEventClick) {
      const fsObject = find(this.props.items, x => x.name === "Shared")
      this.props.onShareEventClick(fsObject)
    }
  }

  @autobind
  private uploadFiles(files: FileList) {
    if (files){
      this.setState({ 
        backgroundContextMenuMouseEvent: null,
      })
      let uploads = []
      for(let i=0; i<files.length; i++) {
        uploads.push(uploadFile(this.state.currentFolder, files.item(i).name, files.item(i)))
      }
      Promise.all(uploads).then( () => {
        this.props.onFileUpload(this.state.currentFolder)
      })
    }
  }

  @autobind
  private renderFileInput(item: any, onDismiss: () => void ) {
    return (
      <div className={styles.contextualMenuItemContainer}>
        <div className={styles.contextualMenuUploadInput}>
          <input 
            type="file" 
            onChange={evt => {this.uploadFiles(evt.target.files)}}
            className={styles.contextualMenuFileselectItem}
            multiple
            data-is-focusable={false}
          />
        </div>
        <button 
          className={styles.contextualMenuUploaderButton} 
          onClick={() => {return null}}
        >
          <span>Upload from Computer</span>
        </button>
      </div>
    )
  }
  //dropzone
  private onDrop(uploadUserFile) {
    let uploads = []
    this.setState({listStyle:false})
    let _uploadFile=_userDragFile.length>0?_userDragFile:uploadUserFile; 
    var customCurrentFolder=this.state.currentFolder?this.state.currentFolder:_currentFolder;
    if(_userDragFile.length>0)
    {
      uploads.push(shareFile(customCurrentFolder, _uploadFile[0]))
    }
    else
    {
      for(let i=0; i<_uploadFile.length; i++) {
        let Extension=_uploadFile[i].name.split(/\.(?=[^\.]+$)/);
        if(Extension[1]!="url")
        uploads.push(addFile(customCurrentFolder, _uploadFile[i].name,_uploadFile[i]))
      }     
    }
    Promise.all(uploads).then( () => {
      this.props.onFileUpload(customCurrentFolder);
    })
    
  }
  private _getDragDropEvents(e): IDragDropEvents {
    return {
    
      canDrop: (dropContext?: IDragDropContext, dragContext?: IDragDropContext) => {
        
          return true; 
        
      },
      canDrag: (item1?: any) => {
        if(item1.type!="folder"){
            return true;
        }
        else
            return false;
        },
      
      onDragStart: (item1?: any, itemIndex?: number, selectedItems?: any[], event?: MouseEvent) => {
       
        _userDragFileByEvent=selectedItems;
        _currentFolderByEvent= this.state.currentFolder;
        return true;
      },
      onDragEnd: (item1?: any, event?: DragEvent) => {
        //_draggedItemByEvent = null;
        _userDragFileByEvent=[];
      },
     
    };
  }
  //Active dropzone
  private setActive()
  {
    this.setState({listStyle:true})
  }
  //deactivate dropzone
  private setDeactivate()
  {
    this.setState({listStyle:false})
  }
  //end
  public render() {
    const {
      sort,
      columns,
      breadcrumbs,
      currentFolder,
      rowContextMenuMouseEvent,
      backgroundContextMenuMouseEvent
    } = this.state
    
    let items = this.props.items.slice()
    // Filter items to show only items in current folder
    const currentFolderPaths = currentFolder.split("/")
    items = items.filter(item => {
      const paths = item.serverRelativeUrl.split("/")
      return (
        paths.length === currentFolderPaths.length + 1 &&
        paths[paths.length - 2] === currentFolderPaths[currentFolderPaths.length - 1]
      )
    })

    const { fieldName } = columns[sort.index]

    // Sorting
    if (fieldName === "modifiedDate") {
      sortMomentsBy(items, fieldName, sort.asc ? "asc" : "desc")
    } else {
      sortBy(items, fieldName, sort.asc ? "asc" : "desc")
    }

    // Rendering
    return (
      <div>
        <Breadcrumbs
          items={breadcrumbs}
          onClick={(item: IBreadcrumbItem, index) => {
            for (let i = breadcrumbs.length - 1; i > index; i -= 1) {
              breadcrumbs.splice(i, 1)
            }
            const lastBreadcrumb = breadcrumbs[breadcrumbs.length - 1]
            const newFolder = (lastBreadcrumb as any).serverRelativeUrl
            
            this.setState({
              breadcrumbs,
              currentFolder: newFolder
            })
          }}
          canShare={this.props.canShare}
          onShare={this.onShareEventClick}
        />
        {/* add dropzone for file upload. */}
        <Dropzone onDragOver={this.setActive.bind(this)} onDragLeave={this.setDeactivate.bind(this)}
        onDrop={this.onDrop.bind(this)}  style={styles.dropzoneStyle} disableClick={true} activeClassName={styles.dropzoneStyleActive}>
         
        <div className={styles.DetailsList}
          ref={this.detailsListContainerRef.bind(this)}
          onContextMenu={(evt) => {
            evt.preventDefault()
            // We only want to show the menu when the target
            // is the DetailsList div and we're not at the top
            // level (because this menu isn't allowed)
            if ((evt.nativeEvent.target as any).className.indexOf("DetailsList") === 0) {
              this.setState({ backgroundContextMenuMouseEvent: evt.nativeEvent })
            }
          }}>
          {/* console.log(items, 'WHAT ARE ITEMS') */}
          <DetailsList
            dragDropEvents={ this._getDragDropEvents(this) }
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
            }} />
          {rowContextMenuMouseEvent && 
            (this.state.activeItem.type == 'file') && (
            <ContextualMenu
              items={[
                { key: "1", name: "Download" },
                { key: "2", name: "Open in Browser" },
                { key: "3", itemType: ContextualMenuItemType.Divider },
                { key: "4", name: "Rename" },
                { key: "5", name: "Duplicate" },
                { key: "6", itemType: ContextualMenuItemType.Divider },
                { key: "7", name: "Delete" },
              ]}
              target={rowContextMenuMouseEvent}
              onItemClick={(evt, item: IContextualMenuItem) => {
                switch (item.key) {
                  case "1": this.onDownloadFileClick(); break
                  case "2": this.onOpenFileInBrowserClick(); break
                  case "4": this.onRenameFileClick(); break
                  case "5": this.onDuplicateFileClick(); break
                  case "7": this.onDeleteFileClick(); break
                  default: break
                }
              }}
              onDismiss={() => {
                this.setState({ rowContextMenuMouseEvent: null })
              }} />
          )}
          {rowContextMenuMouseEvent && 
          (this.state.activeItem.type == 'folder') && (
            <ContextualMenu
              items={[
                { key: "1", name: "Rename" },
                { key: "2", name: "Delete" },
              ]}
              target={rowContextMenuMouseEvent}
              onItemClick={(evt, item: IContextualMenuItem) => {
                switch (item.key) {
                  case "1": this.onRenameFileClick(); break
                  case "2": this.onDeleteFileClick(); break
                  default: break
                }
              }}
              onDismiss={() => {
                this.setState({ rowContextMenuMouseEvent: null })
              }} />
          )}
          {backgroundContextMenuMouseEvent && (
            <ContextualMenu
            items={[
              { key: "1", name: "New Folder" },
              { key: "2", name: "Upload from Computer", onRender: this.renderFileInput},
              { key: "3", name: "Sync to OneDrive" },
            ]}
            target={backgroundContextMenuMouseEvent}
            onItemClick={(evt, item: IContextualMenuItem) => {
              switch (item.key) {
                case "1": this.onNewFolderClick(); break
                case "2": this.onUploadFileFromComputerClick(); break
                case "3": this.onSyncToOneDriveClick(); break
                default: break
              }
            }}
            onDismiss={() => {
              this.setState({ backgroundContextMenuMouseEvent: null })
            }} />        
          )}
        </div>
        
        </Dropzone>
        <div className={styles.dragZoneOperation}>
              { (this.state.listStyle)? <div className={styles.displayImg}> <img src={ `${cdnAssetsBaseUrl}/images/copyFile.png` } /><div>Copy</div></div> :""  }
        </div>
      </div>
    )
  }  
}
