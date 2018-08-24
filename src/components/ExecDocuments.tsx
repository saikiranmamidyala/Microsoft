//dropZone
import Dropzone from 'react-dropzone'

import {
  IDragDropEvents,
  IDragDropContext
} from 'office-ui-fabric-react/lib/utilities/dragdrop/interfaces';
let _draggedItem: any = null;
let _draggedIndex = -1;
export let _userDragFile:any=[];
export let _currentFolder="";
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { _userDragFileByEvent, _currentFolderByEvent } from './EventDocuments'
//end
import * as React from 'react'

import Loading from './Loading'

import {
  IColumn,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
  ConstrainMode,
  //dropZone
  Selection
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
  moveFile,
  isFileNameExists,
  addFile,
  cdnAssetsBaseUrl,
  shareFile
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

import styles from './ExecDocuments.module.scss'
import { autobind } from '@uifabric/utilities';
import { FileAddResult, ItemUpdateResult} from 'sp-pnp-js/lib/sharepoint';
import { FolderUpdateResult } from 'sp-pnp-js/lib/sharepoint/folders'
import { find, debounce } from 'lodash';
import { css } from 'office-ui-fabric-react/lib/Utilities';

///////////////////////| Breadcrumbs |///////////////////////

interface IBreadCrumbsProps { 
  items: IBreadcrumbItem[]
  onClick: (item: IBreadcrumbItem, i: number) => void
  onShare: (item: IBreadcrumbItem) => void
}

class Breadcrumbs extends React.PureComponent<IBreadCrumbsProps> {
  public render() {
    const { items, onClick, onShare } = this.props

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
            if ((item as any).isExecFolder) {
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

export interface IExecDocumentsProps {
  items: IFSObject[]
  rootFolder: string
  startInsideFolder: IFSObject
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

export interface IExecDocumentsState {
  sort: ISortableList
  rowContextMenuMouseEvent: MouseEvent
  backgroundContextMenuMouseEvent: MouseEvent
  breadcrumbs: any[]
  currentFolder: string
  columns: IColumn[]
  activeItem: IFSObject
  loading: boolean
  //dropzone2
  dropToFolder: any;
  selectionDetails?: string;
  isLocallyDrop:boolean;
  listStyle:boolean;
 
}

export interface ISortableList {
  index: number
  asc: boolean
}

export default class ExecDocuments extends React.Component<IExecDocumentsProps, IExecDocumentsState> {
  
  constructor(props: IExecDocumentsProps) {
    super(props)
    //dropzone2
    this._selection = new Selection();
    this.state = {
      //dropZone2
      //flags used for dropzone
      dropToFolder:null,
      selectionDetails:null,
      isLocallyDrop:true,
      listStyle:false,
      //end
      loading: false,
      activeItem: null,
      sort: {
        index: 2,
        asc: true
      },
      rowContextMenuMouseEvent: null,
      backgroundContextMenuMouseEvent: null,
      breadcrumbs: [
        { text: "Private Documents", serverRelativeUrl: props.rootFolder }
      ],
      currentFolder: props.rootFolder,
      columns: [
        {
          key: "column0",
          fieldName: "share",
          name: "Share",
          minWidth: 32,
          maxWidth: 32,
          isResizable: false,
          onRender: this.renderShareIcon.bind(this)
        }, {
          key: "column1",
          fieldName: "type",
          name: "Type",
          minWidth: 42,
          maxWidth: 42,
          isResizable: false,
          onRender: this.renderTypeIcon.bind(this)          
        }, {
          key: "column2",
          fieldName: "name",
          name: "Name",
          minWidth: 150,
          maxWidth: 500,
          isResizable: true,
          onColumnClick: () => this.sortColumn(2),
          onRender: this.renderName.bind(this)
        }, {
          key: "column3",
          fieldName: "modified",
          name: "Modified Date",
          minWidth: 100,
          maxWidth: 100,
          isResizable: false,
          onColumnClick: () => this.sortColumn(3),
          onRender: item => item.modified.format("YYYY-MM-DD"),
        }, {
          key: "column4",
          fieldName: "modifiedBy",
          name: "Modified By",
          minWidth: 150,
          maxWidth: 150,
          isResizable: true,
          onColumnClick: () => this.sortColumn(4),
          onRender: item => item.modifiedBy
        }, {
          key: "column5",
          fieldName: "size",
          name: "File Size",
          minWidth: 50,
          maxWidth: 50,
          isResizable: false,
          onColumnClick: () => this.sortColumn(5),
          onRender: item => item.size ? humanFileSize(item.size) : null
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
    if (this.props.startInsideFolder) {
      this.onFolderClick(this.props.startInsideFolder)
    }

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
          isExecFolder: this.isExecFolderLevel()
        }
      ],
      currentFolder: item.serverRelativeUrl
    })
  }

  @autobind
  private isExecFolderLevel(): boolean {
    return this.state.currentFolder.split("/").length === 5
  }

  @autobind
  private isInsideExecFolder(): boolean {
    return this.state.currentFolder.split("/").length > 5
  }

  @autobind
  private renderShareIcon(item: IFSObject) {
    if (
      this.isInsideExecFolder() &&
      item.type === "file"
    ) {
      return (
        <Icon
          className={css({
            [styles.shareIcon]: true,
            [styles.shareIconUniquePerms]: (
              item.hasUniquePermissions &&
              item.directAccessUsers > 0
            )
          })}
          iconName="Share"
          onClick={() => this.onShareIconClick(item)}
        />
      )
    }
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
      const fsObject = find(this.props.items, x => x.name === breadcrumb.text)
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
  //dropZone2
  private _selection: Selection;
  //dropzone
  //used when file locally drag or in case DetailList contain 0 item
  private onDrop(uploadUserFile,fromLocal) {
    this.setState({listStyle:false})
    if (this.state.breadcrumbs.length > 1)
    {
        let uploads = []
        if(_userDragFileByEvent.length>0)
        {
          uploads.push(shareFile(this.state.currentFolder, _userDragFileByEvent[0]))
        }
        else
        {
          if(fromLocal){
            
              for(let i=0; i<uploadUserFile.length; i++) {
                let customCurrentFolder=this.state.currentFolder?this.state.currentFolder:_currentFolder;
                let Extension=uploadUserFile[i].name.split(/\.(?=[^\.]+$)/);
                if(Extension[1]!="url")
                uploads.push(addFile(customCurrentFolder, uploadUserFile[i].name, uploadUserFile[i]))   
              }
            
          }
          else{
            let customCurrentFolder=this.state.dropToFolder.serverRelativeUrl
            uploads.push(moveFile(customCurrentFolder,uploadUserFile[0]))
          }
        }
        Promise.all(uploads).then( () => {
          this.props.onFileUpload(this.state.currentFolder);
          this.setState({isLocallyDrop:true})
        })
    }
  
  }
  //dropZone2
  //All events for DetailList
 
  private _getDragDropEvents(e): IDragDropEvents {
    return {
    
      canDrop: (dropContext?: IDragDropContext, dragContext?: IDragDropContext) => {
        if (this.state.breadcrumbs.length > 1)
          return true; 
        else
          return false;
      },
      canDrag: (item1?: any) => {
        if(item1.type!="folder"){
          if (this.state.breadcrumbs.length > 1 )
          { 
            return true;
          }
        }
       
        return false;
        },
      onDrop: (item1?: any, event?: DragEvent) => {
        this.setState({isLocallyDrop:true});
        this.setState({listStyle:false})
        this.setState({ dropToFolder: item1});
        if (_draggedItem.length>0) {
          this.onDrop(_draggedItem,false);
        }
      },
      onDragStart: (item1?: any, itemIndex?: number, selectedItems?: any[], event?: MouseEvent) => {
        this.setState({listStyle:false})
        if(item1.type!="folder")
        this.setState({isLocallyDrop:false})
        _draggedItem = selectedItems;
        _draggedIndex = itemIndex!;
        _userDragFile=selectedItems;
        _currentFolder= this.state.currentFolder;
        return true;
      },
      onDragEnd: (item1?: any, event?: DragEvent) => {
        this.setState({listStyle:false})
        this.setState({isLocallyDrop:true})
        _draggedItem = null;
        _userDragFile=[];
      },
       onDragEnter: (item1?: any, event?: DragEvent) => {
        const selectionCount = this._selection.getSelectedCount();
          this.setState({listStyle:false});
          if(selectionCount>0)
          this.setState({isLocallyDrop:false});
          return "";
           }, 
          onDragLeave: (item1?: any, event?: DragEvent) => {
           this.setState({listStyle:false})
             return; },
    };
  }
  //Active dropzone
  private setActive(event,file)
  {
    if(this.state.isLocallyDrop )
    {
    this.setState({listStyle:true})
    }
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
          onShare={this.onShareEventClick}
        />
        <div >
          {/* Add file Dropzone*/}
         <Dropzone   onDragOver={this.setActive.bind(this)} onDragLeave={this.setDeactivate.bind(this)}
        onDrop={this.onDrop.bind(this)}   style={styles.dropzoneStyle} disableClick={true}  activeClassName={(this.state.listStyle)&&(this.state.breadcrumbs.length > 1)?styles.dropzoneStyleActive:"" } >
         
        <div className={styles.DetailsList}
          ref={this.detailsListContainerRef.bind(this)}
          onContextMenu={(evt) => {
            evt.preventDefault()
            // We only want to show the menu when the target
            // is the DetailsList div and we're not at the top
            // level (because this menu isn't allowed)
            if (
              (evt.nativeEvent.target as any).className.indexOf("DetailsList") === 0 &&
              this.state.breadcrumbs.length > 1
            ) {
              this.setState({ backgroundContextMenuMouseEvent: evt.nativeEvent })
            }
          }}>
           
          {/* console.log(this.state.columns, 'COLS', items, 'ITEMS') */}
          {/*dropZone*/}
          {/* enable selection for detailList */}
          <MarqueeSelection selection={ this._selection }> 
          <DetailsList
            //add events for dropzone
            dragDropEvents={ this._getDragDropEvents(this) }
             selection={this._selection}

            columns={this.state.columns}
            items={items}
            layoutMode={DetailsListLayoutMode.justified}
            constrainMode={ConstrainMode.unconstrained}
            checkboxVisibility={CheckboxVisibility.hidden}
            selectionMode={SelectionMode.none}
            onItemContextMenu={(item, i, evt: MouseEvent) => {
              if (this.state.breadcrumbs.length > 1) {
                this.setState({ 
                  rowContextMenuMouseEvent: evt,
                  activeItem: item
                })
              }
            }} />
          </MarqueeSelection>        
          {rowContextMenuMouseEvent && 
          (this.state.activeItem.type == 'file') && 
           (this.state.breadcrumbs.length > 1) && (
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
          (this.state.activeItem.type == 'folder') && 
           (this.state.breadcrumbs.length > 1) && (
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
          {backgroundContextMenuMouseEvent && 
            (this.state.breadcrumbs.length > 1) && (
            <ContextualMenu
            items={[
              { key: "1", name: "New Folder" },
              { key: "4", name: "Upload from Computer", onRender: this.renderFileInput },
              { key: "3", name: "Sync to OneDrive" },
            ]}
            target={backgroundContextMenuMouseEvent}
            onItemClick={(evt, item: IContextualMenuItem) => {
              switch (item.key) {
                case "1": this.onNewFolderClick(); break
                case "4": this.onUploadFileFromComputerClick(); break
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
          
            < div className={styles.dragZoneOperation}>
               {
                 (this.state.listStyle)&&(this.state.breadcrumbs.length > 1)? <div className={styles.displayImg}> <img src={ `${cdnAssetsBaseUrl}/images/addFile.png` } /><div>Upload</div></div>:""  
                
               }
               { (!this.state.isLocallyDrop)&&(this.state.breadcrumbs.length > 1)? <div className={styles.displayImg}> <img src={ `${cdnAssetsBaseUrl}/images/moveFile.png` } /><div>Move</div></div> :""  }
               
            </div>
          </div>
      </div>
    )
  }  
}
