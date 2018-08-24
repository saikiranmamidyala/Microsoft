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

import * as moment from 'moment';

import {
  IFSObject,
  redirectToSitePage,
  uploadFile,
  renameFile,
  siteDomain,
  siteUrl,
  getFileDownloadLink,
  execAssetsLibraryTitle
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

import styles from '../webparts/assetsPage/components/AssetsPage.module.scss'
import { autobind } from '@uifabric/utilities';
import { FileAddResult, ItemUpdateResult} from 'sp-pnp-js/lib/sharepoint';
import { FolderUpdateResult } from 'sp-pnp-js/lib/sharepoint/folders'
import { find, debounce } from 'lodash';
import { css } from 'office-ui-fabric-react/lib/Utilities';
// import { MouseEvent } from 'react';



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


export interface ISortableList {
  index: number
  asc: boolean
}

export interface IExecAssetsProps {
  description?: string;
  execs: object[];
  getExecChildren: any;
  atExecRoot: boolean;
  onOpenFileInBrowserClick?: (item: IFSObject) => void;
  onRenameFileClick?: (item: IFSObject) => void;
  onDuplicateFileClick?: (item: IFSObject) => void;
  onDeleteFileClick?: (item:IFSObject) => void;
  onUploadFileFromComputerClick?: (currentFolder: string) => void
  onSyncToOneDriveClick?: (currentFolder: string) => void
  onFileUpload?: (currentFolder: string) => void;
  onNewFolderClick?: (currentFolder: string) => void
}

export interface IExecAssetsState {
  columns: IColumn[];
  defaultColumns: IColumn[];
  execs: object;
  atExecRoot: boolean;
  currentFiles: IFSObject[],
  currentFolder: string,
  currentExec: string,
  folderBreadCrumbs: any[],
  defaultBreadCrumb: any[],
  currentLocation: string,
  currentUrl: string,
  previousFolders: object,
  rowContextMenuMouseEvent: MouseEvent,
  backgroundContextMenuMouseEvent: MouseEvent,
  selectedItem: IFSObject,
  rows: any,
  execRows: any,
  rightClicked: boolean,
  activeItem: any,
  activeItemIndex: number,
}

export default class ExecAssets extends React.Component<IExecAssetsProps, IExecAssetsState> {
  constructor(props) {
    super(props);

    let defaultBc = [
      {
        key: 'defaultBc',
        text: 'Exec Assets',
        onClick: (e, item) => {
          this.setState(prevState => {
            return {
              atExecRoot: true,
              columns: prevState.defaultColumns,
              folderBreadCrumbs: prevState.defaultBreadCrumb,
            }
          })
        }
      }
    ]


    this.state = {
      columns: null,
      defaultColumns: null,
      execs: this.props.execs,
      atExecRoot: this.props.atExecRoot, //indicates we are at exec root folders on initial page load
      currentFiles: [],
      currentExec: '',
      currentLocation: '',
      currentFolder: '',
      folderBreadCrumbs: defaultBc,
      defaultBreadCrumb: defaultBc,
      currentUrl: '',
      previousFolders: {},
      rowContextMenuMouseEvent: null,
      backgroundContextMenuMouseEvent: null,
      selectedItem: null,
      rows: [],
      execRows: [],
      rightClicked: false,
      activeItem: null,
      activeItemIndex: null,
    }
  }

  @autobind
  private onDeleteFileClick() {
    if (this.props.onDeleteFileClick) {
      this.props.onDeleteFileClick(this.state.activeItem)
    }
  }  

  @autobind
  private downloadFile(item) {
    // console.log(item.Name.props.children)
    let fsObject = this.state.execs[this.state.currentExec].filter(file => {
      return file.name === item.Name.props.children
    })[0];
    // console.log(fsObject, 'WHAT IS ITEM')
    const url = getFileDownloadLink(fsObject);
    window.open(url, '_blank')
  }
  
  @autobind
  private setInitialExecRows(execs) {
    return Object.keys(execs).reduce((acc, execKey) => {
      //moment(execs[execKey][0].modified).format('YYYY-MM-DD'),
      acc.push({
        Name: execKey,
        'Modified Date': moment(execs[execKey][0].modified).format('YYYY-MM-DD'),
        'Modified By': execs[execKey][0].modifiedBy,
      })
      return acc;
    }, [])
  }

  @autobind
  private isSubFolderRenderChildren(childItem) {
    // if (!event.bubbles) return;
    if (childItem.type === 'folder') {
        this.setState(prevState => {
          // console.log(prevState[pre], 'WHAT ARE FILES')
          let currentLocation = childItem.name;
          let files = prevState.execs[prevState.currentExec];
          let currentFiles = this.reformatFiles(files, currentLocation)
          // if (!currentFiles.length) return;
          return {
            currentLocation,
            currentFiles,
            previousFolders: this.setPreviousFoldersCache(currentFiles, prevState, currentLocation),
          }
        })
      
    } 
  }
  
  @autobind
  private isRootRenderChildren(exec) {
    this.setState(prevState => {
      let currentFiles = this.reformatFiles(prevState.execs[exec.Name], exec.Name)
      return {
        atExecRoot: false,
        currentExec: exec.Name,
        currentLocation: exec.Name,
        currentFiles,
        previousFolders: this.setPreviousFoldersCache(currentFiles, prevState, exec.Name),
        // columns: this.renderExecChildrenColHeaders(),
      }
    })
  }

  public componentWillReceiveProps(nextProps) {
    this.setState(prevState => {
      return {
        execs: nextProps.execs,
        atExecRoot: nextProps.atExecRoot,
        folderBreadCrumbs: nextProps.atExecRoot ? this.state.defaultBreadCrumb : prevState.folderBreadCrumbs,
        rows: this.setInitialExecRows(nextProps.execs),
        execRows: this.setInitialExecRows(nextProps.execs),
      }
    })
  }

  @autobind 
  private setPreviousFoldersCache(files, prevState, currentLocation) {
    prevState.previousFolders[currentLocation] = files;
    return prevState.previousFolders;
  }

  @autobind
  private setNewBreadCrumbs(prevState, currentLocation) {
    let folderBreadCrumbs;
    // console.log(currentLocation)
    // debugger

    //when user clicks a folder, is that folder already the current breadcrumb? If not, do this logic
    if (prevState.folderBreadCrumbs.map(e => e.text).indexOf(currentLocation) <= -1) {
      folderBreadCrumbs = prevState.folderBreadCrumbs.concat({
        key: `crumb${prevState.folderBreadCrumbs.length}`,
        text: currentLocation,
        // isCurrentItem: true,
        onClick: (e, item) => {
          // console.log(item, 'WHAT IS CLICKED BREADCRUMB')
          item.isCurrentItem = true;
          // if (this.state.currentLocation === item.text) return;
          let currentFiles = this.state.previousFolders[currentLocation];
          this.setState(prev => {
            return {
              currentFiles,
              currentLocation: item.text,
              folderBreadCrumbs: prev.folderBreadCrumbs.slice(0, prev.folderBreadCrumbs.map(bc => bc.text).indexOf(item.text) + 1),
              previousFolders: this.setPreviousFoldersCache(currentFiles, prev, item.text)
            }
          })
        }
      })
    } else {
      folderBreadCrumbs = prevState.folderBreadCrumbs;
    }
    return folderBreadCrumbs;
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

  @autobind
  private renderExecChildrenColHeaders() {
    let cols = [
      {
        key: 'column0',
        fieldName: 'Type',
        name: 'Type',
        minWidth: 45,
        maxWidth: 45,
        isResizable: false,
      },
      {
        key: 'column1',
        fieldName:  'Name',
        name: 'Name',
        minWidth: 150,
        maxWidth: 500,
        isResizable: false,
      },
      {
        key: 'column2',
        fieldName:  'File Size',
        name: 'File Size',
        minWidth: 150,
        maxWidth: 500,
        isResizable: false,
      },
      {
        key: 'column3',
        fieldName:  'Modified Date',
        name: 'Modified Date',
        minWidth: 150,
        maxWidth: 500,
        isResizable: false,
      },
      {
        key: 'column4',
        fieldName:  'Modified By',
        name: 'Modified By',
        minWidth: 150,
        maxWidth: 500,
        isResizable: false,
      }
    ]
    return cols;
  }

  @autobind
  private setLink(item) {
    if (item.type === 'file') {
      this.openFile(item);
    } else {
      this.isSubFolderRenderChildren(item);
      
      this.setState(prevState => {
        return {
          folderBreadCrumbs: this.setNewBreadCrumbs(prevState, item.name)
        }
      })
    }
  }

  @autobind
  private reformatFiles(files, currentLocation) {
    return files.filter(file => this.isCurrentFolderLevel(file.serverRelativeUrl, currentLocation))
    .map(file => {
      return {
        Type: this.renderTypeIcon(file),
        Name: <Link
                  onClick={() => {
                    this.setLink(file);
                  }}
              >{file.name}</Link>,
        'File Size': this.convertFileSizes(file.size),
        'Modified Date': moment(file.modified).format('YYYY-MM-DD'),
        'Modified By': file.modifiedBy,
        
        'name': file.name,
        'type': file.type,
        'id': file.id,
        'serverRelativeUrl': file.serverRelativeUrl,   
        modifiedBy: file.modifiedBy,
        modified: file.modified,
        uniqueFolderId: file.uniqueFolderId,
        uniqueFileId: file.uniqueFileId,
        hasUniquePermissions: file.hasUniquePermissions,
        serverRedirectedEmbedUrl: file.serverRedirectedEmbedUrl,
        directAccessUsers: file.directAccessUsers,
        fileExtension: file.fileExtension,
        isDocument: file.isDocument,
        isContainer: file.isContainer,
        author: file.author
                  
      }
    })
  }
 
  @autobind
  private isCurrentFolderLevel(url, currentLocation): boolean {
    return url.split('/')[url.split('/').length - 2] === currentLocation;
  }


  
  @autobind
  private renderExecChildrenColData(execChild) {
    let cols = [
     {
       key: 'column0',
       fieldName: execChild.type as string,
       name: this.renderTypeIcon(execChild),
       minWidth: 45,
       maxWidth: 45,
       isResizable: false,
       onRender: this.renderTypeIcon.bind(this)
     },
     {
       key: 'execChildName',
       fieldName: 'execChildName',
       name: execChild.name,
       minWidth: 150,
       maxWidth: 500,
       isResizable: false,
       onColumnClick: () => {
         if (cols[0].fieldName === 'folder') {
          let currentLocation = execChild.name;
          //TODO: apply same functionality here grabbing from this.state.execs instead of api call
          //LOGIC: click on child folder inside an exec, grab all the files from that folder that are at that folder's root

          this.setState(prevState => {
            return {
              columns: this.renderExecChildrenColHeaders(),
              // atExecRoot: false,
              currentFiles: this.state.execs[this.state.currentExec].filter(file => this.isCurrentFolderLevel(file.serverRelativeUrl, currentLocation)),
              currentLocation,
              previousFolders: prevState.previousFolders,
              folderBreadCrumbs: this.setNewBreadCrumbs(prevState, currentLocation),
            }
          })
         } else {
           this.openFile(execChild);
         }
       }
     },
     {
       key: 'column2',
       fieldName: 'fileSize',
       name: this.convertFileSizes(execChild.size),
       minWidth: 150,
       maxWidth: 500,
       isResizable: false,
     },
     {
      key: 'column3',
      fieldName: 'modified',
      name: moment(execChild.modified).format('YYYY-MM-DD'),
      minWidth: 150,
      maxWidth: 500,
      isResizable: false,
     },
     {
      key: 'column3',
      fieldName: 'modifiedBy',
      name: execChild.modifiedBy,
      minWidth: 150,
      maxWidth: 500,
      isResizable: false,
     }

    ]
    return cols;
  }

  @autobind
  private openFile(item) {
    const ext = getFileExt(item.name)
    let prefix = ""

    if (ext === "pptx") {
      prefix = "ms-powerpoint:ofe%7Cu%7C"
    } else if (ext === "docx") {
      prefix = "ms-word:ofe%7Cu%7C"
    } else if (ext === "xslx") {
      prefix = "ms-excel:ofe%7Cu%7C"
    } 
    const href = prefix + siteDomain + item.serverRelativeUrl
    window.open(href);

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

  public getActiveItem() {
    return this.state.activeItem
  }

  @autobind
  private onNewFolderClick() {
    if (this.props.onNewFolderClick) {
      this.props.onNewFolderClick(this.state.currentLocation)
    }
  }

  @autobind
  private onUploadFileFromComputerClick() {
    if (this.props.onUploadFileFromComputerClick) {
      this.props.onUploadFileFromComputerClick(this.state.currentLocation)
    }
  }

  @autobind
  private onSyncToOneDriveClick() {
    if (this.props.onSyncToOneDriveClick) {
      this.props.onSyncToOneDriveClick(this.state.currentLocation);
    }
  }

  @autobind
  //referenced this: https://stackoverflow.com/questions/15900485/correct-way-to-convert-size-in-bytes-to-kb-mb-gb-in-javascript
  private convertFileSizes(bytes: number) {
    if (bytes === 0) return '';

    let k = 1024
    let dm = 2;
    let sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];
    let i  = Math.floor(Math.log(bytes) / Math.log(k));

    return `${ (bytes / Math.pow(k, i)).toFixed(dm)} ${sizes[i]}`
  }

  @autobind
  private renderExecRootFolderData(execsDict, execKey) {
    let cols = [
      {
        key: 'column0',
        fieldName: execKey.split(' ')[0],
        name:  execKey,
        //minWidth: 150,
        //maxWidth: 500,'
        minWidth:45,
        maxWidth:45,
        isResizable: false,
        onColumnClick: () => {
          // debugger
          //TODO: Breadcrumbs don't work, fix those, check out the other execChildfunction to do same thing there
          let currentFiles = this.state.execs[execKey].filter(file => this.isCurrentFolderLevel(file.serverRelativeUrl, execKey));
          this.setState(prevState => {
            return {
              columns: this.renderExecChildrenColHeaders(),
              atExecRoot: false,
              currentFiles,
              currentExec: execKey,
              currentLocation: execKey,
              //TODO: setting prev state prev files here
              previousFolders: this.setPreviousFoldersCache(currentFiles, prevState, execKey),
              folderBreadCrumbs: this.setNewBreadCrumbs(prevState, execKey),
            }
          })
        },
      },
      {
        key: 'column1',
        fieldName:  'modifiedDate',
        name: execsDict[execKey].modifiedDate,
        minWidth: 150,
        maxWidth: 500,
        isResizable: false,
      },
      {
        key: 'column2',
        fieldName:  'modifiedBy',
        name: execsDict[execKey].modifiedBy,
        minWidth: 150,
        maxWidth: 500,
        isResizable: false,
      }
    ]

    return cols;
  }


  private assetsListContainerRef(div: HTMLDivElement) {
    if (div) {
      const windowHeight = window.innerHeight
      const divTop = div.getBoundingClientRect().top
      const height = windowHeight - divTop
      div.style.minHeight = `${height}px`
    }
  }

  public render() {
    const {
      rowContextMenuMouseEvent,
      backgroundContextMenuMouseEvent
    } = this.state

    return (
      <div>
        <div>
          <Breadcrumb 
            items={this.state.atExecRoot ? this.state.defaultBreadCrumb : this.state.folderBreadCrumbs}
          />
        </div>
        <div>
          {
            this.state.atExecRoot &&
              <DetailsList 
                // columns={ this.state.atExecRoot ?  this.state.defaultColumns : this.state.columns }
                columns={ this.renderExecChildrenColHeaders()}
                items={this.state.rows}
                layoutMode={DetailsListLayoutMode.justified}
                constrainMode={ConstrainMode.unconstrained}
                checkboxVisibility={CheckboxVisibility.hidden}
                selectionMode={SelectionMode.none}
                onActiveItemChanged={(item, index, evt) => {
                  this.isRootRenderChildren(item)
                  
                  this.setState(prevState => {
                    return {
                      folderBreadCrumbs: this.setNewBreadCrumbs(prevState, prevState.currentLocation)
                    }
                  })
                 
                }}
              />
          }
        </div>
 
    {!this.state.atExecRoot &&    
      <div className={styles.execRow}>
        <div className={styles.DetailsList}
        ref={this.assetsListContainerRef.bind(this)}
        onContextMenu={(evt) => {
          evt.preventDefault()
          // We only want to show the menu when the target
          // is the DetailsList div and we're not at the top
          // level (because this menu isn't allowed)
          if ((evt.nativeEvent.target as any).className.indexOf("DetailsList") === 0) {
            this.setState({ backgroundContextMenuMouseEvent: evt.nativeEvent })
          }
        }}>
          <DetailsList 
            // columns={ this.renderExecChildrenColData(File) }
            columns={ this.renderExecChildrenColHeaders()}
            items={this.state.currentFiles}
            layoutMode={DetailsListLayoutMode.justified}
            constrainMode={ConstrainMode.unconstrained}
            checkboxVisibility={CheckboxVisibility.hidden}
            selectionMode={SelectionMode.none}
            onItemContextMenu={(item, i, evt: MouseEvent) => {
              this.setState({ 
                rightClicked: true,
                activeItem: item,
                activeItemIndex: i,
                rowContextMenuMouseEvent: evt,
              })
              evt.preventDefault();
              evt.stopImmediatePropagation();

            }}

          />
        </div>
      </div>
    }
    {
        rowContextMenuMouseEvent &&
                this.state.rightClicked && this.state.activeItem['File Size'] !== '' && (
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
        target={this.state.rowContextMenuMouseEvent}
        onItemClick={(evt, item: IContextualMenuItem) => {
          switch (item.key) {
            case "1": this.downloadFile(this.state.activeItem); break
            case "2": this.onOpenFileInBrowserClick(); break
            case "4": this.onRenameFileClick(); break
            case "5": this.onDuplicateFileClick(); break
            case "7": this.onDeleteFileClick(); break
            default: break
          }
        }}
        onDismiss={() => {
          this.setState({ 
            rowContextMenuMouseEvent: null,
            rightClicked: false,
          })
        }} />
      )
    }
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
                //case "2": this.onUploadFileFromComputerClick(); break
                //case "3": this.onSyncToOneDriveClick(); break
                default: break
              }
            }}
            onDismiss={() => {
              this.setState({ backgroundContextMenuMouseEvent: null })
            }} />        
          )}
    </div>
    )
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
        //this.props.onFileUpload(this.state.currentFolder)
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
}

