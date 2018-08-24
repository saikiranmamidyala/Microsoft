import * as React from 'react';
import styles from './AssetsPage.module.scss';


//Office UI Fabric components
import { Breadcrumb, IBreadcrumb, IBreadcrumbItem } from 'office-ui-fabric-react/lib/Breadcrumb'
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button'
import { registerDefaultFontFaces } from '@uifabric/styling/lib/styles/DefaultFontStyles';
import { Dialog } from 'office-ui-fabric-react/lib/Dialog';

//our components
import ExecFilter from '../../../components/ExecFilter'
import Loading from '../../../components/Loading'
import ExecDocuments from '../../../components/ExecDocuments'
import EventDocuments from '../../../components/EventDocuments'
import SharingDialog from '../../../components/SharingDialog'
import TextInputDialog from '../../../components/TextInputDialog'
import DeleteConfirmationDialog from '../../../components/DeleteConfirmationDialog'
import EventDetailsDialog from "../../../components/EventDetailsDialog"
import  ExecAssets from "../../../components/ExecAssets"

import {
  redirectToSitePage,
  siteDomain,
  IEvent, getEventById,
  createFolder,
  getFolderUrl,
  getFiles,
  IFSObject,
  IEventListDetails, getEventListDetails,
  deleteItem,
  renameFile,
  getEventICSLink,
  getFileDownloadLink,
  IExecutive,
  duplicateFile,
  getAllExecutives,
  execAssetsLibraryTitle,
  isCurrentUserAnAdmin
} from '../../../shared/SharePoint'

import { getQueryParameters } from '../../../shared/util'
import { PageContext } from '@microsoft/sp-page-context'
import { find } from 'lodash';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';

import { mapExecsToFolders, getExecChildren } from './AssetsPage.request';

import {
  IColumn,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
  ConstrainMode
} from 'office-ui-fabric-react/lib/DetailsList'
//import {  } from '../../../../lib/shared/SharePoint';



export interface IExecAssetsPageProps {
  description: string;
  context: PageContext
}

export interface IExecAssetsPageState {
  loading: boolean
  breadcrumbItems: any[]
  eventDetails: IEvent
  eventFolderUrl: string
  execFiles: IFSObject[]
  execDocumentsStartFolder: string
  startInsideFolder: IFSObject
  eventFiles: IFSObject[]
  eventDocumentsStartFolder: string
  showFileSharingDialog: IFSObject
  showEventSharingDialog: IFSObject
  showEventDetailsDialog: boolean
  eventListDetails: IEventListDetails
  showTextInputDialog: string
  showDeleteConfirmationDialog: IFSObject
  showRenameDialog: IFSObject
  openInBrowser: IFSObject
}

export interface IAssetsPageProps {
  description: string;
  context: PageContext
}


export interface IAssetsState {
  //userIsAdmin: boolean,
  topBreadCrumbs: any,
  execsDict: any,
  execsFilter: any,
  atExecRoot: boolean,
  showFileSharingDialog: IFSObject
  showEventSharingDialog: IFSObject
  showEventDetailsDialog: boolean
  eventListDetails: IEventListDetails
  showTextInputDialog: string
  showDeleteConfirmationDialog: IFSObject
  showRenameDialog: IFSObject
  openInBrowser: IFSObject
  
}

export default class AssetsPage extends React.Component<IAssetsPageProps, IAssetsState> {
  constructor(props) {
    super(props);

    let topBreadCrumbs = [ 
      {
      key: 'crumb0',
      text: 'Exec Assets',
      isCurrentItem: true,
      onClick: () => {
        // console.log('INSIDE TOP BC CLICK HANDLER')
        this.setState(prevState => {
          // console.log(prevState)
          return {
            atExecRoot: true,
          }
        })
      }
      } as IBreadcrumbItem,
   ]

 

    this.state = {
      //userIsAdmin: false,
      topBreadCrumbs,
      execsDict: [],
      execsFilter: [],
      atExecRoot: true,
      showFileSharingDialog: null,
      showEventSharingDialog: null,
      eventListDetails: null,
      showEventDetailsDialog: false,
      showTextInputDialog: null,
      showDeleteConfirmationDialog: null,
      showRenameDialog: null,
      openInBrowser: null,
    }
  }

  public componentDidMount() {
    getAllExecutives()
    .then(execs => {
      // console.log(execs)
      this.setState({
        execsFilter: execs,
      })
      return mapExecsToFolders(execs);
    })
    .then(items => {
      this.setState( {
        execsDict: items
      })
    })
  }

  @autobind
  private openInBrowser(item: IFSObject) {
    const url = item.serverRedirectedEmbedUrl.split('&')[0];
    window.open(url, "_blank")
  }

  @autobind
  private showRenameDialog(item: IFSObject) {
    this.setState({atExecRoot: false})
    this.showDialog("showRenameDialog", item)
  }

  @autobind
  private async deleteSelectedItem(item: IFSObject) {
    await deleteItem(item, execAssetsLibraryTitle)
    //this.fetchFiles()
    // this.setState({
    //   execFiles: this.state.execFiles.filter(e => e !== item),
    //   eventFiles: this.state.eventFiles.filter(e => e !== item)
    // })
  }

  @autobind
  private showDeleteItemDialog(item: IFSObject) {
    this.showDialog("showDeleteConfirmationDialog", item)
  }
  
  @autobind
  private async renameSelectedItem(item: IFSObject, newName: string) {
    const parentFolder = item.serverRelativeUrl.substring(0, item.serverRelativeUrl.lastIndexOf('/'))
    await renameFile(execAssetsLibraryTitle, item, newName)
    
    //TODO : refresh the files list and statue
    const files = await getFiles(parentFolder, execAssetsLibraryTitle)
    const curLocation = item.serverRelativeUrl
    console.log("currlocation: " + curLocation)
    
    //await this.fetchFiles()
  }
  @autobind
  private getExecChildAssets(path) {
    let currentFolderPath = path ? path : ""
    getExecChildren(path)
  }

  
  @autobind
  private async duplicateFileItem(item: IFSObject) {
    await duplicateFile(item)

    // let currentFolder = item.serverRelativeUrl.substring(0, item.serverRelativeUrl.lastIndexOf("/"))
    // const files = await getFiles(currentFolder)
    // this.refreshList(currentFolder, files)
    //this.fetchFiles()
  }

  @autobind
  private showNewFolderDialog(currentFolder: string) {
    this.setState({ showTextInputDialog: currentFolder })
  }

  @autobind
  private async createNewFolder(newFolderName: string) {
    const currentFolder = this.state.showTextInputDialog
    this.hideDialog("showTextInputDialog")
    const results = await createFolder(currentFolder, newFolderName, execAssetsLibraryTitle)
    
    console.log("Assets Page - createFolder results: " + results)
    //const files = await getFiles(currentFolder)
    //this.refreshList(currentFolder, files)
    //await this.fetchFiles()
  }

  

  public render(): React.ReactElement<IAssetsPageProps> {
    
    const {
      showFileSharingDialog,
      showEventSharingDialog,
      showEventDetailsDialog,
      showTextInputDialog,
      showDeleteConfirmationDialog,
      showRenameDialog
    } = this.state

    return (
      <div className={ styles.assetsPage }>
        <div>
          <Breadcrumb 
            items={ this.state.topBreadCrumbs }
          />
        </div>
        <div className={ styles.execFilter }>
          <ExecFilter 
            execs={ this.state.execsFilter }
            onChange={ () => { /*console.log('clicked')*/ }}
          />
        </div>
      
        <div>
          <ExecAssets 
            execs={this.state.execsDict}
            getExecChildren={() => {
              return (
                this.state.atExecRoot ? 
                this.getExecChildAssets :  null
              )
            }}
            atExecRoot={ this.state.atExecRoot }
            onOpenFileInBrowserClick={ this.openInBrowser }
            onRenameFileClick={this.showRenameDialog}
            onDuplicateFileClick={this.duplicateFileItem}
            onDeleteFileClick={this.showDeleteItemDialog}
            onNewFolderClick={this.showNewFolderDialog}
            //onFileUpload={this.onFileUpload}
          />
        </div>
        {
          showRenameDialog && 
          <TextInputDialog
            title="Rename item"
            value=""
            extension={
              showRenameDialog.name.lastIndexOf('.') > -1 ? 
              showRenameDialog.name.substring(showRenameDialog.name.lastIndexOf('.') + 1 ) : ""
            }
            onDismiss={() => this.hideDialog("showRenameDialog")}
            onSave={text => {
              this.renameSelectedItem(showRenameDialog, text)
              this.getExecChildAssets(showRenameDialog.serverRelativeUrl.split('/').slice(0, -1).join('/'))
            }}
          />
        }
        {
          showDeleteConfirmationDialog && (
            <DeleteConfirmationDialog 
              item={showDeleteConfirmationDialog}
              onDismiss={() => this.hideDialog("showDeleteConfirmationDialog")}
              onDelete={() => this.deleteSelectedItem(showDeleteConfirmationDialog)}
            />
          )
        }
        {showTextInputDialog && (
          <TextInputDialog
            title="Folder Name"
            value=""
            onDismiss={() => this.hideDialog("showTextInputDialog")}
            onSave={text => this.createNewFolder(text)}
          />
        )}
      </div>
    );
  }
  @autobind
  private showDialog(key: string, data: any = true) {
    this.setState(({ [key]: data }) as any)
  }

  @autobind
  private hideDialog(key: string, data: any = null) {
    this.setState(({ [key]: data }) as any)
  }
}
