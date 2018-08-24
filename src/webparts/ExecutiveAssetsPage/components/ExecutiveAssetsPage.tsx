import * as React from 'react'

import { Breadcrumb, IBreadcrumb, IBreadcrumbItem } from 'office-ui-fabric-react/lib/Breadcrumb'
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button'
import { registerDefaultFontFaces } from '@uifabric/styling/lib/styles/DefaultFontStyles';
import { Dialog } from 'office-ui-fabric-react/lib/Dialog';

import Loading from '../../../components/Loading'
//import ExecDocuments from '../../../components/ExecutiveAssets'
import ExecutiveAssets from '../../../components/ExecutiveAssets'
import SharingDialog from '../../../components/SharingDialog'
import TextInputDialog from '../../../components/TextInputDialog'
import DeleteConfirmationDialog from '../../../components/DeleteConfirmationDialog'

import {
  redirectToSitePage,
  siteDomain,
  IEvent, getEventById,
  createFolder,
  getFolderUrl,
  getFiles,
  IFSObject,
  //IEventListDetails, getEventListDetails,
  deleteItem,
  renameFile,
  //getEventICSLink,
  getFileDownloadLink,
  IExecutive,
  duplicateFile,
  execAssetsLibraryTitle,
  siteCollectionUrl,
  siteUrl,  
  getListGUID
} from '../../../shared/SharePoint'

import { getQueryParameters } from '../../../shared/util'
import { PageContext } from '@microsoft/sp-page-context'
import styles from './ExecutiveAssets.module.scss'
import styles2 from './ExecutiveAssetsPage.module.scss'
import { find } from 'lodash'
import { autobind, css } from 'office-ui-fabric-react/lib/Utilities'
import { isCurrentUserAnAdmin } from '../../../shared/SharePoint'
import { sp } from 'sp-pnp-js/lib/pnp';


///////////////////////| ExecutiveAssets Page |///////////////////////

export interface IExecutiveAssetsPageProps {
  description: string
  context: PageContext
}

export interface IExecutiveAssetsPageState {
  loading: boolean
  //userIsAdmin: boolean
  breadcrumbItems: any[]
  assetsListDetails: any
  //eventFolderUrl: string
  executiveAssets: IFSObject[]
  executiveAssetsStartFolder: string
  executiveAssetsRootFolderId: number
  startInsideFolder: IFSObject
  showFileSharingDialog: IFSObject
  showEventSharingDialog: IFSObject
  showEventDetailsDialog: boolean
  //eventListDetails: IEventListDetails
  showTextInputDialog: string
  showDeleteConfirmationDialog: IFSObject
  showRenameDialog: IFSObject
  openInBrowser: IFSObject
}

export default class ExecutiveAssetsPage extends React.Component<IExecutiveAssetsPageProps, IExecutiveAssetsPageState> {

  constructor(props) {
    super(props)



    this.state = {
      loading: true,
      //userIsAdmin: false,
      breadcrumbItems: [],
      assetsListDetails: null,
     // eventFolderUrl: "",
      executiveAssets: [],
      executiveAssetsStartFolder: "",
      executiveAssetsRootFolderId: null,
      startInsideFolder: null,
      //eventFiles: [],
      //eventDocumentsStartFolder: "",
      showFileSharingDialog: null,
      showEventSharingDialog: null,
      //eventListDetails: null,
      showEventDetailsDialog: false,
      showTextInputDialog: null,
      showDeleteConfirmationDialog: null,
      showRenameDialog: null,
      openInBrowser: null,
    }
  }

  public async componentDidMount() {
    //onst params = getQueryParameters()
    // console.log("params:", params)
    //if (!params.eventId) return

    //const eventId = parseInt(params.eventId, 10)
    //if (isNaN(eventId)) return

    //const eventDetails = await this.getEventMetadata(eventId)
    //if (!eventDetails) return

    let executiveAssets: IFSObject[] = [] 
    let executiveAssetsStartFolder = siteUrl + '/Exec Assets'
  

    let eventFolderUrl: string
    let startInsideFolder = null
    let executiveAssetsRootFolderId = await sp.web.lists.getByTitle(execAssetsLibraryTitle)
      .rootFolder.get()
      .then((rootFolder) => parseInt(rootFolder.UniqueId))

    const fetchedFiles = await this.fetchFiles(executiveAssetsRootFolderId, false)
    executiveAssets = fetchedFiles.executiveAssets
    //executiveAssetsStartFolder = fetchedFiles.executiveAssetsStartFolder
    //startInsideFolder = fetchedFiles.startInsideFolder
    //eventFolderUrl = fetchedFiles.eventFolderUrl
    //breadcrumb.serverRelativeUrl = eventFolderUrl
    

    //const eventListDetails = await getEventListDetails()
    const userIsAdmin = await isCurrentUserAnAdmin()

    this.setState({
      loading: false,
      //userIsAdmin,
      //breadcrumbItems: [breadcrumb],
      //eventFolderUrl,
      executiveAssets,
      executiveAssetsStartFolder,
      executiveAssetsRootFolderId,
      //startInsideFolder,
      //eventFiles,
      //eventDocumentsStartFolder,
      //eventListDetails,
    })
  }

  @autobind
  private async fetchFiles(folderId?: number, shouldSetState = true) {
   if (!folderId) {
     folderId = this.state.executiveAssetsRootFolderId
   }

    let executiveAssets: IFSObject[] = [] 
    //let executiveAssetsStartFolder = ""
    try {
      executiveAssets = await getFiles(this.state.executiveAssetsStartFolder, execAssetsLibraryTitle)
    } catch (err) {
      // console.log("err:", err)
      executiveAssets = []
    }
    
   

    // How many exec folders are there?
    const executiveFolders = executiveAssets.filter(x => {
      return (
        x.type === "folder" &&
        x.serverRelativeUrl.split("/").length === 6
      )
    })

    // console.log("execFolders:", execFolders)

    let startInsideFolder = null
    if (executiveFolders.length === 1) {
      startInsideFolder = executiveFolders[0]
    }

    if (shouldSetState) {
      this.setState({
        executiveAssets,
        //executiveAssetsStartFolder,
        //startInsideFolder,
        //eventFiles,
        //eventDocumentsStartFolder,
        //eventFolderUrl          
      })
    }

    return {
      executiveAssets,
      //executiveAssetsStartFolder,
      startInsideFolder,
      //eventFiles,
      //eventDocumentsStartFolder,
      //eventFolderUrl
    }
  }

  public render() {
    if (this.state.loading) {
      return <Loading />
    }

    const {
      //userIsAdmin,
      breadcrumbItems,
      //eventDetails,
      executiveAssets,
      executiveAssetsStartFolder,
      startInsideFolder,
      //eventFiles,
      //eventDocumentsStartFolder,
      showFileSharingDialog,
      showEventSharingDialog,
      showEventDetailsDialog,
      showTextInputDialog,
      showDeleteConfirmationDialog,
      showRenameDialog
    } = this.state

    const sharedDocumentsWidthClass = executiveAssets.length ? "ms-u-sm4" : "ms-u-sm12"

    return (
      <div className={styles2.ExecutiveAssetsPage}>
        <div className={styles2.flex}>
          <div className={styles2.flexGrow}>
            <div className={styles2.pageBreadcrumb}>
              <Breadcrumb
                items={[{
                  key: "crumb0",
                  text: "Exec Assets",
                  onClick: () => redirectToSitePage("Assets")
                }]} />
            </div>
            
          </div>
          <div className={styles2.flexColumn}>
          </div>
        </div>
        <div className="ms-Grid">
          <div className="ms-Grid-row">
            {executiveAssets.length > 0 && (
              <div className="ms-Grid-col ms-u-s12">
                <ExecutiveAssets
                  items={executiveAssets}
                  canShare={false}
                  rootFolder={executiveAssetsStartFolder}
                  //startInsideFolder={startInsideFolder}
                  onDeleteFileClick={this.showDeleteItemDialog}
                  onDownloadFileClick={this.downloadFile}
                  onDuplicateFileClick={this.duplicateFileItem}
                  onFileUpload={this.onFileUpload}
                  onNewFolderClick={this.showNewFolderDialog}
                  onOpenFileInBrowserClick={this.openInBrowser}
                  onRenameFileClick={this.showRenameDialog}
                  onShareEventClick={item => this.showDialog("showEventSharingDialog", item)}
                  onShareIconClick={item => this.showDialog("showFileSharingDialog", item)}
                  onSyncToOneDriveClick={this.syncToOneDrive}
                  />
              </div>
            )}
          </div>
        </div>
        {showTextInputDialog && (
          <TextInputDialog
            title="Folder Name"
            value=""
            onDismiss={() => this.hideDialog("showTextInputDialog")}
            onSave={text => this.createNewFolder(text)}
          />
        )}
        {showDeleteConfirmationDialog && (
          <DeleteConfirmationDialog 
            item={showDeleteConfirmationDialog}
            onDismiss={() => this.hideDialog("showDeleteConfirmationDialog")}
            onDelete={() => this.deleteSelectedItem(showDeleteConfirmationDialog)}
          />
        )}
        {showRenameDialog && 
          <TextInputDialog
            title="Rename item"
            value={
              showRenameDialog.name.substring(0, showRenameDialog.name.lastIndexOf('.'))
            }
            extension={showRenameDialog.name.lastIndexOf('.') > -1 ? 
            showRenameDialog.name.substring(showRenameDialog.name.lastIndexOf('.') + 1 ) : ""}
            onDismiss={() => this.hideDialog("showRenameDialog")}
            onSave={text => this.renameSelectedItem(showRenameDialog, text)}
          />
        }
      </div>
    )
  }

  /********************************************************************************************/
  //  Events
  /********************************************************************************************/

  @autobind
  private async syncToOneDrive(currentFolder: string) {
    const item = find(this.state.executiveAssets, x => (
      x.type === "folder" &&
      x.serverRelativeUrl === currentFolder
    ))

    if (!item) {
      throw new Error("item not found")
    }

    const listGUID = await getListGUID(execAssetsLibraryTitle)
    const { context: ctx } = this.props
    const { assetsListDetails } = this.state
    const encode = encodeURIComponent

    let syncHref = (
      "odopen://sync?scope=OPENFOLDER" +
      "&siteId=" + ctx.site.id +
      "&webId=" + ctx.web.id + 
      "&webTitle=" + encode(ctx.web.title) +
      "&listId=" +  listGUID +
      "&listTitle=" + encode(execAssetsLibraryTitle) +
      "&userEmail=" + encode(ctx.user.email) +
      "&listTemplateTypeId=101" +
      "&webUrl=" + encode(ctx.web.absoluteUrl) +
      "&webLogoUrl=" + encode(ctx.web.logoUrl) +
      "&webTemplate=" + encode(ctx.web.templateName) +
      "&isSiteAdmin=" + (ctx.legacyPageContext.isSiteAdmin ? "1":"0") +
      "&folderId=" + item.uniqueFolderId + 
      "&folderName=" + item.name + 
      "&folderUrl=" + encode(currentFolder)
    )
              
    window.open(syncHref, "_new")
  }

  @autobind
  private showRenameDialog(item: IFSObject) {
    this.showDialog("showRenameDialog", item)
  }

  @autobind
  private async renameSelectedItem(item: IFSObject, newName: string) {
    const parentFolder = item.serverRelativeUrl.substring(0, item.serverRelativeUrl.lastIndexOf('/'))
    await renameFile(execAssetsLibraryTitle, item, newName)
    await this.fetchFiles()
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
    console.log("Event Page - createFolder results: " + results)
    await this.fetchFiles()
  }

  @autobind
  private showDeleteItemDialog(item: IFSObject) {
    this.showDialog("showDeleteConfirmationDialog", item)
  }

  @autobind
  private async deleteSelectedItem(item: IFSObject) {
    await deleteItem(item, execAssetsLibraryTitle)
    this.fetchFiles()
  }

  @autobind
  private downloadFile(item: IFSObject) {
    const url = getFileDownloadLink(item);
    window.open(url, '_blank')
  }

  @autobind
  private openInBrowser(item: IFSObject) {
    const url = item.serverRedirectedEmbedUrl.split('&')[0];
    window.open(url, "_blank")
  }

  @autobind
  private async duplicateFileItem(item: IFSObject) {
    await duplicateFile(item)
    this.fetchFiles()
  }
  
  @autobind
  private async onFileUpload(currentFolder: string){
    this.fetchFiles()
  }

  /********************************************************************************************/
  //  Helpers
  /********************************************************************************************/
  
  @autobind
  private showDialog(key: string, data: any = true) {
    this.setState(({ [key]: data }) as any)
  }

  @autobind
  private hideDialog(key: string, data: any = null) {
    this.setState(({ [key]: data }) as any)
  }
}