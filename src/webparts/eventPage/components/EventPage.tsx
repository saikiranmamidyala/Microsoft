import * as React from 'react'

import { Breadcrumb, IBreadcrumb, IBreadcrumbItem } from 'office-ui-fabric-react/lib/Breadcrumb'
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button'
import { registerDefaultFontFaces } from '@uifabric/styling/lib/styles/DefaultFontStyles';
import { Dialog } from 'office-ui-fabric-react/lib/Dialog';

import Loading from '../../../components/Loading'
import ExecDocuments from '../../../components/ExecDocuments'
import EventDocuments from '../../../components/EventDocuments'
import SharingDialog from '../../../components/SharingDialog'
import TextInputDialog from '../../../components/TextInputDialog'
import DeleteConfirmationDialog from '../../../components/DeleteConfirmationDialog'
import EventDetailsDialog from "../../../components/EventDetailsDialog"

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
  eventDocumentLibraryTitle
} from '../../../shared/SharePoint'

import { getQueryParameters } from '../../../shared/util'
import { PageContext } from '@microsoft/sp-page-context'
import styles from './EventPage.module.scss'
import { find } from 'lodash'
import { autobind, css } from 'office-ui-fabric-react/lib/Utilities'
import { isCurrentUserAnAdmin } from '../../../shared/SharePoint'

///////////////////////| EventMetadata |///////////////////////

class EventMetadata extends React.PureComponent<IEvent> {
  public render() {
    const {
      startDate,
      endDate,
      location,
      commsManagers
    } = this.props

    return (
      <div className={styles.metadata}>
        {startDate && (
          <div>
            <div className={`${styles.metadataLabel} ms-font-m`}>
              Start Date
            </div>
            <div className="ms-font-m-plus">
              {startDate.format("YYYY-MM-DD")}
            </div>
          </div>
        )}
        {endDate && (
          <div>
            <div className={`${styles.metadataLabel} ms-font-m`}>
              End Date
            </div>
            <div className="ms-font-m-plus">
              {endDate.format("YYYY-MM-DD")}
            </div>
          </div>
        )}
        {location && (
          <div>
            <div className={`${styles.metadataLabel} ms-font-m`}>
              Location
            </div>
            <div className="ms-font-m-plus">
              {location}
            </div>
          </div>
        )}
        {commsManagers.length > 0 && (
          <div>
            <div className={`${styles.metadataLabel} ms-font-m`}>
              Comms Manager(s)
            </div>
            <div className="ms-font-m-plus">
              {commsManagers.map(x => x.Title).join(", ")}
            </div>
          </div>
        )}
      </div>    
    )
  }
}

///////////////////////| EventPage |///////////////////////

export interface IEventPageProps {
  description: string;
  context: PageContext
}

export interface IEventPageState {
  loading: boolean
  userIsAdmin: boolean
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

export default class EventPage extends React.Component<IEventPageProps, IEventPageState> {

  constructor(props) {
    super(props)

    this.state = {
      loading: true,
      userIsAdmin: false,
      breadcrumbItems: [],
      eventDetails: null,
      eventFolderUrl: "",
      execFiles: [],
      execDocumentsStartFolder: "",
      startInsideFolder: null,
      eventFiles: [],
      eventDocumentsStartFolder: "",
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

  private async getEventMetadata(eventId: number) {
    let eventDetails: IEvent
    try {
      eventDetails = await getEventById(eventId)
    } catch (err) {
      // console.log("err:", err)
      eventDetails = null
    }
    // console.log("eventDetails:", eventDetails)

    return eventDetails
  }

  public async componentDidMount() {
    const params = getQueryParameters()
    // console.log("params:", params)
    if (!params.eventId) return

    const eventId = parseInt(params.eventId, 10)
    if (isNaN(eventId)) return

    const eventDetails = await this.getEventMetadata(eventId)
    if (!eventDetails) return

    let execFiles: IFSObject[] = [] 
    let execDocumentsStartFolder = ""
    let eventFiles: IFSObject[] = []
    let eventDocumentsStartFolder = ""

    let breadcrumb = {
      text: eventDetails.eventName,
      serverRelativeUrl: ""
    }

    let eventFolderUrl: string
    let startInsideFolder = null

    if (eventDetails.folderId) {
      const fetchedFiles = await this.fetchFiles(eventDetails.folderId, false)
      execFiles = fetchedFiles.execFiles
      execDocumentsStartFolder = fetchedFiles.execDocumentsStartFolder
      startInsideFolder = fetchedFiles.startInsideFolder
      eventFiles = fetchedFiles.eventFiles
      eventDocumentsStartFolder = fetchedFiles.eventDocumentsStartFolder
      eventFolderUrl = fetchedFiles.eventFolderUrl
      breadcrumb.serverRelativeUrl = eventFolderUrl
    }

    const eventListDetails = await getEventListDetails()
    const userIsAdmin = await isCurrentUserAnAdmin()

    this.setState({
      loading: false,
      userIsAdmin,
      breadcrumbItems: [breadcrumb],
      eventDetails,
      eventFolderUrl,
      execFiles,
      execDocumentsStartFolder,
      startInsideFolder,
      eventFiles,
      eventDocumentsStartFolder,
      eventListDetails,
    })
  }

  @autobind
  private async fetchFiles(folderId?: number, shouldSetState = true) {
    if (!folderId) {
      folderId = this.state.eventDetails.folderId
    }

    let execFiles: IFSObject[] = [] 
    let execDocumentsStartFolder = ""
    let eventFiles: IFSObject[] = []
    let eventDocumentsStartFolder = ""
    let eventFolderUrl: string

    try {
      eventFolderUrl = await getFolderUrl(folderId, eventDocumentLibraryTitle)
    } catch (err) {
      // console.log("err:", err)
      eventFolderUrl = ""
    }
    // console.log("eventFolderUrl:", eventFolderUrl)
    if (!eventFolderUrl) return

    execDocumentsStartFolder = eventFolderUrl

    try {
      execFiles = await getFiles(eventFolderUrl, eventDocumentLibraryTitle)
    } catch (err) {
      // console.log("err:", err)
      execFiles = []
    }
    
    const sharedFolderUrl = `${eventFolderUrl.toLowerCase()}/shared`
    const sharedFolderUrlPaths = sharedFolderUrl.split("/")

    for (let i = execFiles.length - 1; i >= 0; i -= 1) {
      const eventFile = execFiles[i]

      // Is this a shared file or folder?
      if (eventFile.serverRelativeUrl.toLowerCase().indexOf(sharedFolderUrl) >= 0) {
        if (eventFile.serverRelativeUrl.split("/").length === sharedFolderUrlPaths.length) {
          // `eventFile` is the Shared folder itself
          eventDocumentsStartFolder = eventFile.serverRelativeUrl
          ;(eventFile as any).isSharedFolder = true
        }
        // Move shared files into shared files array
        execFiles.splice(i, 1)
        eventFiles.push(eventFile)
      }
    }

    // How many exec folders are there?
    const execFolders = execFiles.filter(x => {
      return (
        x.type === "folder" &&
        x.serverRelativeUrl.split("/").length === 6
      )
    })

    // console.log("execFolders:", execFolders)

    let startInsideFolder = null
   // directly jump to exec folder if only one exec.
    // if (execFolders.length === 1) {
    //   startInsideFolder = execFolders[0]
    // }

    if (shouldSetState) {
      this.setState({
        execFiles,
        execDocumentsStartFolder,
        startInsideFolder,
        eventFiles,
        eventDocumentsStartFolder,
        eventFolderUrl          
      })
    }

    return {
      execFiles,
      execDocumentsStartFolder,
      startInsideFolder,
      eventFiles,
      eventDocumentsStartFolder,
      eventFolderUrl
    }
  }

  public render() {
    if (this.state.loading) {
      return <Loading />
    }

    const {
      userIsAdmin,
      breadcrumbItems,
      eventDetails,
      execFiles,
      execDocumentsStartFolder,
      startInsideFolder,
      eventFiles,
      eventDocumentsStartFolder,
      showFileSharingDialog,
      showEventSharingDialog,
      showEventDetailsDialog,
      showTextInputDialog,
      showDeleteConfirmationDialog,
      showRenameDialog
    } = this.state

    const sharedDocumentsWidthClass = execFiles.length ? "ms-u-sm4" : "ms-u-sm12"

    return (
      <div className={styles.EventPage}>
        <div className={styles.flex}>
          <div className={styles.flexGrow}>
            <div className={styles.pageBreadcrumb}>
              <Breadcrumb
                items={[{
                  key: "crumb0",
                  text: "Events",
                  onClick: () => redirectToSitePage("FutureEvents")
                },{
                  key: "crumb1",
                  text: eventDetails.eventName,
                  isCurrentItem: true
                }]} />
            </div>
            <EventMetadata {...eventDetails} />
          </div>
          <div className={styles.flexColumn}>
            <PrimaryButton
              onClick={async () => {
                const icsLink = await getEventICSLink(eventDetails.id)
                window.open(icsLink, "_New")    
              }}
              iconProps={{ iconName: "Add" }}
            /* PrimaryButton */>
              OUTLOOK
            </PrimaryButton>
            {userIsAdmin && (
              <PrimaryButton
                onClick={() => this.showDialog("showEventDetailsDialog")}
                iconProps={{ iconName: "Edit" }}
              /* PrimaryButton */>
                EVENT
              </PrimaryButton>
            )}
          </div>
        </div>
        <div className="ms-Grid">
          <div className="ms-Grid-row">
            {execFiles.length > 0 && (
              <div className="ms-Grid-col ms-u-sm8">
                <ExecDocuments
                  items={execFiles}
                  rootFolder={execDocumentsStartFolder}
                  startInsideFolder={startInsideFolder}
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
            <div className={css("ms-Grid-col", sharedDocumentsWidthClass)}>
              <EventDocuments
                items={eventFiles}
                canShare={execFiles.length > 0}
                rootFolder={eventDocumentsStartFolder} 
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
          </div>
        </div>
        {(showFileSharingDialog || showEventSharingDialog) && (
          <SharingDialog
            event={eventDetails}
            file={showFileSharingDialog}
            folder={showEventSharingDialog}
            onDismiss={() => {
              this.fetchFiles()
              if (showFileSharingDialog) {
                this.hideDialog("showFileSharingDialog")
              } else {
                this.hideDialog("showEventSharingDialog")
              }
            }} />
        )}
        {showEventDetailsDialog && (
          <EventDetailsDialog
            eventId={eventDetails.id}
            onDismiss={() => this.hideDialog("showEventDetailsDialog")}
            onSuccess={async eventId => {
              this.hideDialog("showEventDetailsDialog")
              // TODO: Update events to show updated event
              const metadata = await this.getEventMetadata(eventDetails.id)
              this.setState({ eventDetails: metadata })
            }}
          />
        )}
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
  private syncToOneDrive(currentFolder: string) {
    const item = find(this.state.execFiles, x => (
      x.type === "folder" &&
      x.serverRelativeUrl === currentFolder
    ))

    if (!item) {
      throw new Error("item not found")
    }

    const { context: ctx } = this.props
    const { eventListDetails } = this.state
    const encode = encodeURIComponent

    let syncHref = (
      "odopen://sync?scope=OPENFOLDER" +
      "&siteId=" + ctx.site.id +
      "&webId=" + ctx.web.id + 
      "&webTitle=" + encode(ctx.web.title) +
      "&listId=" +  eventListDetails.id +
      "&listTitle=" + encode(eventListDetails.name) +
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
    await renameFile(eventDocumentLibraryTitle,item, newName)
    // const files = await getFiles(parentFolder)
    // this.refreshList(parentFolder, files)
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
    const results = await createFolder(currentFolder, newFolderName, eventDocumentLibraryTitle)
    console.log("Event Page - createFolder results: " + results)
    //const files = await getFiles(currentFolder)
    //this.refreshList(currentFolder, files)
    await this.fetchFiles()
  }

  // @autobind
  // private refreshList(currentFolder: string, files: IFSObject[]) {
  //   if (currentFolder.indexOf(this.state.execDocumentsStartFolder + "/Shared") > -1) {
  //     this.setState({ eventFiles : files })
  //   } else {
  //     this.setState({ execFiles: files })
  //   }
  // }

  @autobind
  private showDeleteItemDialog(item: IFSObject) {
    this.showDialog("showDeleteConfirmationDialog", item)
  }

  @autobind
  private async deleteSelectedItem(item: IFSObject) {
    await deleteItem(item, eventDocumentLibraryTitle)
    this.fetchFiles()
    // this.setState({
    //   execFiles: this.state.execFiles.filter(e => e !== item),
    //   eventFiles: this.state.eventFiles.filter(e => e !== item)
    // })
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
    // let currentFolder = item.serverRelativeUrl.substring(0, item.serverRelativeUrl.lastIndexOf("/"))
    // const files = await getFiles(currentFolder)
    // this.refreshList(currentFolder, files)
    this.fetchFiles()
  }
  
  @autobind
  private async onFileUpload(currentFolder: string){
    // const files = await getFiles(currentFolder)
    // this.refreshList(currentFolder, files)
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