import * as moment from 'moment'

import pnp, { 
  CamlQuery,
  HttpClient, 
  Web, 
  WebEnsureUserResult, 
  sp, 
  Folder, 
  Item, 
  PermissionKind,
  FolderAddResult, 
  FileAddResult, 
  ChunkedFileUploadProgressData,
  SharingRole,
  TypedHash,
  ContextInfo,
  ItemUpdateResult,
  ItemAddResult,
  Items,
  ODataQueryable,
  MoveOperations,
  Files
} from 'sp-pnp-js'

import { FolderUpdateResult } from 'sp-pnp-js/lib/sharepoint/folders';
import { RoleDefinition } from 'sp-pnp-js/lib/sharepoint/roles';
import { SiteUser, CurrentUser } from 'sp-pnp-js/lib/sharepoint/siteusers';
import { SiteGroup } from 'sp-pnp-js/lib/sharepoint/sitegroups';
import { IPersona, IPersonaProps } from 'office-ui-fabric-react/lib/components/Persona'
import { createQueryParameters } from './util'
import { uniqBy, find } from 'lodash'
//const build = require('@microsoft/sp-build-web');
//export const cdnAssetsBaseUrl = "https://publiccdn.sharepointonline.com/microsoft.sharepoint.com/teams/MSITAppCDN/SiteAssets/ExecComms"
//export const cdnAssetsBaseUrl = "https://blulinksolutions.sharepoint.com/sites/ExecCommsvNext/App"
export const cdnAssetsBaseUrl = "https://microsoft.sharepoint.com/teams/dev_execcomms/cdn"
//export const cdnAssetsBaseUrl = build.writeManifests.taskConfig.cdnBasePath;
//export const cdnAssetsBaseUrl = "https://microsoft.sharepoint.com/teams/ExecComms656/cdn"
// `siteDomain` example: https://blulinksolutions.sharepoint.com
export const siteDomain = location.origin
// `siteUrl` example: /sites/ExecCommsvNext
export const siteUrl = location.pathname.split("/").splice(0, 3).join("/")
// `siteCollectionUrl` example: https://blulinksolutions.sharepoint.com/sites/ExecCommsvNext
export const siteCollectionUrl = siteDomain + siteUrl

//export const env = siteUrl.indexOf("dev") >= 0 ? "dev" : "test"

//export const archiveSiteUrl = "/sites/execcomms2"
//export const archiveSiteCollectionUrl = siteDomain + archiveSiteUrl
export const archiveSiteCollectionUrl = "https://microsoft.sharepoint.com/teams/ExecComms"
const archiveWeb = new Web(archiveSiteCollectionUrl)

export const eventDocumentLibraryTitle: string = "Event Documents";

export const execAssetsLibraryTitle: string = "Exec Assets"

const documentLibraryTitle = "Event Documents"
const calendarTitle = "Executive Events"




// This is for debugging and allows you to execute
// requests to SharePoint in the browser dev tools
;(window as any).sp = sp

export function redirectToSitePage(name: string, params = {}) {
  location.assign(`${siteCollectionUrl}/sitepages/${name}.aspx${createQueryParameters(params)}`)
}

export async function getEventICSLink(itemID: number): Promise<string> {
  const listGUID =  await getListGUID(calendarTitle)
  return `${siteCollectionUrl}/_vti_bin/owssvr.dll?CS=109&Cmd=Display&List=${listGUID}&CacheControl=1&ID=${itemID}&Using=event.ics`
}

export function getFileDownloadLink(item: IFSObject) : string {
  return `${siteCollectionUrl}/_layouts/download.aspx?SourceUrl=${item.serverRelativeUrl}`
}

export async function getListGUID(listTitle: string): Promise<string> {
  return (
    sp.web
      .lists.getByTitle(`${listTitle}`).get()
      .then((list) => {
        return list.Id;
      })
  )
}

///////////////////////////////////////////////////////////////
//
//  Interfaces and Classes
//
///////////////////////////////////////////////////////////////

export interface IEventListDetails {
  id: string
  name: string
  url: string
}

export interface IExecutive {
  id: number
  name: string
  groupId: number
  archiveGroupId: number
  initials: string
  imageUrl: string
}

export interface IEvent {
  id: number
  eventName: string
  startDate: moment.Moment
  endDate: moment.Moment
  location: string
  principalIds: number[]
  execs?: IExecutive[]
  commsManagers?: any[]
  folderId?: number
  externalLink?: string
}

export interface IArchivedEvent extends IEvent {
  executiveName: string
}

export interface IArchiveExecConfig {
  execName: string,
  libraryUrl: string,
  execId?: number
}

export interface IFSObject {
  name: string
  type: "file" | "folder"
  id: number
  serverRelativeUrl: string
  size: number
  modifiedBy: string
  modified: moment.Moment
  uniqueFolderId: string
  uniqueFileId: string
  hasUniquePermissions: boolean
  serverRedirectedEmbedUrl: string,
  directAccessUsers: number,
  fileExtension?: string,
  isDocument?: boolean,
  isContainer?: boolean,
  author?: string
}

export interface IUserEntityData {
  IsAltSecIdPresent: string;
  ObjectId: string;
  Title: string;
  Email: string;
  MobilePhone: string;
  OtherMails: string;
  Department: string;
}

export interface IClientPeoplePickerSearchUser {
  Key: string;
  Description: string;
  DisplayText: string;
  EntityType: string;
  ProviderDisplayName: string;
  ProviderName: string;
  IsResolved: boolean;
  EntityData: IUserEntityData;
  MultipleMatches: any[];
}

export interface IEnsureUser {
  Email: string;
  Id: number;
  IsEmailAuthenticationGuestUser: boolean;
  IsHiddenInUI: boolean;
  IsShareByEmailGuestUser: boolean;
  IsSiteAdmin: boolean;
  LoginName: string;
  PrincipalType: number;
  Title: string;
  UserId: {
      NameId: string;
      NameIdIssuer: string;
  };
}

export interface IEnsurableSharePointUser
    extends IClientPeoplePickerSearchUser, IEnsureUser {}

export class SharePointUserPersona implements IPersona {
    private _user: IEnsurableSharePointUser;

    public get User(): IEnsurableSharePointUser {
      return this._user;
    }

    public set User(user: IEnsurableSharePointUser) {
      this._user = user;
      this.primaryText = user.Title;
      this.secondaryText = user.EntityData.Title;
      this.tertiaryText = user.EntityData.Department;
      this.imageShouldFadeIn = true;
      this.imageUrl = `/_layouts/15/userphoto.aspx?size=S&accountname=${this.User.Key.substr(this.User.Key.lastIndexOf('|') + 1)}`;
    }

    constructor (user: IEnsurableSharePointUser) {
      this.User = user;
    }

    public primaryText: string;
    public secondaryText: string;
    public tertiaryText: string;
    public imageUrl: string;
    public imageShouldFadeIn: boolean;
}

export interface IEventFormValues {
  id?: number
  execs?: IExecutive[]
  eventName?: string
  startDate?: Date
  endDate?: Date
  location?: string
  commsManagers?: IPersonaProps[]
}

export interface ICreateNewEventParams {
  event: IEventFormValues
  progress?: (message: string) => void
}

interface ICreateNewEventFoldersResult {
  eventFolderItemId: number
  sharedFolderItemId: number
  execFolders: Array<{
    execId: number
    itemId: number
  }>
}

interface ICreateNewEventGroupsResult {
  sharedGroupId: number
  execGroups: Array<{
    execId: number
    groupId: number
  }>
}

export interface ICreateGroupOpts extends TypedHash<any> {
  Title: string
  Description: string
  AllowMembersEditMembership: boolean
  AllowRequestToJoinLeave: boolean
  AutoAcceptRequestToJoinLeave: boolean
  OnlyAllowMembersViewMembership: boolean
}

export enum PrincipalType {
  user,
  group
}

export interface IUser {
  type: PrincipalType
  id: number
  email: string
  name: string
  loginName: string
}

export interface IGroup {
  type: PrincipalType
  id: number
  name: string
  principals: IPrincipal[]
}

export interface IPrincipal extends IUser, IGroup {}

export interface IUser {
  Id: number
  Title: string
  Email: string
}

export interface ISiteGroup {
  Id: number
  Title: string
  Users?: IUser[]
  Description?: string
}

interface UniqueRoleAssignment {
  groupIds: number[]
  roleDefIds: number[]
}

///////////////////////////////////////////////////////////////
//
//  Functions
//
///////////////////////////////////////////////////////////////

export async function getEventListDetails(): Promise<IEventListDetails> {
  const data = await sp.web.lists.getByTitle(eventDocumentLibraryTitle).get()

  return {
    id: data.Id,
    name: data.Title,
    url: `${data.ParentWebUrl}/${data.Title}`
  } as IEventListDetails
}

export async function getAllExecutives(): Promise<IExecutive[]> {
  let items

  try {
    items = await sp.web.lists
      .getByTitle("Executives").items
      .orderBy('Title')
      .get()
  } catch (err) {
    items = []
  }
    
  return items.map(item => ({
    id: item.Id,
    name: item.Title,
    groupId: item.GroupId || -1,
    archiveGroupId: item.ArchiveGroupId || -1,
    initials: (
      (item.Title as string)
        .split(" ")
        .map(name => name.substring(0, 1).toUpperCase())
        .join("")
    ),
    imageUrl: (item.ImageUrl && (item.ImageUrl.Url || item.ImageUrl)) || ""
  }) as IExecutive)
}

export async function getFutureEvents(): Promise<IEvent[]> {
  const dateFilter = moment().startOf("day").subtract(7, "days").toISOString()
  let items 

  try {
    // First we try requesting `RoleAssignments`. If the user does not
    // have the access to enumerate permissions, we make the request
    // again (below) without `RoleAssignments`    
    items = await sp.web
      .lists.getByTitle(calendarTitle)
      .items
      .expand(
        "RoleAssignments",
        "RoleAssignments/Member",
        "RoleAssignments/RoleDefinitionBindings"
      )
      .filter(`EventDate ge datetime'${dateFilter}' and EventFiles ne null`)
      .get()

  } catch (err) {
    items = await sp.web.lists
    .getByTitle(calendarTitle).items
    // .expand(
    //   "RoleAssignments",
    //   "RoleAssignments/Member",
    //   "RoleAssignments/RoleDefinitionBindings"
    // )
    .filter(`EventDate ge datetime'${dateFilter}' and EventFiles ne null`)
    .get()
  }
    
  if (!items || !items.length) {
    return []
  }

  return items.map(item => ({
    id: item.Id,
    eventName: item.Title,
    startDate: item.EventDate ? moment(item.EventDate) : "",
    endDate: item.EndDate ? moment(item.EndDate) : "",
    location: item.Location || "",
    principalIds: (
      item.RoleAssignments &&
      item.RoleAssignments.map(roleAssign => roleAssign.Member.Id)
    ) || [],
  }) as IEvent)

    
}    

export async function getPastEvents(): Promise<IEvent[]> {
  const dateFilter = moment().startOf("day").subtract(7, "days").toISOString()
  let items

  try {
    // First we try requesting `RoleAssignments`. If the user does not
    // have the access to enumerate permissions, we make the request
    // again (below) without `RoleAssignments`        
    items = await sp.web
      .lists.getByTitle(calendarTitle)
      .items
      .expand(
        "RoleAssignments",
        "RoleAssignments/Member",
        "RoleAssignments/RoleDefinitionBindings"
      )
      .filter(`EventDate lt datetime'${dateFilter}' and EventFiles ne null`)
      .get()

  } catch (err) {
    items = await sp.web
      .lists.getByTitle(calendarTitle)
      .items
      // .expand(
      //   "RoleAssignments",
      //   "RoleAssignments/Member",
      //   "RoleAssignments/RoleDefinitionBindings"
      // )
      .filter(`EventDate lt datetime'${dateFilter}' and EventFiles ne null`)
      .get()
  }

  if (!items || !items.length) {
    return []
  }

  return items.map(item => ({
    id: item.Id,
    eventName: item.Title,
    startDate: item.EventDate ? moment(item.EventDate) : "",
    endDate: item.EndDate ? moment(item.EndDate) : "",
    location: item.Location || "",
    principalIds: (
      item.RoleAssignments &&
      item.RoleAssignments.map(roleAssign => roleAssign.Member.Id)
    ) || []
  }) as IEvent)
}  

export async function getArchivedEvents(): Promise<IEvent[]> {
  let events
  
  try {
    events = await archiveWeb.lists
      .getByTitle(calendarTitle).items
      .select(
        "Id",
        "ecEventName",
        "ecStartDate", 
        "ecEndDate",
        "Location",
        "Venue",
        "ecExecNameLookup/Title",
        "RoleAssignments"
      )
      .expand(
          "RoleAssignments",
          "RoleAssignments/Member",
          "RoleAssignments/RoleDefinitionBindings",
          "ecExecNameLookup/Title"
      )
      .get()

  } catch (err) {
    events = null
  }

  if (!events) {
    try {
      events = await archiveWeb.lists
      .getByTitle(calendarTitle).items
      .select(
        "Id",
        "ecEventName",
        "ecStartDate", 
        "ecEndDate",
        "Location",
        "Venue",
        "ecExecNameLookup/Title",
        //"RoleAssignments"
      )
      .expand(
          // "RoleAssignments",
          // "RoleAssignments/Member",
          // "RoleAssignments/RoleDefinitionBindings",
          "ecExecNameLookup/Title"
      )
      .get()
    } catch (err) {
      events = null
    }
  }

  if (!events || !events.length) {
    return []
  }

  return events.map(event => ({
    id: event.Id,
    eventName: event.ecEventName,
    startDate: event.ecStartDate ? moment(event.ecStartDate) : "",
    endDate: event.ecEndDate ? moment(event.ecEndDate) : "",
    location: event.Location || event.Venue || "",
    executiveName: event.ecExecNameLookup ? event.ecExecNameLookup.Title : "",
    //if there is no executive name (ecExecNameLookup, then it is a Key Event and is in a different library)
    externalLink: (
      event.ecExecNameLookup
        ? "https://microsoft.sharepoint.com/teams/ExecComms/" + (event.ecExecNameLookup.Title.replace('-', '')) + "/Forms/Exec%20Comms%20Document%20Set/docsethomepage.aspx?RootFolder=%2Fteams/execcomms/" + (event.ecExecNameLookup.Title.replace('-', '')) + "/" + event.ecEventName
        : "https://microsoft.sharepoint.com/teams/ExecComms/teams/ExecComms/Exec%20Comm%20Key%20Events/Forms/Exec%20Comm%20Key%20Events/docsethomepage.aspx?RootFolder=/teams/ExecComms/Exec%20Comm%20Key%20Events/" + event.ecEventName
    ),
    principalIds: (
      event.RoleAssignments && 
      event.RoleAssignments.map(roleAsign => roleAsign.PrincipalId)
    ) || []
  }) as IArchivedEvent) 
}

export async function getArchivedExecutiveConfig(): Promise<IArchiveExecConfig[]> {
  let items

  try {
    items = await archiveWeb.lists
      .getByTitle("Executive Config").items
      .get()
  } catch (err) {
    items = []
  }

  return items.map(item => ({
    execName: item.Title,
    libraryUrl: item.LibraryURL
  }) as IArchiveExecConfig)
}

export async function getEventById(id: number): Promise<IEvent> {
  let canEnumeratePerms = false
  try {
    canEnumeratePerms = await sp.web.lists
      .getByTitle(calendarTitle)
      .items.getById(id)
      .currentUserHasPermissions(PermissionKind.EnumeratePermissions)
  } catch (err) {
    canEnumeratePerms = false
  }
  console.log("canEnumeratePerms:", canEnumeratePerms)

  let query = sp.web.lists
    .getByTitle(calendarTitle)
    .items.getById(id)
    .select(
      "*",
      "CommsManagers/Id",
      "CommsManagers/Title",
      "CommsManagers/EMail",
      //"CommsManagers/JobTitle",
     // "CommsManagers/Department",
    )
   .expand("CommsManagers")

  if (canEnumeratePerms) {
    query = query.expand(
      "CommsManagers", 
      //"CommsManagers/Id",
      //"CommsManagers/Title",
      //"CommsManagers/EMail",
      //"CommsManagers/JobTitle",
      //"CommsManagers/Department"
      "RoleAssignments",
      "RoleAssignments/Member",
      "RoleAssignments/RoleDefinitionBindings",
      "RoleAssignments/Title"
    )
  }
  
  const evt = await query.get()
    
  return {
    id: evt.Id,
    eventName: evt.Title,
    startDate: evt.EventDate ? moment(evt.EventDate) : null,
    endDate: evt.EndDate ? moment(evt.EndDate) : null,
    location: evt.Location || "",
    folderId: evt.EventFilesId,
    commsManagers: (
      evt.CommsManagers && ((
        evt.CommsManagers.results &&
        evt.CommsManagers.results
      ) || (
        evt.CommsManagers.length &&
        evt.CommsManagers
      ))
    ) || [],
    principalIds: (
      evt.RoleAssignments && ((
        evt.RoleAssignments.results &&
        evt.RoleAssignments.results.map(roleAssign => roleAssign.Member.Id)
      ) || (
        evt.RoleAssignments.length &&
        evt.RoleAssignments.map(roleAssign => roleAssign.Member.Id)
      ))
    ) || []
  } as IEvent
}

export async function createFolder(currentFolderPath: string, newFolderName: string, listName: string): Promise<FolderAddResult> { 
  return await sp.web.lists.getByTitle(listName)
    .rootFolder.folders.add(currentFolderPath +  "/" + newFolderName)
}

export async function uploadFile(currentFolderPath: string, filePath: string, fileData: File ) : Promise<FileAddResult> {
    return await sp.web.getFolderByServerRelativeUrl(currentFolderPath).files.add(filePath, fileData, true)
 }

export async function renameFile(listTitle: string, item: IFSObject, newName: string) : Promise<ItemUpdateResult> {
  return await sp.web.lists
    .getByTitle(listTitle).items
    .getById(item.id)
    .update({ FileLeafRef: newName })
 }
 
export async function duplicateFile(item: IFSObject) : Promise<void> {
  const newItemPath = `${item.serverRelativeUrl.substring(0, item.serverRelativeUrl.lastIndexOf('.'))}1${item.serverRelativeUrl.substring(item.serverRelativeUrl.lastIndexOf('.'))}`  

  return await sp.web.getFileByServerRelativeUrl(item.serverRelativeUrl)
    .copyTo(newItemPath, false)
 }
 //dropzone
 export async function isFileNameExists (destinationPath: string, fileName: string): Promise<any>{
   var fileExists=false;
  await pnp.sp.web.getFolderByServerRelativeUrl(destinationPath).files.getByName(fileName).get()
  .then((fileInfo: any) => {
    fileExists= true;
    
  })
  .catch((error: any) => {
    fileExists= false;
  })
  if(fileExists)
  {
   var userConformation= confirm("A file named "+fileName+" already exists in this library.do you want to replace file?")
   if(userConformation)
      return true;
    else
      return false;
  }
  else
    return true;
}
export async function addFile(currentFolderPath: string, filePath: string, fileData: File ) : Promise<FileAddResult> {
  if(await isFileNameExists(currentFolderPath,filePath))
  {
    return await sp.web.getFolderByServerRelativeUrl(currentFolderPath).files.add(filePath, fileData, true);
  }
}

 export async function moveFile(destinationPath:string, item: IFSObject) : Promise<void> {
    const copyDestinationPath=`${destinationPath}/${item.name}`;
    if(await isFileNameExists(destinationPath,item.name))
    {
      return await sp.web.getFileByServerRelativeUrl(item.serverRelativeUrl)
      .moveTo(copyDestinationPath,1);
    }  
 }
 export async function shareFile(sharePath:string,item: IFSObject) : Promise<void> {
  const newSharedPath = `${sharePath}/${item.name}`;
  if(await isFileNameExists(sharePath,item.name))
  {
    return await sp.web.getFileByServerRelativeUrl(item.serverRelativeUrl)
      .copyTo(newSharedPath,true)
      .then((res) => console.log(res) )
  }
 }
 //end
export async function getFolderUrl(folderId: number, listTitle: string): Promise<string> {
  const folder = await sp.web.lists
    .getByTitle(listTitle)
    .items.getById(folderId)
    .expand("Folder")
    .get()
    
  return (folder.Folder && folder.Folder.ServerRelativeUrl) || null
}

export async function getFiles(folderUrl: string, listTitle: string): Promise<IFSObject[]> {
let items
const folderTitle = folderUrl.split('/').pop()
const top = 5000;
try {
// First we try requesting `RoleAssignments`. If the user does not
// have the access to enumerate permissions, we make the request
// again (below) without `RoleAssignments` 
// add if else for filtering data according to folderTitle with respective list
if(listTitle=="Event Documents")
  {
    items = await sp.web.lists.getByTitle(listTitle)
    .items
    .filter(
    `substringof('${folderTitle}', FileRef)`
    )
    .select(
    "*",
    "File/ModifiedBy/Title",
    "HasUniqueRoleAssignments",
    "ServerRedirectedEmbedUrl",
    "FileRef",
    )
    .expand(
    "File",
    "File/Name",
    "File/UniqueId",
    "File/ModifiedBy",
    "File/ModifiedBy/Title",
    "File/ServerRelativeUrl",
    "File/Length",
    "Folder",
    "Folder/ServerRelativeUrl",
    "RoleAssignments",
    "RoleAssignments/Member",
    "RoleAssignments/RoleDefinitionBindings",
    )
    .top(top)
    .get();
  }
  else
  {
    items = await sp.web.lists.getByTitle(listTitle)
    .items
    .select(
    "*",
    "File/ModifiedBy/Title",
    "HasUniqueRoleAssignments",
    "ServerRedirectedEmbedUrl",
    "FileRef",
    )
    .expand(
    "File",
    "File/Name",
    "File/UniqueId",
    "File/ModifiedBy",
    "File/ModifiedBy/Title",
    "File/ServerRelativeUrl",
    "File/Length",
    "Folder",
    "Folder/ServerRelativeUrl",
    "RoleAssignments",
    "RoleAssignments/Member",
    "RoleAssignments/RoleDefinitionBindings",
    )
    .top(top)
    .get();
  }
} catch (err) {
items = await sp.web.lists
.getByTitle("Event Documents")
.items
.filter(
`substringof('${folderTitle}', FileRef)`
)
.select(
"*",
"File/ModifiedBy/Title",
"HasUniqueRoleAssignments",
"ServerRedirectedEmbedUrl",
"FileRef",
)
.expand(
"File",
"File/Name",
"File/UniqueId",
"File/ModifiedBy",
"File/ModifiedBy/Title",
"File/ServerRelativeUrl",
"File/Length",
"Folder",
)
.top(top)
.get();
}
if (!items || !items.length) {
return []
}
return items.map(item => ({
id: item.Id,
name: item.Folder ? item.Folder.Name : item.File.Name,
type: item.Folder ? "folder" : "file",
uniqueFolderId: item.Folder ? item.Folder.UniqueId : "",
uniqueFileId: item.File ? item.File.UniqueId : "",
serverRelativeUrl: item.Folder ? item.Folder.ServerRelativeUrl : item.File.ServerRelativeUrl,
size: (item.File && item.File.Length) ? parseInt(item.File.Length, 10) : 0,
modifiedBy: (item.File && item.File.ModifiedBy && item.File.ModifiedBy.Title) || "",
modified: item.Modified ? moment(item.Modified) : "",
hasUniquePermissions: item.HasUniqueRoleAssignments,
serverRedirectedEmbedUrl: item.ServerRedirectedEmbedUrl,
// Count how many users have direct access
directAccessUsers: (
item.File &&
item.RoleAssignments &&
item.RoleAssignments.length &&
item.RoleAssignments.reduce((count, roleAssign) => {
if (
roleAssign.Member &&
roleAssign.Member["odata.type"] &&
roleAssign.Member["odata.type"] === "SP.User"
) {
return count + 1
} else {
return count
}
}, 0)
) || 0
}) as IFSObject)
}


export async function deleteItem (item: IFSObject, listTitle: string): Promise<string>{
  if (item.type == "file") {
    return await sp.web.lists
      .getByTitle(listTitle).items
      .getById(item.id)
      .recycle()
  } else {
    return await sp.web
      .getFolderByServerRelativeUrl(item.serverRelativeUrl)
      .recycle()
  }
}

////////////////////| PeoplePicker section |////////////////////

export async function searchPeople(text: string) : Promise<IPersonaProps[]> {
  const http = new HttpClient()

  const res = await http.post(`${siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`, {
    headers: { "Accept": "application/json;odata=nometadata"},
    body: JSON.stringify({
      queryParams: {
        __metadata: { type: "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters" },
        AllowEmailAddresses: true,
        AllowMultipleEntities: false,
        AllUrlZones: false,
        MaximumEntitySuggestions: 50,
        PrincipalSource: 15,
        PrincipalType: 15,
        QueryString: text
      }
    })
  })

  const response = await res.json()
  const queryResults: IClientPeoplePickerSearchUser[] = JSON.parse(response.value)

  const persons = queryResults
    .map(x => new SharePointUserPersona(x as IEnsurableSharePointUser))
    .filter(x => x.User.EntityType === "User")

  const batch = sp.web.createBatch()

  const batchPromises = persons.map(person => {
    return new Promise(batchResolve => {
      sp.web.inBatch(batch)
        .ensureUser(person.User.Key)
        .then((result: WebEnsureUserResult) => {
          batchResolve(result.data as IEnsureUser)
        })
        .catch(err => {
          batchResolve(null)
        })
    })
  })

  await batch.execute()

  const userResults = await Promise.all(batchPromises)

  userResults.forEach((result: IEnsureUser) => {
    if (result) {
      const persona = find(persons, p => p.User.Key === result.LoginName)
      if (persona && persona.User) {
        persona.User = {
          ...persona.User,
          ...result
        }
      }  
    }
  })

  return persons
}

////////////////////| Create a new event section |////////////////////

export const RoleDefinitionIds = {
  FullControl: 1073741829,
  Design: 1073741828,
  Edit: 1073741830,
  Contribute: 1073741827,
  Read: 1073741826,
  LimitedAccess: 1073741825
}

export async function updateEventMetadata(params: ICreateNewEventParams) {
  const progress = params.progress || (() => {})
  console.log("updateEventMetadata() params:", params)
  progress("Updating calendar event...")
  const event = params.event || {}

  try {
    const updateResult = await sp.web
      .lists.getByTitle(calendarTitle)
      .items.getById(event.id)
      .update({
        EventDate: event.startDate.toISOString(),
        EndDate: event.endDate ? event.endDate.toISOString() : undefined,
        Location: event.location || "",
        CommsManagersId: {
          results: event.commsManagers.map((manager: any) => {
            return (
              (manager.User && manager.User.Id) ||
              manager.id
            )
          })
        }      
      })
    console.log("updateEventMetadata() updateResult:", updateResult)
  } catch (err) {
    console.error("updateEventMetadata() err:", err)
  }
}

export async function createNewEvent(params: ICreateNewEventParams): Promise<number> {
  const progress = params.progress || (() => {})
  console.log("createNewEvent() params:", params)

  progress("Creating calendar event...")
  const calendarItemId = await createCalendarEvent(params)
  console.log("createNewEvent() newEventId:", calendarItemId)

  progress("Creating folders...")
  const folders = await createNewEventFolders(params)
  console.log("createNewEvent() folders:", folders)

  progress("Creating groups...")
  const groups = await createNewEventGroups(calendarItemId, params)

  progress("Setting permissions...")
  await setNewEventPermissions(calendarItemId, params, folders, groups)

  await updateCalendarItem(calendarItemId, folders.eventFolderItemId)

  return calendarItemId
}

async function createNewEventFolders(params: ICreateNewEventParams): Promise<ICreateNewEventFoldersResult> {
  const result: ICreateNewEventFoldersResult = {
    eventFolderItemId: -1,
    sharedFolderItemId: -1,
    execFolders: []
  }
  
  const eventList = await getEventListDetails()

  // Create the event folder
  result.eventFolderItemId = await createFolder2(eventList.url, params.event.eventName)

  // Create all the execs folders
  for (const exec of params.event.execs) {
    const execFolderId = await createFolder2(`${eventList.url}/${params.event.eventName}`, exec.name)
    
    result.execFolders.push({
      execId: exec.id,
      itemId: execFolderId
    })
  }

  // Create the shared folder
  result.sharedFolderItemId = await createFolder2(`${eventList.url}/${params.event.eventName}`, "Shared")

  return result
}

async function createFolder2(currentFolderPath: string, newFolderName: string): Promise<number> {
  const newFolder = encodeURIComponent(`${currentFolderPath}/${newFolderName}`)
  const { data, folder } = await sp.web.getFolderByServerRelativeUrl(currentFolderPath)
    .folders.add(newFolder)

  const urlAndQuery = await folder.listItemAllFields.select("Id").toUrlAndQuery()
  const item = await folder.listItemAllFields.select("Id").get()
  console.log("createFolder2() urlAndQuery:", urlAndQuery, "item:", item)

  return item.Id as number
}


async function createNewEventGroups(calendarItemId: number, params: ICreateNewEventParams): Promise<ICreateNewEventGroupsResult> {
  try {
    const result: ICreateNewEventGroupsResult = {
      sharedGroupId: -1,
      execGroups: []
    }

    const siteVisitorsGroup = await sp.web.associatedVisitorGroup.get()
    //const copyUsersFromGroupIds = params.event.execs.map(x => x.groupId)
    result.sharedGroupId = await createEventGroup(`Event ${calendarItemId} Shared Visitors`, siteVisitorsGroup.Id) //, null, copyUsersFromGroupIds)

    for (let i = 0; i < params.event.execs.length; i += 1) {
      const exec = params.event.execs[i]
      const groupId = await createEventGroup(`Event ${calendarItemId} ${exec.name} Members`, siteVisitorsGroup.Id)
      
      result.execGroups.push({
        execId: exec.id,
        groupId
      })
    }
    
    return result

  } catch (err) {
    console.log("createNewGroups() err:", err)
    throw err
  }
}


export async function createEventGroup(name: string, ownersGroupId?: number /*, copyUsersFromGroupIds?: number[]*/): Promise<number> {
  try {
    const result = await sp.web.siteGroups.add({
      Title: name,
      Description: name,
      AllowMembersEditMembership: true,
      AllowRequestToJoinLeave: false,
      AutoAcceptRequestToJoinLeave: false,
      OnlyAllowMembersViewMembership: false,
    } as ICreateGroupOpts)

    if (!ownersGroupId) {
      const ownersGroup = await sp.web.associatedOwnerGroup.get()
      ownersGroupId = ownersGroup.Id
    }
    const contextInfo: ContextInfo = await sp.site.getContextInfo()
    const spGroupGuidConstant = "740c6a0b-85e2-48a0-a494-e0f1759d4aa7"
    const newGroupId = result.data.Id
    const { Id: siteId } = await sp.site.select("Id").get()
    const processQueryUrl = siteUrl + "/_vti_bin/client.svc/ProcessQuery"

    console.log("createSharedGroup() siteId:", siteId, "processQueryUrl:", processQueryUrl, "newGroupId:", newGroupId, "digest:", contextInfo.FormDigestValue.toString())

    const http = new HttpClient()
    const response = await http.post(processQueryUrl, {
      headers: {
        "Content-Type": "text/xml",
        "Accept": "application/json; odata=verbose",
        "X-RequestDigest": contextInfo.FormDigestValue.toString()
      },
      body: (`
        <Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="15.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
          <Actions>
            <SetProperty Id="1" ObjectPathId="2" Name="Owner">
              <Parameter ObjectPathId="3" />
            </SetProperty>
            <Method Name="Update" Id="4" ObjectPathId="2" />
          </Actions>
          <ObjectPaths>
            <Identity Id="2" Name="${spGroupGuidConstant}:site:${siteId}:g:${newGroupId}" />
            <Identity Id="3" Name="${spGroupGuidConstant}:site:${siteId}:g:${ownersGroupId}" />
          </ObjectPaths>
        </Request>
      `)
    })
    const result2 = await response.json()
    console.log("createSharedGroup() result2:", result2)

    // if (copyUsersFromGroupIds && copyUsersFromGroupIds.length) {
    //   // Copy users from exec group into event exec group
    //   let usersToCopy = []
    //   for (const copyUsersFromGroupId of copyUsersFromGroupIds) {
    //     const users = await sp.web.siteGroups.getById(copyUsersFromGroupId).users.get()
    //     usersToCopy = [...usersToCopy, ...users]
    //   }
    //   for (const user of usersToCopy) {
    //     await sp.web.siteGroups.getById(newGroupId).users.add(user.LoginName)
    //   }
    // }

    return newGroupId

  } catch (err) {
    console.log("createSharedGroup() err:", err)
    throw err    
  }
}

async function setNewEventPermissions(calendarItemId: number, params: ICreateNewEventParams, folders: ICreateNewEventFoldersResult, groups: ICreateNewEventGroupsResult): Promise<void> {
  // Get permission
  const roleDefs: any[] = await sp.web.roleDefinitions.get()
  const managePerms = find(roleDefs, x => x.Name === "Manage Permissions")
  if (!managePerms) {
    console.error(`Custom permission level "Manage Permissions" is missing.`)
    throw new Error(`Custom permission level "Manage Permissions" is missing.`)
  }

  const currentUser = await sp.web.currentUser.get()
  const ownersGroup = await sp.web.associatedOwnerGroup.get()

  const execGroupIds = params.event.execs.map(x => x.groupId)
  const execEventGroupIds = groups.execGroups.map(x => x.groupId)

  await setUniquePermissions2(
    documentLibraryTitle,
    folders.eventFolderItemId,
    [{
      groupIds: [...execGroupIds, ...execEventGroupIds, groups.sharedGroupId],
      roleDefIds: [RoleDefinitionIds.Read, managePerms.Id]
    },{
      groupIds: [ownersGroup.Id],
      roleDefIds: [RoleDefinitionIds.FullControl]
    }],
    currentUser.Id
  )

  // Set permissions on each exec folder
  for (const execFolder of folders.execFolders) {
    const execGroupId = params.event.execs.filter(x => x.id === execFolder.execId).map(x => x.groupId)
    const execEventGroupId = groups.execGroups.filter(x => x.execId === execFolder.execId).map(x => x.groupId)
    
    await setUniquePermissions2(
      documentLibraryTitle,
      execFolder.itemId,
      [{
        groupIds: [...execGroupId, ...execEventGroupId],
        roleDefIds: [RoleDefinitionIds.Contribute, managePerms.Id]
      },{
        groupIds: [ownersGroup.Id],
        roleDefIds: [RoleDefinitionIds.FullControl]
      }],
      currentUser.Id
    )
  }

  await setUniquePermissions2(
    documentLibraryTitle,
    folders.sharedFolderItemId,
    [{
      groupIds: [...execGroupIds, ...execEventGroupIds],
      roleDefIds: [RoleDefinitionIds.Contribute, managePerms.Id]
    },{
      groupIds: [groups.sharedGroupId],
      roleDefIds: [RoleDefinitionIds.Read]
    },{
      groupIds: [ownersGroup.Id],
      roleDefIds: [RoleDefinitionIds.FullControl]
    }],
    currentUser.Id
  )

  // Update calendar item permissions
  for (const execEventGroupId of execEventGroupIds) {
    await sp.web
      .lists.getByTitle(calendarTitle)
      .items.getById(calendarItemId)
      .roleAssignments.add(execEventGroupId, RoleDefinitionIds.Read)
  }
  await sp.web
    .lists.getByTitle(calendarTitle)
    .items.getById(calendarItemId)
    .roleAssignments.add(groups.sharedGroupId, RoleDefinitionIds.Read)
}

async function setUniquePermissions2(libraryTitle: string, itemId: number, assignments: UniqueRoleAssignment[], removeUserId?: number) {
  // Break inheritance
  await sp.web.lists
    .getByTitle(libraryTitle).items
    .getById(itemId)
    .breakRoleInheritance()

  for (const assignment of assignments) {
    for (const groupId of assignment.groupIds) {
      for (const roleDefId of assignment.roleDefIds) {
        await sp.web.lists
          .getByTitle(libraryTitle).items
          .getById(itemId)
          .roleAssignments.add(groupId, roleDefId)      
      }
    }
  }

  if (removeUserId && removeUserId > 0) {
    // Remove self
    await sp.web.lists
      .getByTitle(libraryTitle).items
      .getById(itemId)
      .roleAssignments.remove(removeUserId, RoleDefinitionIds.FullControl)
  }
}

async function createCalendarEvent(params: ICreateNewEventParams): Promise<number> {
  const ownersGroup = await sp.web.associatedOwnerGroup.get()
  const itemId = await createCalendarItem(params.event)
  const execGroupIds = params.event.execs.map(x => x.groupId)

  // Get current user
  const currentUser = await sp.web.currentUser.get()

  await setUniquePermissions2(
    calendarTitle,
    itemId,
    [{
      groupIds: [ownersGroup.Id],
      roleDefIds: [RoleDefinitionIds.FullControl]
    },{
      groupIds: execGroupIds,
      roleDefIds: [RoleDefinitionIds.Read]
    },{
      groupIds: [],
      roleDefIds: []
    }],
    currentUser.Id
  )

  return itemId
}

async function updateCalendarItem(itemId: number, eventFilesId: number) {
  await sp.web.lists
    .getByTitle(calendarTitle).items
    .getById(itemId).update({
      EventFilesId: eventFilesId
    })
}

async function createCalendarItem(event: IEventFormValues): Promise<number> {
  const result = await sp.web.lists
    .getByTitle(calendarTitle).items
    .add({
      Title: event.eventName,
      EventDate: event.startDate.toISOString(),
      EndDate: event.endDate ? event.endDate.toISOString() : undefined,
      Location: event.location || undefined,
      CommsManagersId: {
        results: event.commsManagers.map((manager: SharePointUserPersona) => {
          return manager.User.Id
        })
      }
    })
  
  return result.data.Id
}

// This will add a t=123 query string to bust
// the server side cache if there is one
function noServerCaching(request: ODataQueryable): ODataQueryable {
  request.query.add("t", new Date().getTime().toString())
  return request
}

export async function isCurrentUserAnAdmin() {
  let currentUser = await sp.web.currentUser.get()
  if (currentUser.IsSiteAdmin) {
    return true
  }
  try {
    currentUser = await sp.web.associatedOwnerGroup
      .users.getByLoginName(currentUser.LoginName).get()
    console.log("currentUser:", currentUser)
    return true
  } catch (err) {
    return false
  }
}

export async function isCurrentUserInAnExecGroup(): Promise<boolean> {
  const items = await sp.web.lists
    .getByTitle("Executives").items
    .get()
    
  return items && items.length > 0
}

export async function getGroupsAndMembers(fsObject: IFSObject): Promise<IPrincipal[]> {
  const req = sp.web.lists
    .getByTitle("Event Documents").items
    .getById(fsObject.id)
    .expand(
      "RoleAssignments",
      "RoleAssignments/Member",
      "RoleAssignments/Member/Users",
      "RoleAssignments/RoleDefinitionBindings"
    )
  // Cache bust
  req.query.add("t", new Date().getTime().toString())
  const item = await req.get()

  return item.RoleAssignments
    .filter(assignment => {
      return assignment.RoleDefinitionBindings
        .filter(bindings => bindings.Name === "Limited Access")
        .length === 0
    })
    .map(assignment => {
      if (assignment.Member["odata.type"] === "SP.User") {
        return {
          type: PrincipalType.user,
          id: assignment.Member.Id,
          email: assignment.Member.Email,
          name: assignment.Member.Title
        } as IPrincipal
      } else if (assignment.Member["odata.type"] === "SP.Group") {
        return {
          type: PrincipalType.group,
          id: assignment.Member.Id,
          name: assignment.Member.Title,
          principals: assignment.Member.Users.map(user => ({
            type: PrincipalType.user,
            id: user.Id,
            email: user.Email,
            name: user.Title
          } as IPrincipal))
        } as IPrincipal
      }
    })
}

export async function addUsersToFile(file: IFSObject, users: { id: number, login: string }[]) {
  const item = await sp.web.lists
      .getByTitle(eventDocumentLibraryTitle).items
      .getById(file.id)
    
  const uniqueItem = await noServerCaching(
    item.select("HasUniqueRoleAssignments")
  ).get()

  if (!uniqueItem || !uniqueItem.HasUniqueRoleAssignments) {
    await item.breakRoleInheritance(true, false)
  }

  for (let i = 0; i < users.length; i += 1) {
    await item.roleAssignments.add(users[i].id, RoleDefinitionIds.Contribute)
  }
}

export async function addUsersToGroup(groupId: number, userLoginNames: string[]) {
  for (let i = 0; i < userLoginNames.length; i += 1) {
    const loginName = userLoginNames[i]
    await sp.web.siteGroups
      .getById(groupId).users
      .add(loginName)
  }
}

export async function removeUserFromGroup(groupId: number, userId: number): Promise<any> {
  return (
    sp.web.siteGroups
    .getById(groupId).users
    .removeById(userId)
  )
}

export async function getGroupMembers(groupId: number): Promise<IUser[]> {
  const users = await sp.web.siteGroups.getById(groupId).users.get()

  return users
    .filter(user => (
      user.PrincipalType === 1 &&
      user.Title !== "System Account"
    ))
    .map(user => ({
      Id: user.Id,
      Title: user.Title,
      Email: user.Email
    }) as IUser)
}

//TODO: Expand users
export async function getExecutiveGroups(oDataFilter: string): Promise<any>{
  let siteGroups: ISiteGroup[] = []
  return sp.web.siteGroups.select("Id","Title","Description", "OwnerTitle","Users/Id", "Users/Title","Users/UserPrincipalName", "Users/Email").expand("Users").filter(oDataFilter).get()
  .then( groups => {
    groups.forEach(group => {
      let g: ISiteGroup
      let users: IUser[]
      
      g = {
        Id: group.Id,
        Title: group.Title,
        //Owner: group.Owner,
        Description: group.Description,
        Users: []
      } as ISiteGroup

      group.Users.forEach(user => {
        g.Users.push ({Id: user.Id, Title: user.Title, Email: user.Email} as IUser)
      })
      siteGroups.push(g) 
    })
    return siteGroups
  })
}
