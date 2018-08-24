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
  ODataQueryable
} from 'sp-pnp-js'

import { IFSObject, IExecutive, getAllExecutives } from '../../../shared/SharePoint';

import  { siteCollectionUrl } from '../../../shared/SharePoint';

const execAssetsLibraryTitle: string = 'Exec Assets';

import * as moment from 'moment';



const isExecFolderLeveL = (serverRelativeUrl) => {
  return serverRelativeUrl.split("/").length === 5
}

const isInsideExecFolder = (serverRelativeUrl) => {
  return serverRelativeUrl.split("/").length > 5
}

export async function getFiles(folderUrl: string, listName: string): Promise<IFSObject[]> {
  let items

  try {
    // First we try requesting `RoleAssignments`. If the user does not
    // have the access to enumerate permissions, we make the request
    // again (below) without `RoleAssignments`
    items = await sp.web.lists
    .getByTitle(listName)
    .items
    .select(
    "*",
    "File/ModifiedBy/Title",
    "HasUniqueRoleAssignments",
    "ServerRedirectedEmbedUrl",
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
    "RoleAssignments",
    "RoleAssignments/Member",
    "RoleAssignments/RoleDefinitionBindings",
    )
    .get();
  } catch (err) {
    items = await sp.web.lists
    .getByTitle(listName)
    .items
    .select(
    "*",
    "File/ModifiedBy/Title",
    "HasUniqueRoleAssignments",
    "ServerRedirectedEmbedUrl",
    )
    .expand(
    "File",
    "File/Name",
    "File/UniqueId",
    "File/ModifiedBy",
    "File/ModifiedBy/Title",
    "File/ServerRelativeUrl",
    "File/Length",
    "Folder"
    )
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

async function getExecAssets() {
  let items;
  try {
    items = await sp.web.lists
    .getByTitle(execAssetsLibraryTitle)
    .items
    .select(
    "*",
    "File/ModifiedBy/Title",
    "HasUniqueRoleAssignments",
    "ServerRedirectedEmbedUrl",
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
    "RoleAssignments",
    "RoleAssignments/Member",
    "RoleAssignments/RoleDefinitionBindings",
    )
    .get();
    
    return items;

  } catch (err) {
    // console.warn(err);
  }
}

export function mapExecsToFolders(execs) {
  return Promise.resolve()
  .then(() => {
    return getExecAssets();
  })
  .then(items => {

    // console.log(items);

    let files = items.map(item => {
      return {
        name: item.File ? item.File.Name : item.Folder.Name,
        type: item.File ? 'file' : 'folder',
        id: item.Id,
        serverRelativeUrl: item.File ? item.File.ServerRelativeUrl : item.Folder.ServerRelativeUrl,
        size: (item.File && item.File.Length) ? parseInt(item.File.Length, 10) : 0,
        modifiedBy: (item.File && item.File.ModifiedBy && item.File.ModifiedBy.Title) || "",
        modified: item.Modified ? moment(item.Modified) : "",
        hasUniquePermissions: item.HasUniqueRoleAssignments,
        serverRedirectedEmbedUrl: item.ServerRedirectedEmbedUrl,
        uniqueFolderId: item.Folder ? item.Folder.UniqueId : "",
        uniqueFileId: item.File ? item.File.UniqueId : "",
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


      } as IFSObject;
    })

      let execsDict = execs.reduce((acc, exec) => {
        if (!acc[exec.name]) {
          acc[exec.name] = {};
        }
        return acc;
      }, {})

    //get folders synced to exec in the dict created above
    files.filter(file => file.type === 'folder' && execsDict[file.name])
    // .filter((folder) => execsDict[folder.name])
    .forEach(folder => {
      execsDict[folder.name] =  {
        root: folder.name,
        folderUrl: folder.serverRelativeUrl,
        id: folder.uniqueFolderId,
        modifiedDate: moment(folder.modified).format('YYYY-MM-DD')
      }
    })

    Object.keys(execsDict).forEach(key => {
      let matchedFiles = files.filter(file => {
        return file.name !== key;
      })
      .filter(file => {
        return file.serverRelativeUrl.split('/')[4] === key;
      })

      execsDict[key] = matchedFiles;
    })

      return execsDict;

  })
}

export function getExecChildren(execFolderUrl) {
  return getFiles(execFolderUrl, execAssetsLibraryTitle)
  .then(files => {
    // console.log(files, 'DID WE GET THE FILES???')
    return files;
  })

}
