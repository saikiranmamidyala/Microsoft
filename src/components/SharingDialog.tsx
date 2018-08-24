import * as React from 'react'

import {
  IFSObject,
  getGroupsAndMembers,
  IPrincipal,
  PrincipalType,
  addUsersToFile,
  IUser,
  IExecutive,
  SharePointUserPersona,
  IEvent,
  addUsersToGroup,
  RoleDefinitionIds,
  siteDomain,
  siteCollectionUrl,
  isCurrentUserAnAdmin,
  isCurrentUserInAnExecGroup
} from '../shared/SharePoint';

import {
  Dialog,
  DialogType,
  DialogFooter
} from 'office-ui-fabric-react/lib/Dialog'

import {
  SelectionMode,
  Selection,
  SelectionZone
} from 'office-ui-fabric-react/lib/utilities/selection/index'

import { GroupedList, IGroup } from 'office-ui-fabric-react/lib/GroupedList'
import { PrimaryButton, IconButton, DefaultButton } from 'office-ui-fabric-react/lib/Button'
import { DetailsRow, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { Label } from 'office-ui-fabric-react/lib/Label'
import { uniqBy, find, endsWith, startsWith } from 'lodash'
import PeoplePicker, { IPerson } from './PeoplePicker';
import styles from './SharingDialog.module.scss'
import { sp, PermissionKind, SharingRole } from 'sp-pnp-js/lib/pnp';
import { RoleDefinition } from 'sp-pnp-js/lib/sharepoint/roles';
import Loading from './Loading';

export interface ISharingDialogProps {
  event: IEvent
  file?: IFSObject
  folder?: IFSObject
  onDismiss: () => void
}

export interface ISharingDialogState {
  items: IRemovablePrincipal[]
  groups: IGroup[]
  invitees: IPerson[]
  userIsAdmin: boolean
  userIsInExecGroup: boolean
  canShare: boolean
  loading: boolean
}

export interface IRemovablePrincipal extends IPrincipal {
  isRemovable: boolean
}

export default class SharingDialog extends React.Component<ISharingDialogProps, ISharingDialogState> {
  private _selection = new Selection()

  constructor(props: ISharingDialogProps) {
    super(props)

    this.state = {
      items: [],
      groups: [],
      invitees: [],
      userIsAdmin: false,
      userIsInExecGroup: false,
      canShare: false,
      loading: true
    }
  }

  public async componentDidMount() {
    this.loadPermissions()
  }

  private async loadPermissions() {
    const target: IFSObject = this.props.file || this.props.folder
    let items: IRemovablePrincipal[] = []
    const groups: IGroup[] = []
    const outsiders: IRemovablePrincipal[] = []

    const ownersGroup = await sp.web.associatedOwnerGroup.get()

    const userIsAdmin = await isCurrentUserAnAdmin()
    // console.log("userIsAdmin:", userIsAdmin)

    const userIsInExecGroup = await isCurrentUserInAnExecGroup()
    // console.log("userIsInExecGroup:", userIsInExecGroup)

    let allowedPrincipals
    try {
      allowedPrincipals = await getGroupsAndMembers(target)
    } catch (err) {
      allowedPrincipals = null
    }
    // console.log("allowedPrincipals:", allowedPrincipals)

    if (!allowedPrincipals || !allowedPrincipals.length) {
      this.setState({ loading: false })
    }

    allowedPrincipals.forEach((principal: IRemovablePrincipal, i) => {
      if (principal.type === PrincipalType.user) {
        principal.isRemovable = true
        outsiders.push(principal)
      } else if (
        principal.type === PrincipalType.group &&
        principal.id !== ownersGroup.Id
      ) {
        groups.push({
          key: principal.id.toString(),
          name: principal.name,
          startIndex: items.length,
          count: principal.principals.length,
          isCollapsed: true,
          data: { groupId: principal.id }
        })
        if (principal.principals.length) {
          principal.principals.forEach((pri: IRemovablePrincipal) => {
            // This check is pretty fragile
            if (startsWith(principal.name, "Event")) {
              pri.isRemovable = true
            } else {
              pri.isRemovable = false
            }
            items.push(pri)
          })
        }
      }
    })

    if (outsiders.length) {
      groups.push({
        key: "outsiders",
        name: "Direct access",
        startIndex: items.length,
        count: outsiders.length,
        isCollapsed: true
      })
      items = items.concat(outsiders)
    }

    // console.log("init() groups:", groups)
    
    this.setState({
      items,
      groups,
      invitees: [],
      userIsAdmin,
      userIsInExecGroup,
      loading: false,
      canShare: groups.length > 0
    })
  }

  private createAccessLabel(userCount: number, isFile: boolean) {
    let label = "No one has"
    
    if (userCount > 0) {
       label = userCount.toString()
       if (userCount === 1) {
         label += " person has"
       } else {
         label += " people have"
       }
    }
    
    label += " access to"
    
    if (isFile) {
      label += " this document."
    } else {
      label += " all the documents in this space."
    }
    
    return label    
  }

  public render() {
    const { items, groups, userIsAdmin, userIsInExecGroup, canShare, loading } = this.state
    const { event, file, folder } = this.props
    const target = file || folder
    const uniqueUsers = uniqBy(items, "id").length
    const uniqueLabel = this.createAccessLabel(uniqueUsers, !!file)
    const dialogTitle = `${event.eventName}: ${file ? file.name : folder.name}`

    return (
      <Dialog
        hidden={false}
        modalProps={{
          className: styles.SharingDialog
        }}
        dialogContentProps={{
          type: DialogType.normal,
          title: dialogTitle
        }}
        onDismiss={() => {
          this.props.onDismiss()
        }}>

        {loading && (
          <Loading />
        )}
        {!loading && !canShare && (
          <h1>You don't have permission to share.</h1>
        )}
        {!loading && canShare && (
          <div>
            <Label>Invite people to collaborate</Label>
            <div className={styles.flexRow}>
              <div className={styles.flexGrow}>
                <PeoplePicker
                  placeholder="Start typing a name..."
                  selectedItems={this.state.invitees}
                  onChange={invitees => this.setState({ invitees })}
                  autoFocus={true}
                />
              </div>
              <PrimaryButton text="Invite"
                className={styles.inviteButton}
                onClick={() => this.inviteUsers()}
              />
            </div>

            {(userIsAdmin || userIsInExecGroup) && (
              <div className={styles.permissions}>
                <Label>{uniqueLabel}</Label>
                <SelectionZone
                  selection={this._selection}
                  selectionMode={SelectionMode.none}
                /* SelectionZone */>
                  <GroupedList
                    items={items}
                    groups={groups}
                    selection={this._selection}
                    selectionMode={SelectionMode.none}
                    groupProps={{
                      showEmptyGroups: true,
                    }}
                    onRenderCell={(nestingDepth?: number, item?: IRemovablePrincipal, index?: number) => {
                      return (
                        <div className={styles.flexRow}>
                          <div className={styles.flexGrow}>
                            <DetailsRow
                              columns={[
                                { key: "name", name: "Name", fieldName: "name", minWidth: 300 },
                                { key: "email", name: "Email", fieldName: "email", minWidth: 300 },
                              ]}
                              item={item}
                              itemIndex={index}
                              groupNestingDepth={nestingDepth}
                              selection={this._selection}
                              selectionMode={SelectionMode.none}
                            />
                          </div>
                          {item && item.isRemovable && (
                            <IconButton
                              iconProps={{ iconName: "Cancel" }}
                              onClick={() => this.removePermission(item, index)}
                            />
                          )}
                        </div>
                      )
                    }}
                    />
                </SelectionZone>
              </div>
            )}
          </div>
        )}

        <DialogFooter>
          <DefaultButton text="Close"
            onClick={() => this.props.onDismiss()}
            />
        </DialogFooter>
      </Dialog>
    )
  }

  private async inviteUsers() {
    const { file, folder } = this.props
    const { invitees, groups } = this.state

    // console.log("invitees:", invitees)

    if (invitees.length) {
      if (file) {
        const users = (invitees as SharePointUserPersona[]).map(x => ({
          id: x.User.Id,
          login: x.User.LoginName
        }))
        await addUsersToFile(file, users)

        const sharedWiths = (invitees as SharePointUserPersona[]).map(x => ({
          name: x.User.DisplayText,
          email: x.User.Email
        }))

        await this.sendNotificationEmail(
          file.name,
          `${siteDomain}/${file.serverRelativeUrl}`,
          sharedWiths
        )
      } else {
        // Figure out which group object we're working with
        let groupToAddUsersTo
        
        if (folder.name === "Shared") {
          groupToAddUsersTo = find(groups, x => endsWith(x.name, "Shared Visitors"))
        } else {
          groupToAddUsersTo = find(groups, x => startsWith(x.name, "Event"))
        }
        // console.log("groupToAddUsersTo:", groupToAddUsersTo, "groups:", groups)
        if (!groupToAddUsersTo) {
          return
        }

        // Gather all the user login names we need to invite
        // and add them to event exec group
        const loginNames = (invitees as SharePointUserPersona[]).map(x => x.User.LoginName)
        await addUsersToGroup(groupToAddUsersTo.data.groupId, loginNames)

        // Also add them to the SharePoint visitors group
        const visitorsGroup = await sp.web.associatedVisitorGroup.get()
        await addUsersToGroup(visitorsGroup.Id, loginNames)

        // Lastly, add them to the Shared folder


        // Gather info for the notification email
        const sharedWiths = (invitees as SharePointUserPersona[]).map(x => ({
          name: x.User.DisplayText,
          email: x.User.Email
        }))
        const evt = this.props.event

        // Send it
        await this.sendNotificationEmail(
          evt.eventName,
          `${siteCollectionUrl}/SitePages/Event.aspx?eventId=${evt.id}`,
          sharedWiths
        )
      }
      // Refresh permissions
      this.loadPermissions()
    }
  }

  private async removePermission(principal: IRemovablePrincipal, itemsIndex: number) {
    const group = this.findGroup(itemsIndex)
    // console.log("removePermission() principal:", principal, "itemsIndex:", itemsIndex, "group:", group)
    if (group) {
      if (group.name === "Direct access") {
        await this.removeUserFromDirectAccess(principal)
      } else {
        await this.removeUserFromGroup(principal, group)
      }
      this.loadPermissions()
    }
  }

  private findGroup(index: number) {
    for (let i = 0; i < this.state.groups.length; i += 1) {
      const group = this.state.groups[i]
      if (index >= group.startIndex && index <= group.startIndex + (group.count - 1)) {
        return group
      }
    }
  }

  private async removeUserFromGroup(user: IRemovablePrincipal, group: IGroup) {
    const siteGroup = await sp.web.siteGroups.getById(group.data.groupId)
    await siteGroup.users.removeById(user.id)
  }

  private async removeUserFromDirectAccess(user: IRemovablePrincipal) {
    const fsObj = this.props.file || this.props.folder
    await sp.web.lists
      .getByTitle("Event Documents").items
      .getById(fsObj.id).roleAssignments
      .remove(user.id, RoleDefinitionIds.Contribute)
  }

  private async sendNotificationEmail(
    sharedResourceName: string,
    sharedResourceUrl: string,
    sharedWiths: Array<{
      name: string,
      email: string
    }>
  ) {
    const currentUser = await sp.web.currentUser.get()
    const fromName = currentUser.Title

    for (let i = 0; i < sharedWiths.length; i += 1) {
      const sharedWith = sharedWiths[i]

      await sp.utility.sendEmail({
        To: [sharedWith.email],
        Subject: `${fromName} just shared "${sharedResourceName}" with you`,
        Body: (
          `Hello ${sharedWith.name},<br/><br/>` +
          `${fromName} just shared <a href="${sharedResourceUrl}">${sharedResourceName}</a> with you.<br/><br/>` +
          `Thanks,<br/>` +
          `Exec Comms Team`
        )
      })
    }
  }
}