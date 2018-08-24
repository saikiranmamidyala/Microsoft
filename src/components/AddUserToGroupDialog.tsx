import * as React from 'react'

import {
  Dialog,
  DialogType,
  DialogFooter
} from 'office-ui-fabric-react/lib/Dialog'

import {
  Label
} from 'office-ui-fabric-react/lib/Label'

import {
  TextField
} from 'office-ui-fabric-react/lib/TextField'

import {
  NormalPeoplePicker
} from 'office-ui-fabric-react/lib/Pickers'

import {
  PrimaryButton,
  DefaultButton
} from 'office-ui-fabric-react/lib/Button'

import {
  Spinner,
  SpinnerSize
} from 'office-ui-fabric-react/lib/Spinner'

import {
  Icon
} from 'office-ui-fabric-react/lib/Icon'

import PeoplePicker, {
  IPerson
} from '../components/PeoplePicker'

import {
  searchPeople,
  IExecutive,
  IUser,
  ISiteGroup,
  addUsersToGroup,
  SharePointUserPersona
} from '../shared/SharePoint'

import { sp } from 'sp-pnp-js/lib/pnp';

import styles from './AddUserToGroupDialog.module.scss'
import { autobind } from 'office-ui-fabric-react/lib/Utilities';

export interface IAddUserToGroupDialogState {
  activeGroup: ISiteGroup
  newUsers: IPerson[]
}

export interface IAddUserToGroupDialogProps {
  activeGroup: ISiteGroup
  //save: (newUser:IPerson[]) => void
  onDismiss?: () => void
}

export default class AddUserToGroupDialog extends React.Component<IAddUserToGroupDialogProps, IAddUserToGroupDialogState> {
  constructor(props: IAddUserToGroupDialogProps) {
    super(props)
    this.state = {
      activeGroup: null,
      newUsers: null
    }
  }

  public render(){
    const dialogTitle = "Add user(s) to this group"
    return (
      <Dialog
        hidden={false}
        modalProps={{
          className: styles.AddUserToGroupDialog
        }}
        dialogContentProps={{
          type: DialogType.normal,
          title: dialogTitle
        }}
        onDismiss={() => {
          this.props.onDismiss()
        }}>
          <div>
            <Label>Invite people to collaborate</Label>
            <div className={styles.flexRow}>
              <div className={styles.flexGrow}>
                <PeoplePicker
                  placeholder="Start typing a name..."
                  selectedItems={this.state.newUsers}
                  onChange={newUsers => this.setState({ newUsers })}
                  autoFocus={true}
                />
              </div>
          </div>
        </div>
        <DialogFooter>
        <PrimaryButton text="Add"
                className={styles.inviteButton}
                onClick={() => {
                  this.addUsers()
                  .then(() => this.props.onDismiss())
                }}
              />
          <DefaultButton text="Close"
            onClick={() => {
              this.props.onDismiss()
            }}
            />
        </DialogFooter>
      </Dialog>
    )
  }
  private async addUsers(): Promise<any>{
    const loginNames = (this.state.newUsers as SharePointUserPersona[]).map( x => ({
      id: x.User.Id,
      login: x.User.LoginName
    }))
    .map(u => {
      return u.login
    })
    return addUsersToGroup(this.props.activeGroup.Id, loginNames)
  }
}