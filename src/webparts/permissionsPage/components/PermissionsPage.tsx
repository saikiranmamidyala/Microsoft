import * as React from 'react';
import styles from './PermissionsPage.module.scss';
import { IPermissionsPageProps } from './IPermissionsPageProps';
import { escape, findIndex } from '@microsoft/sp-lodash-subset';
import Loading from '../../../components/Loading'

import {
  Persona,
  PersonaSize,
  PersonaPresence
} from 'office-ui-fabric-react/lib/Persona'

import { sp } from 'sp-pnp-js'

import {
  getExecutiveGroups,ISiteGroup,
  IExecutive, getAllExecutives, 
  getGroupMembers, IUser, removeUserFromGroup,cdnAssetsBaseUrl
} from '../../../shared/SharePoint'

import { SiteGroup, SiteGroups } from 'sp-pnp-js/lib/sharepoint/sitegroups';
import { DetailsList, DetailsListLayoutMode, ConstrainMode, CheckboxVisibility, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Breadcrumb } from 'office-ui-fabric-react/lib/Breadcrumb';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import AddUserToGroupDialog from '../../../components/AddUserToGroupDialog';
import * as _ from 'lodash';
import { CSSProperties } from 'react';

export interface IPermissionPageProps {
  description: string
  siteTitle: string
  //siteOwnersGroup: SiteGroup
  //siteMembersGroup: SiteGroup
}

export interface IPermissionPageState {
  loading: boolean
  execs: IExecutive[]
  execGroups: ISiteGroup[]
  //siteOwnersGroup: ISiteGroup
  //siteMembersGroup: ISiteGroup
  activeGroup: ISiteGroup
  showAddUserToGroupDialog: boolean
}

export default class PermissionsPage extends React.Component<IPermissionPageProps, IPermissionPageState> {
  constructor(props){
    super(props)
    this.state = {
      loading: true,
      execs: [],
      execGroups: [],
      //siteOwnersGroup: null,
      //siteMembersGroup: null,
      activeGroup: null,
      showAddUserToGroupDialog: false
    }
  }

  public async componentDidMount(){
    const execs = await getAllExecutives()
    //this.setState({})

    let execIds = execs.map(exec => exec.groupId)
    let filter =  execIds.map(id => "Id eq " + id )
    
    const execGroups =  await getExecutiveGroups(filter.join(" or ")) as ISiteGroup[]
    //let siteOwnersGroup  = await sp.web.associatedOwnerGroup.select('Id',"Title", "Description" ).get().then((g) => { return {Id: g.Id, Title: g.Title, Description: g.Description, Users: [] }})
    //siteOwnersGroup.Users = await getGroupMembers(siteOwnersGroup.Id)

    //let siteMembersGroup = await sp.web.associatedMemberGroup.select("Id","Title", "Description", "Users").get().then((g) => { return {Id: g.Id, Title: g.Title, Description: g.Description, Users: [] }})
    //siteMembersGroup.Users = await getGroupMembers(siteMembersGroup.Id)
    this.setState({
      execs: execs,
      execGroups: execGroups, 
      activeGroup: (execGroups.length > 0) ? execGroups[0] : null,
      //siteOwnersGroup
    }) 
  }

  private showMembers(exec: ISiteGroup) {
    // console.log("showMembers(id): " + exec.Id.toString())
  }

  //private getAllGroups(oDataFilter: string): Promise<SiteGroups> {
    //return  getExecutiveGroups(oDataFilter)
  //}

  private getExecInfoByGroupId(id: number): IExecutive {
    let exec: IExecutive = null
    this.state.execs.forEach((e) => {
      if(e.groupId === id) {
        exec = e
      }
    })
    return exec;
  }

  private getGroupButton(group: ISiteGroup, isExec?:boolean){
    const className = (this.state.activeGroup && this.state.activeGroup.Id === group.Id) ? styles.activeGroup : ""
    if(isExec) {
      const siteGroupExec = this.getExecInfoByGroupId(group.Id)
      return  (
        <div 
          className={styles.permButton + " " + className}
          onClick={ () => {
            this.setState({activeGroup: group})
            //this.showMembers(execGroup)}
          }}
        >
          <Persona
            size={PersonaSize.size28}
            presence={PersonaPresence.none}
            imageInitials={siteGroupExec.initials}
            imageUrl={siteGroupExec.imageUrl==""?`${cdnAssetsBaseUrl}/images/DefaultProfile.png` :siteGroupExec.imageUrl}
            
            hidePersonaDetails={true}
          />
          <div className={styles.userName}>
            {group.Title}
          </div>
      </div>
      )
    }
    else{
      return (
        <div 
          className={styles.permButton + " " + className}
          onClick={ () => {
            this.setState({activeGroup: group})
            //this.showMembers(execGroup)}
          }}
        >
          <div className={styles.userName}>
            {group.Title}
          </div>
        </div>
      )
    }
  }
  private renderDeleteButton(user: IUser) {
    return (  
      <div className="groupSelector">
        <DefaultButton
          className={styles.transparent}
          onClick={() => {
            removeUserFromGroup(this.state.activeGroup.Id, user.Id)
            .then(() => this.updateActiveGroupState())
          }}
        >X
        </DefaultButton>
      </div>
    )
  }

  private async updateActiveGroupState(){
    const activeExecGroupIndex = _.findIndex(this.state.execGroups, (sg) => {return sg.Id === this.state.activeGroup.Id})
    const activeGroup =  await getExecutiveGroups(`Id eq ${this.state.activeGroup.Id}`) as ISiteGroup[]
    let tempExecs = this.state.execGroups
    tempExecs[activeExecGroupIndex] = activeGroup[0]

    this.setState({
      execGroups: tempExecs,
      activeGroup: activeGroup[0]
    })
  }

  private renderExpanderButton(expanded: boolean) {
    return (
      <div>

      </div>
    )
  }

  private addUserToGroup() {
    this.setState({showAddUserToGroupDialog: true})
  }
  
  @autobind
  private showAddUserToGroupDialog() {
    //this.showDialog("showRenameDialog", item)
  }

  private resizeRightPanel(): CSSProperties{
    if(this.state.activeGroup) {
      let pageHeight = document.querySelector(".SPPageChrome").getBoundingClientRect().height
      let rightTop = document.querySelector("#leftContentContainer").getBoundingClientRect().top
      let topToBottom = pageHeight - rightTop - 150

      return {minHeight: topToBottom} as CSSProperties
    }
    else {
      return {} as CSSProperties
    }
  }

  public render(): React.ReactElement<IPermissionsPageProps> {
    const showAddUserToGroupDialog = this.state.showAddUserToGroupDialog;
    return (
      <div className={styles.permissionsPage}>
         <div className={styles.outer}>
         <div className={styles.flexShrink}>
            <div className={styles.pageBreadcrumb}>
              <Breadcrumb
                items={[
                  { key: "crumb0", text: "Permissions", isCurrentItem: true }
                ]} />
            </div>
            <div className="tabRow">
                <div className={styles.tab}>Permission Groups</div>
            </div>
          </div>
          <div className={styles.inner}>
            <div className={styles.left}>
              <div className={styles.leftContentContainer} id="leftContentContainer">
                <div className={styles.sectionHeader}>Executive Groups</div>
                <div className={styles.collapsableSection}>
                  {this.state.execGroups.map( e => this.getGroupButton(e, true))}
                </div>
              {/*} <div className={styles.sectionHeader}>Common Groups</div>
                <div className={styles.collapsableSection}>
                  {this.state.siteOwnersGroup && this.getGroupButton(this.state.siteOwnersGroup)}
                </div>*/}
              </div> 
            </div>
            {this.state.activeGroup && (
            <div className={styles.right}>
              <div className={styles.rightContentContainer} id="rightContentContainer" style={this.resizeRightPanel()}>
                <div className={styles.rightTopContainer}>
                  <PrimaryButton
                    className={styles.blueButton}
                    onClick={() => {
                      this.setState({ showAddUserToGroupDialog: true }) 
                      // console.log("add new user to group botton clicked")
                    }}
                    iconProps={{ iconName: "AddFriend" }}>
                    ADD NEW USER
                    
                  </PrimaryButton>
                </div>
                {this.state.activeGroup && (
                <div>
                  <DetailsList
                    columns = {[
                    {
                        key: "column1",
                        fieldName: "name",
                        name: "Name",
                        minWidth: 150,
                        maxWidth:250,
                        isResizable: true,
                        onRender: (user) => user.Title    
                      }, {
                        key: "column2",
                        fieldName: "email",
                        name: "Email address",
                        minWidth: 150,
                        maxWidth: 200,
                        isResizable: true,
                        onRender: (user) => user.Email
                      }, {
                        key: "column3",
                        fieldName: "remove ",
                        name: "Remove",
                        minWidth: 100,
                        maxWidth: 100,
                        isResizable: false,
                        onRender: this.renderDeleteButton.bind(this)
                      }
                    ]}
                    items={this.state.activeGroup ? this.state.activeGroup.Users : []}
                    layoutMode={DetailsListLayoutMode.justified}
                    constrainMode={ConstrainMode.unconstrained}
                    checkboxVisibility={CheckboxVisibility.hidden}
                  />
                </div>
                )}
              </div>
            </div>
            )} 
          </div> 
        </div>
        {
          showAddUserToGroupDialog && (
            <AddUserToGroupDialog
              activeGroup={this.state.activeGroup}
              onDismiss={async () => {
                this.setState({
                  showAddUserToGroupDialog: false,
                })
                this.updateActiveGroupState()
              }}
            >test</AddUserToGroupDialog>
          )
        }
      </div>
    )
  }
}