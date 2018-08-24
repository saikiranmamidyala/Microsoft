import * as React from 'react'
import Teams from '../../../shared/Teams'
import { siteCollectionUrl } from '../../../shared/SharePoint'
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown'

export interface IConfigPageProps {
  description: string;
}

export interface IConfigPageState {
  selectedTab: IDropdownOption
}

export default class ConfigPage extends React.Component<IConfigPageProps, IConfigPageState> {
  constructor(props: IConfigPageProps) {
    super(props)

    this.state = {
      selectedTab: null
    }
  }

  public render() {
    return (
      <div>
        <Dropdown
          label="Select a tab to add"
          options={[
            { key: "1", text: "Events", entityId: "FutureEvents" },
            { key: "2", text: "Exec Assets", entityId: "Assets" },
            { key: "3", text: "Archive", entityId: "PastEvents" },
            { key: "4", text: "Calendar" },
            { key: "5", text: "Search" },
            { key: "6", text: "Permissions" },
          ]}
          onChanged={this.onChange.bind(this)} />
      </div>
    )
  }

  private onChange(item: IDropdownOption) {
    // console.log("onChange item:", item)
    if (Teams.connected) {
      this.setState({ selectedTab: item })
      Teams.api.settings.setValidityState(true)
      Teams.api.settings.registerOnSaveHandler(this.saveConfig.bind(this))
    }
  }

  private saveConfig(saveEvent) {
    const { selectedTab: tab } = this.state
    const entityId = (tab as any).entityId || tab.text

    Teams.api.settings.setSettings({
      suggestedDisplayName: tab.text,
      entityId,
      contentUrl: `${siteCollectionUrl}/SitePages/${entityId}.aspx`,
      removeUrl: `${siteCollectionUrl}/SitePages/RemoveTab.aspx`
    });
    saveEvent.notifySuccess();    
  }
}
