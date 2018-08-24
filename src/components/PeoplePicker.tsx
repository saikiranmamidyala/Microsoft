import * as React from 'react'

import {
  IPersonaProps
} from 'office-ui-fabric-react/lib/components/Persona'

import {
  NormalPeoplePicker
} from 'office-ui-fabric-react/lib/Pickers'

import {
  searchPeople
} from '../shared/SharePoint'

export interface IPerson extends IPersonaProps {}

export interface IPeoplePickerProps {
  onChange: (items: IPerson[]) => void
  placeholder?: string
  selectedItems?: IPerson[]
  autoFocus?: boolean
}

export default class PeoplePicker extends React.Component<IPeoplePickerProps, {}> {
  public render() {
    const {
      onChange,
      placeholder,
      autoFocus
    } = this.props

    const selectedItems: IPersonaProps[] = this.props.selectedItems || []

    return (
      <NormalPeoplePicker
        inputProps={{ placeholder, autoFocus }}
        selectedItems={selectedItems}
        onResolveSuggestions={this.onResolveSuggestions.bind(this)}
        pickerSuggestionsProps={{
          suggestionsHeaderText: "Suggested People",
          loadingText: "Loading",
          noResultsFoundText: "No results found"
        }}
        onChange={items => onChange(items as IPerson[])} />
    )
  }

  private onResolveSuggestions(text: string, selectedItems?: IPersonaProps[]): Promise<IPersonaProps[]> {
    if (text && text.length > 2) {
      return searchPeople(text)
    }
    return Promise.resolve([])
  }  
}