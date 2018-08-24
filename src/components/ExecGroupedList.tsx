import * as React from 'react';
import {
  GroupedList,
  IGroup,
  IGroupDividerProps
} from 'office-ui-fabric-react/lib/GroupedList';
import './ExecGroupedList.module.scss';
import { IExecDocumentsProps } from './ExecDocuments';
import { Persona } from 'office-ui-fabric-react/lib/Persona';
import { cdnAssetsBaseUrl } from '../shared/SharePoint';

export interface IExecGroupedListProps {
  onRenderCell: () => void
  items: any[]
}

export interface IExecGroupedListState {
  items: any[]
  selectedCellId: string
}

export class GroupedListCustomExample extends React.Component<IExecGroupedListProps, IExecGroupedListState> {

  constructor(props) {
    super(props)
    this.state = {
      items: this.props.items,
      selectedCellId: this.state.selectedCellId
    }
  
  }

  public render() {
    return (
      <GroupedList
        ref='groupedList'
        items={ this.state.items }
        onRenderCell={ this._onRenderCell }
      />
    );
  }

  private _onRenderCell(nestingDepth: number, item: any, itemIndex: number) {
    return (
      <div data-selection-index={ itemIndex }>
        <Persona 
          hidePersonaDetails={false}
          imageInitials={item.initials}
          primaryText={""}
          imageUrl={item.imageUrl==""?`${cdnAssetsBaseUrl}/images/DefaultProfile.png` :item.imageUrl}
              
          className={'persona-Coin'}
        />
        <span className={'exec-Name'}>
          { item.name }
        </span>
      </div>
    );
  }

}