import * as React from 'react'
import * as moment from 'moment'

import ListFaces, {
  IListFacesProps
} from './ListFaces'

import {
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
  ConstrainMode
} from 'office-ui-fabric-react/lib/DetailsList'

import { IEvent } from '../shared/SharePoint'
import { sortBy, sortMomentsBy } from '../shared/util'

import styles from './FutureEventsList.module.scss'
import { Link } from 'office-ui-fabric-react/lib/Link';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { debounce } from 'lodash';

export interface IFutureEventsListProps {
  events: IEvent[]
  onEventClick?: (evt: IEvent) => void
  filters?: ((evt: IEvent) => boolean)[]
}

export interface IFutureEventsListState {
  sort: ISortableList
  columns: any[]
}

export interface ISortableList {
  index: number
  asc: boolean
}

export default class FutureEventsList extends React.Component<IFutureEventsListProps, IFutureEventsListState> {
  constructor(props) {
    super(props)

    this.state = {
      sort: {
        index: 2,
        asc: true
      },
      columns: [
        {
          key: "column0",
          fieldName: "execFaces",
          name: "Exec",
          minWidth: 60,
          maxWidth: 100,
          isResizable: false,
          onRender: item => (
            <ListFaces
              overflow={3}
              people={item.execs} />
          )
        },
        {
          key: "column1",
          fieldName: "eventName",
          name: "Name",
          minWidth: 200,
          isResizable: true,
          onColumnClick: () => this.sortColumn(1),
          onRender: item => (
            <Link onClick={() => this.eventClicked(item)}>
              {item.eventName}
            </Link>
          )
        },
        {
          key: "column2",
          fieldName: "startDate",
          name: "Start Date",
          minWidth: 150,
          maxWidth: 150,
          isResizable: true,
          onColumnClick: () => this.sortColumn(2),
          onRender: item => item.startDate.format("YYYY-MM-DD")
        },
        {
          key: "column3",
          fieldName: "endDate",
          name: "End Date",
          minWidth: 150,
          maxWidth: 150,
          isResizable: true,
          onColumnClick: () => this.sortColumn(3),
          onRender: item => item.endDate.format("YYYY-MM-DD")
        },
        {
          key: "column4",
          fieldName: "location",
          name: "Location",
          minWidth: 200,
          maxWidth: 200,
          isResizable: true,
          onColumnClick: () => this.sortColumn(4)
        }
      ]
    }

    // Initial sort UI
    const { columns, sort } = this.state
    const column = columns[sort.index]
    column.isSorted = true
    column.isSortedDescending = !sort.asc
  }

  public componentDidMount() {
    window.addEventListener("resize", debounce(() => {
      this.forceUpdate()
    }, 250))    
  }

  @autobind
  private sortColumn(index: number) {
    const { sort: prevSort, columns } = this.state
    // Default sorting order when column is first clicked
    const nextSort: ISortableList = { index, asc: true }

    // Toggle sorting order if column is clicked again
    if (prevSort.index === index) {
      nextSort.asc = !prevSort.asc
      const col = columns[nextSort.index]
      // Toggle column arrow UI
      col.isSortedDescending = !nextSort.asc
    } else {
      // We get here when the user clicks a column header
      // that is different than the currently sorted column
      
      // Remove arrow UI from currently sorted column
      const prevCol = columns[prevSort.index]
      delete prevCol.isSorted
      delete prevCol.isSortedDescending
      // Add arrow UI for the newly sorted column
      const nextCol = columns[nextSort.index]
      nextCol.isSorted = true
      nextCol.isSortedDescending = !nextSort.asc
    }

    // Re-render
    this.setState({
      sort: nextSort,
      columns
    })
  }

  // private _clickListenerEnabled = false

  // private detailsListRef(component) {
  //   if (component && !this._clickListenerEnabled) {
  //     this._clickListenerEnabled = true
  //     component.refs.root.addEventListener("click", evt => {
  //       if (this._cachedEvents) {
  //         for (let i = 0; i < evt.path.length; i += 1) {
  //           const el = evt.path[i]
  //           if (
  //             el.classList &&
  //             el.classList.contains("ms-List-cell")
  //           ) {
  //             const index = el.dataset.listIndex
  //             //console.log("clicked index:", index, "el:", el, "event:", this._cachedEvents[index])
  //             if (this._cachedEvents[index]) {
  //               this.props.onEventClick(this._cachedEvents[index])
  //             }
  //             return
  //           }
  //         }
  //       }
  //     })
  //   }
  // }

  //private _cachedEvents = null

  public render() {
    const { filters } = this.props
    const { sort, columns } = this.state
    let events = this.props.events.slice()

    // console.log("before filtering, events:", events)
    // Filtering
    filters.forEach(filter => {
      events = events.filter(filter)
    })
    // console.log("after filtering, events:", events)

    const { fieldName } = columns[sort.index]
    
    // Sorting
    if (fieldName === "eventName" || fieldName === "location") {
      sortBy(events, fieldName, sort.asc ? "asc" : "desc")
    } else if (fieldName === "startDate" || fieldName === "endDate") {
      sortMomentsBy(events, fieldName, sort.asc ? "asc" : "desc")
    }

    //this._cachedEvents = events

    // Rendering
    return /*this.props.events.length ?*/ (
      <div className={styles.FutureEventsList}
        ref={this.eventsListContainerRef}
      >
        <DetailsList
          //ref={this.detailsListRef.bind(this)}
          columns={columns}
          items={events}
          layoutMode={DetailsListLayoutMode.justified}
          constrainMode={ConstrainMode.unconstrained}
          checkboxVisibility={CheckboxVisibility.hidden}
          selectionMode={SelectionMode.none} />
      </div>
    ) /*: null*/
  }

  @autobind
  private eventsListContainerRef(div: HTMLDivElement) {
    if (div) {
      const windowHeight = window.innerHeight
      const divTop = div.getBoundingClientRect().top
      const height = windowHeight - divTop
      div.style.minHeight = `${height}px`
    }
  }

  @autobind
  private eventClicked(evt: IEvent) {
    if (this.props.onEventClick) {
      this.props.onEventClick(evt)
    }
  }
}

