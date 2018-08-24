import * as React from 'react'
import * as moment from 'moment'
import { intersection } from 'lodash'

import { Breadcrumb } from 'office-ui-fabric-react/lib/Breadcrumb'
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox'
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button'

import Loading from '../../../components/Loading'
import ExecFilter from '../../../components/ExecFilter'
import EventDetailsDialog from '../../../components/EventDetailsDialog'

import FutureEventsList, {
  IFutureEventsListProps
} from '../../../components/FutureEventsList'

import {
  redirectToSitePage,
  IExecutive, getAllExecutives,
  IEvent, getFutureEvents,
  getEventListDetails,
  IEventListDetails,
  isCurrentUserAnAdmin
} from '../../../shared/SharePoint'

import styles from './FutureEventsPage.module.scss'

export interface IFutureEventsPageProps {
  description: string
}

export interface IFutureEventsPageState {
  loading: boolean
  userIsAdmin: boolean
  events: IEvent[]
  execs: IExecutive[]
  searchText: string
  selectedExecIds: number[]
  showNewEventDialog: boolean
  docLib: IEventListDetails
}

export default class FutureEventsPage extends React.Component<IFutureEventsPageProps, IFutureEventsPageState> {
  constructor(props) {
    super(props)

    this.state = {
      loading: true,
      userIsAdmin: false,
      events: [],
      execs: [],
      searchText: "",
      selectedExecIds: [],
      showNewEventDialog: false,
      docLib: null
    }
  }

  private async refreshEvents() {
    const events = await getFutureEvents()
    this.attachExecsToEvents(events, this.state.execs)
    this.setState({ events })
  }

  // This function is called after the
  // first render() is called
  public componentDidMount() {
    Promise.all([
      getAllExecutives(),
      getFutureEvents(),
      getEventListDetails(),
      isCurrentUserAnAdmin()
    ])
    .then(([ execs, events, docLib, userIsAdmin ]) => {
      if (!execs || !events || !docLib) {
        // console.error("execs, events, or docLib is null or undefined")
        throw new Error("execs, events, or docLib is null or undefined")
      }

      this.attachExecsToEvents(events, execs)

      const selectedExecIds = execs.map(exec => exec.id);

      // console.log("execs:", execs)
      // console.log("events:", events)
      // console.log("docLib:", docLib)

      this.setState({
        loading: false,
        userIsAdmin,
        events,
        execs,
        selectedExecIds,
        docLib
      })
    })
  }

  private attachExecsToEvents(events: IEvent[], execs: IExecutive[]) {
    events.forEach(evt => {
      evt.execs = evt.execs || []
      evt.principalIds.forEach(principalId => {
        for (let i = 0; i < execs.length; i += 1) {
          const exec = execs[i]
          if (principalId === exec.groupId) {
            evt.execs.push({ ...exec })
            return
          }
        }
      })
    })
  }

  public render() {
    if (this.state.loading) {
      return <Loading />
    }

    const { 
      userIsAdmin,
      events,
      execs,
      selectedExecIds,
      searchText,
      showNewEventDialog
    } = this.state

    return (
      <div className={styles.FutureEventsPage}>
        <div className={styles.flex} style={{ paddingBottom: "12px" }}>
          <div className={styles.flexGrow}>
            <div className={styles.pageBreadcrumb}>
              <Breadcrumb
                items={[
                  { key: "crumb0", text: "Events", isCurrentItem: true }
                ]} />
            </div>
            <ExecFilter
              execs={execs} 
              onChange={execIds => {
                // console.log("new selectedExecIds:", execIds)
                this.setState({ selectedExecIds: execIds })
              }} />
          </div>
          <div className={styles.flexColumn}>
            <div className={styles.flexGrow}>
              {userIsAdmin && (
                <PrimaryButton
                  onClick={() => {
                    this.setState({ showNewEventDialog: true })
                  }}
                  iconProps={{ iconName: "Add" }}>
                  EVENT
                </PrimaryButton>
              )}
              </div>
            <SearchBox
              style={{ width: "200px" }}
              onChange={val => this.setState({ searchText: val })}/>
          </div>
        </div>
        <FutureEventsList
          events={events}
          filters={[
            this.filterByExec.bind(this),
            this.filterBySearchText.bind(this)
          ]}
          onEventClick={evt => {
            redirectToSitePage("Event", { eventId: evt.id })
          }} />
        {showNewEventDialog && (
          <EventDetailsDialog
            onDismiss={() => this.setState({ showNewEventDialog: false })}
            onSuccess={eventId => {
              if (eventId) {
                redirectToSitePage("Event", { eventId })
              } else {
                this.setState({ showNewEventDialog: false })
                this.refreshEvents()
              }
            }} />
        )}
      </div>
    )
  }

  private filterByExec(evt: IEvent) {
    if (!evt.execs.length) {
      return true
    }
    return intersection(
      this.state.selectedExecIds,
      evt.execs.map(exec => exec.id)
    ).length > 0
  }

  private filterBySearchText(evt: IEvent) {
    const searchText = this.state.searchText.toLowerCase()
    return (
      evt.eventName.toLowerCase().indexOf(searchText) >= 0 ||
      (evt.location && evt.location.toLowerCase().indexOf(searchText) >= 0) ||
      (evt.startDate && evt.startDate.format("YYYY-MM-DD").toLowerCase().indexOf(searchText) >= 0) ||
      (evt.endDate && evt.endDate.format("YYYY-MM-DD").toLowerCase().indexOf(searchText) >= 0)
    )
  }
}
