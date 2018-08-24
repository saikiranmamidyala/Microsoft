import * as React from 'react';
import * as moment from 'moment'
import { intersection, uniqBy, orderBy } from 'lodash'

import { Breadcrumb } from 'office-ui-fabric-react/lib/Breadcrumb'
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox'
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button'

import Loading from '../../../components/Loading'
import ExecFilter from '../../../components/ExecFilter'

import {
  redirectToSitePage,
  IExecutive, getAllExecutives,
  IEvent, getPastEvents, getArchivedEvents, 
  IArchiveExecConfig, getArchivedExecutiveConfig
} from '../../../shared/SharePoint'

import styles from '../../futureEventsPage/components/FutureEventsPage.module.scss';
import styles2 from './PastEventsPage.module.scss';

import { IPastEventsPageProps } from './IPastEventsPageProps';
import { escape } from '@microsoft/sp-lodash-subset';

import PastEventsList, { IPastEventsListProps } from '../../../components/PastEventsList';


export interface IPastEventsPageProps {
  description: string
}

export interface IPastEventsPageState {
  loading: boolean
  events: IEvent[]
  execs: IExecutive[]
  searchText: string
  selectedExecIds: number[]
}

export default class PastEventsPage extends React.Component<IPastEventsPageProps, IPastEventsPageState> {
  constructor(props: IPastEventsPageProps) {
    super(props)
    // console.log(props.context)
    this.state = {
      loading: true,
      events: [],
      execs: [],
      searchText: "",
      selectedExecIds: []
    }
  }

  // This function is called after the
  // first render() is called
  public componentDidMount() {
    Promise.all([
      getAllExecutives(),
      getPastEvents(),
      getArchivedEvents(),
      getArchivedExecutiveConfig()
    ])
    .then(([ execs, events, archivedEvents, archivedExecutiveConfig ]) => {
      if (!execs || !events || !archivedEvents || !archivedExecutiveConfig) {
        // console.error("execs or events is null or undefined")
        throw new Error("execs or events is null or undefined")
      }
      // console.log("execs:", execs, "events:", events, "archivedEvents:", archivedEvents, "archivedExecutiveConfig:", archivedExecutiveConfig)

      const execLibraries = []
      
      //Add archived events to the events list
      events = [...events, ...archivedEvents]    

      // Connect execs to events
      events.forEach(evt => {
        evt.execs = evt.execs || []
        evt.principalIds.forEach(principalId => {
          for (let i = 0; i < execs.length; i += 1) {
            const exec = execs[i]
            if (principalId === exec.groupId || principalId === exec.archiveGroupId) {
                evt.execs.push({ ...exec })
              return
            }
          }
        })
        evt.execs = uniqBy(evt.execs, 'name')
        evt.execs = orderBy(evt.execs, 'name')
      })

      const selectedExecIds = execs.map(exec => exec.id)

      this.setState({
        loading: false,
        events,
        execs,
        selectedExecIds
      })
    })
    .catch(err => {
      // console.error("componentDidMount() err:", err)
    })
  }
  
  public render() {
    if (this.state.loading) {
      return <Loading />
    }

    const { 
      events,
      execs,
      selectedExecIds,
      searchText
    } = this.state

    return (
      <div className={styles.FutureEventsPage}>
        <div className={styles.flex}>
          <div className={styles.flexGrow}>
            <Breadcrumb
              items={[
                { key: "crumb0", text: "Events", isCurrentItem: true }
              ]} />
            <ExecFilter
              execs={execs} 
              onChange={execIds => {
                // console.log("new selectedExecIds:", execIds)
                this.setState({ selectedExecIds: execIds })
              }} />
          </div>
          <div className={styles.flexColumn} style={{ paddingBottom: "12px" }}>
            <div className={styles.flexGrow}>
            </div>
            <SearchBox
              style={{ width: "200px" }}
              onChange={val => this.setState({ searchText: val })}/>
          </div>
        </div>

        <PastEventsList
          events={events}
          filters={[
            this.filterByExec.bind(this),
            this.filterBySearchText.bind(this)
          ]}
          onEventClick={evt => {
            if(!evt.externalLink) {
              redirectToSitePage("Event", { eventId: evt.id })
            }
            else {
              //location.assign(evt.externalLink)
              window.open('','_new').location.href = evt.externalLink
            }
          }} />
      </div>
    );
  }

  private filterByExec(evt: IEvent) {
    return intersection(
      this.state.selectedExecIds,
      evt.execs.map(exec => exec.id)              
    ).length > 0
  }

  private filterBySearchText(evt: IEvent) {
    const searchText = this.state.searchText.toLowerCase()
    return (
      evt.eventName.toLowerCase().indexOf(searchText) >= 0 ||
      evt.location.toLowerCase().indexOf(searchText) >= 0 ||
      evt.startDate.format("YYYY-MM-DD").toLowerCase().indexOf(searchText) >= 0 ||
      evt.endDate.format("YYYY-MM-DD").toLowerCase().indexOf(searchText) >= 0
    )
  }
}
