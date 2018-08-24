import * as React from 'react';
import { Breadcrumb, IBreadcrumb, IBreadcrumbItem, IBreadCrumbData, IBreadcrumbProps } from 'office-ui-fabric-react/lib/Breadcrumb'


import { ICalendarPageProps } from './ICalendarPageProps';
import ExecFilter from '../../../components/ExecFilter';
import { getAllExecutives, IExecutive, siteCollectionUrl, cdnAssetsBaseUrl } from '../../../shared/SharePoint'

import BigCalendar from 'react-big-calendar';
import * as moment from 'moment';
BigCalendar.momentLocalizer(moment);
import 'react-big-calendar/lib/css/react-big-calendar.css';
import styles from './CalendarPage.module.scss';


import getCalendarEvents from './CalendarPage.request';
import { getCalendarEventsViaRest, calendarIdRequest } from './CalendarPage.request';

const allViews = Object.keys(BigCalendar.Views).map(k => BigCalendar.Views[k]);

export default class CalendarPage extends React.Component<ICalendarPageProps, {events, execs, execIds, originalEvents, breadCrumbs, calendarId}> {
  constructor(props) {
    super(props);

    let breadCrumbs = [ {
      key: 'crumb0',
      text: 'Calendar',
      isCurrentItem: true,
    } as IBreadcrumbItem ] 

    this.state = {
      events: [],
      execs: [],
      execIds: [],
      originalEvents: [],
      breadCrumbs,
      calendarId: '',
    }

    this.attachExecsToEvents = this.attachExecsToEvents.bind(this);
    this.filterEventsOnExecClick = this.filterEventsOnExecClick.bind(this);
    this.cacheOriginalEvents = this.cacheOriginalEvents.bind(this);
    this.attachPrincipalIdsToEvents = this.attachPrincipalIdsToEvents.bind(this);
    this.navigateToEventPage = this.navigateToEventPage.bind(this);
    this.syncEventsToOutlook = this.syncEventsToOutlook.bind(this);
  }

  private attachExecsToEvents(events, execs) {
   return events.map(e => {
      e.execs = e.execs || [];
      e.principalIds.forEach(id => {
        for (let i = 0; i < execs.length; i++) {
          const exec = execs[i];
          if (id === exec.groupId) {
            e.execs.push({...exec })
          }
        }
      })
      return e;
    })
  }

  private cacheOriginalEvents(events) {
    this.setState({
      originalEvents: events,
    })
  }
  
  public componentDidMount() {
    getCalendarEvents()
    .then(events => {
     this.setState({
        events,
     })

     calendarIdRequest()
     .then(id => this.setState({ calendarId: id }))
    })
    .then(() => {
      return getAllExecutives();
    })
    .then(execs => {
      this.setState({
        execs: execs.map(exec => exec as IExecutive),
        // execIds: execs.map(exec => exec.groupId),
      })
      
      return getCalendarEventsViaRest();
    })
    .then(eventsViaRest => {
      let eventsWithPrincipalIds = this.attachPrincipalIdsToEvents(this.state.events, eventsViaRest)
      this.setState({
        events: this.attachExecsToEvents(eventsWithPrincipalIds, this.state.execs),
      })
      this.cacheOriginalEvents(this.state.events)
    })
    .catch(err => {
      //console.warn(err)
    })
  }

  private attachPrincipalIdsToEvents(stateEvents, eventsViaRest) {
    return stateEvents.map(stateEvent => {
      eventsViaRest.forEach(evt => {
        if (evt.id === stateEvent.id) {
          stateEvent.principalIds = evt.principalIds;
        }
      })
      return stateEvent;
    })
  }

  private filterEventsOnExecClick(execIds) {
    this.setState({
      execIds,
    })
    let selectedEvents = execIds.map(id => {
      return this.state.originalEvents.filter(e => {
        return e.execs.map(exec => exec.id).indexOf(id) > -1;
      })
    })
    .reduce((prev, cur) => prev.concat(cur))
    .filter((e, index, self) => self.indexOf(e) === index)
    this.setState({
      events: selectedEvents,
    })
  }

  private navigateToEventPage(calendarEvent) {
    window.open(calendarEvent.url, '_blank');
  }

  private syncEventsToOutlook() {
    let url = `stssync://sts/?ver=1.1&type=calendar&cmd=add-folder&base-url=${siteCollectionUrl}&list-url=%2FLists%2FExecutive%20Events&guid=%7B${this.state.calendarId}%7D&site-name=ExecCommsvNextDev&list-name=Executive%20Events`
    window.open(url)
  }

  
  public render(): React.ReactElement<ICalendarPageProps> {
    return (
      <div className={ styles.calendarPage }>
        <Breadcrumb
         items={ this.state.breadCrumbs } 
         className={ styles.currentBreadCrumb }
         />
        <ExecFilter 
        execs={ this.state.execs }
        onChange={ this.filterEventsOnExecClick }
        />
        <div className={ styles.syncOutlookButton }
        onClick={ this.syncEventsToOutlook }
        > 
          <div className={styles.imageContainer}>
            <img src={ `${cdnAssetsBaseUrl}/images/SyncToOutlook.png` } />
          </div>
          <h3>Sync to Outlook</h3>
        </div>
        <BigCalendar 
        events={ this.state.events }
        defaultDate={ new Date() }
        views={ allViews }
        onSelectEvent={ this.navigateToEventPage }
        />
      
      
    </div>
    );
  }
}
