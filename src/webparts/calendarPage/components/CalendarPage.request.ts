import  { siteCollectionUrl } from './../../../shared/SharePoint'
import * as moment from 'moment'
import { getCalendarSOAPXML } from './CalendarPage.soapXml'
import pnp, { Web } from 'sp-pnp-js'
import * as parse from 'xml-parser';
// console.log(parse)
// const readFileSync = require('fs').readFileSync;

const web = new Web(siteCollectionUrl);
const executiveEventsList = 'Executive Events';
const eventUrl = `${siteCollectionUrl}/sitepages/Event.aspx?eventId=`

//TODO: use async await to get executive events id and set it to state on calendar component
export function calendarIdRequest() {
  return web.lists.getByTitle(executiveEventsList)
  .get()
  .then(items => {
    return items.Id;
  })
}


export function getCalendarEventsViaRest() {
  return web.lists.getByTitle(executiveEventsList).items
  .expand(
    'RoleAssignments',
    'RoleAssignments/Member',
    'RoleAssignments/RoleDefinitionBindings',
  )
  .get()
  .then(items => {
    // console.log(items, 'FUTURE EVENTS VIA REST')
    return items.map(item => {

      return {
        id: item.Id,
        title: item.Title,
        // startDate: item.EventDate ? moment(item.EventDate) : '',
        // endDate: item.EndDate ? moment(item.EndDate) : '',
        // location: item.Location || '',
        principalIds: item.RoleAssignments && item.RoleAssignments.map(ra => ra.Member.Id) || [],

      }
    })
    // return items;
  })
}

export default function getCalendarEvents() {
  return new Promise((resolve, reject) => {
    getCalendarSOAPXML().then((soapXML) => {
      const url = `${siteCollectionUrl}/_vti_bin/Lists.asmx`;
      // alert(url + ' LOCATION')
      const xhr = new XMLHttpRequest();
      xhr.open('POST', url, true);
      xhr.setRequestHeader('Content-Type', 'text/xml; charset=utf-8');
      xhr.setRequestHeader('SOAPAction','http://schemas.microsoft.com/sharepoint/soap/GetListItems');
      xhr.onreadystatechange  = () => {
        if (xhr.readyState === 4) {
          if (xhr.status === 200) {
            // console.log(xhr.response, 'what is response')
            // console.log(parse(xhr.response), 'WHAT IS HERE?')
            let parsedData = parse(xhr.response);
            // console.log(parsedData)

            let parsedEvents = parsedData.root.children[0].children[0].children[0].children[0].children[0].children.map(e => e.attributes)

            let events = parsedEvents.reduce((arr, evt) => {
              let event = {};
              let id = parseInt(evt.ows_ID, 10);
              event['id'] = id;
              event['allDay'] = parseInt(evt.ows_fAllDayEvent) ? parseInt(evt.ows_fAllDayEvent, 10) : 0;
              event['title'] = evt.ows_Title;
              event['start'] = new Date(evt.ows_EventDate);
              if(evt.ows_EndDate) {
                event['end'] = new Date(evt.ows_EndDate);
              }
              else {
                event['end'] = new Date(evt.ows_EventDate)
              }
              event['url'] = `${eventUrl}${id}`;

              arr.push(event);
              return arr;
            }, [])

            resolve(events)
          } else {
            reject(() => {
              // console.warn(xhr.status)
            })
          }
        }
      }
      xhr.send(soapXML)
    })
  })
}