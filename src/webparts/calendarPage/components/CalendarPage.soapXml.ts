import  { getListGUID } from './../../../shared/SharePoint'

//export async function  getCalendarSOAPXML(month:Date): Promise<string> {
export async function  getCalendarSOAPXML(): Promise<string> {
  let calendarGUID = await getListGUID("Executive Events")

let soapXML = "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'> <soap:Body>";
soapXML += "<GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>";
soapXML += "<listName>" +  calendarGUID + "</listName>";
soapXML += "<viewFields>";
soapXML += "<ViewFields xmlns=\"ows_UniqueId\" >";
soapXML += "<FieldRef Name=\"FileRef\" />";
soapXML += "<FieldRef Name=\"FileLeafRef\" />";
soapXML += "<FieldRef Name=\"fRecurrence\" />";
soapXML += "<FieldRef Name=\"RecurrenceData\" />";
soapXML += "<FieldRef Name=\"EventDate\" />";
soapXML += "<FieldRef Name=\"EndDate\" />";
soapXML += "<FieldRef Name=\"Title\" />";
soapXML += "<FieldRef Name=\"Comments\" />";
soapXML += "<FieldRef Name=\"Description\" />";
soapXML += "<FieldRef Name=\"Location\" />";
soapXML += "<FieldRef Name=\"Remote_x0020_Link\" />";
soapXML += "<FieldRef Name=\"EventType0\" />";
soapXML += "<FieldRef Name=\"ID\" />";
soapXML += "</ViewFields>";
soapXML += "</viewFields>";
soapXML += "<rowLimit>5000</rowLimit>"
soapXML += "<query>";
soapXML += "<Query>";
soapXML += "<Where>";
soapXML += "<DateRangesOverlap>";
soapXML += "<FieldRef Name=\"EventDate\" />";
soapXML += "<FieldRef Name=\"EndDate\" />";
soapXML += "<FieldRef Name=\"RecurrenceID\" />";
soapXML += "<Value Type=\"DateTime\"><Year /></Value>";
soapXML += "</DateRangesOverlap>";
soapXML += "</Where>";
soapXML += "<OrderBy>";
soapXML += "<FieldRef Name=\"EventDate\" Ascending=\"True\" />";
soapXML += "</OrderBy>";
soapXML += "</Query>";
soapXML += "</query>";
soapXML += "<queryOptions>";
soapXML += "<QueryOptions>";
soapXML += "<CalendarDate>" + new Date().toISOString() + "</CalendarDate>";  //e.g. 2017-10-11T14:40:35
soapXML += "<RecurrencePatternXMLVersion>v3</RecurrencePatternXMLVersion>";
soapXML += "<ExpandRecurrence>TRUE</ExpandRecurrence>";
soapXML += "</QueryOptions>";
soapXML += "</queryOptions>";
soapXML += "</GetListItems>";
soapXML += "</soap:Body></soap:Envelope>";

return soapXML;
}