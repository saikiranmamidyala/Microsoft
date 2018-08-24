//#region imports

import * as React from 'react'
import { findDOMNode } from 'react-dom'

import {
  Dialog,
  DialogType,
  DialogFooter
} from 'office-ui-fabric-react/lib/Dialog'

import {
  Label
} from 'office-ui-fabric-react/lib/Label'

import {
  TextField
} from 'office-ui-fabric-react/lib/TextField'

import {
  DatePicker,
  ICalendarStrings
} from 'office-ui-fabric-react/lib/DatePicker'

import {
  IPersonaProps
} from 'office-ui-fabric-react/lib/components/Persona'

import {
  NormalPeoplePicker
} from 'office-ui-fabric-react/lib/Pickers'

import {
  PrimaryButton,
  DefaultButton
} from 'office-ui-fabric-react/lib/Button'

import {
  Spinner,
  SpinnerSize
} from 'office-ui-fabric-react/lib/Spinner'

import {
  Icon
} from 'office-ui-fabric-react/lib/Icon'

import PeoplePicker, {
  IPerson
} from '../components/PeoplePicker'

import ExecSelector from '../components/ExecSelector'

import {
  searchPeople,
  IExecutive,
  createNewEvent,
  getAllExecutives,
  IEvent,
  getEventById,
  updateEventMetadata
} from '../shared/SharePoint'

import { find } from 'lodash';

import styles from './EventDetailsDialog.module.scss'
import { autobind } from 'office-ui-fabric-react/lib/Utilities';

//#endregion
//#region EventForm

export interface IEventFormValues {
  id?: number
  execs?: IExecutive[]
  eventName?: string
  startDate?: Date
  endDate?: Date
  location?: string
  commsManagers?: IPerson[]
}

export interface IEventFormProps {
  values?: IEventFormValues
  execs: IExecutive[]
  onDismiss: () => void
  onSave: (event: IEventFormValues) => void
}

export interface IEventFormState {
  event: IEventFormValues,
  eventNameErrorMessage: string
}

export class EventForm extends React.Component<IEventFormProps, IEventFormState> {
  constructor(props: IEventFormProps) {
    super(props)

    this.state = {
      event: {
        execs: [],
        eventName: "",
        startDate: null,
        endDate: null,
        location: "",
        commsManagers: [],
        ...(props.values || {})
      },
      eventNameErrorMessage: ""
    }
  }

  public render() {
    return (
      <div>
        <div className={styles.NewEventForm}>
          <div className={styles.row}>
            <div className={styles.label}>
              <Label required={true}>Exec(s)</Label>
            </div>
            <div className={styles.input}>
              <ExecSelector
                execs={this.props.execs}
                selectedExecIds={this.state.event.execs.map(x => x.id)}
                //disabled={true}
                disabled={false}
                //onChange={this.onExecSelectionChange}
                onChange={this.onExecSelectionChange}
              />
            </div>
          </div>
          <div className={styles.row}>
            <div className={`${styles.label} ${styles.centerVert}`}>
              <Label required={true}>Event Name</Label>
            </div>
            <div className={styles.input}>
              <TextField
                onChanged={this.onEventNameChange}
                value={this.state.event.eventName}
                //disabled={true}
                disabled={false}
                errorMessage={this.state.eventNameErrorMessage} />
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.label}>
              <Label required={true}>Start Date</Label>
            </div>
            <div className={styles.input}>
              <DatePicker
                value={this.state.event.startDate}
                showMonthPickerAsOverlay={true}
                showGoToToday={false}
                placeholder="Select a start date..."
                onSelectDate={this.onStartDateChange} />
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.label}>
              <Label required={false}>End Date</Label>
            </div>
            <div className={styles.input}>
              <DatePicker
                value={this.state.event.endDate}
                showMonthPickerAsOverlay={true}
                showGoToToday={false}
                placeholder="Select an end date..."
                onSelectDate={this.onEndDateChange} />
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.label}>
              <Label required={false}>Location</Label>
            </div>
            <div className={styles.input}>
              <TextField
                placeholder="City, State, Country"
                onChanged={this.onLocationChange}
                value={this.state.event.location} />
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.label}>
              <Label required={false}>Comms Manager(s)</Label>
            </div>
            <div className={styles.input}>
              <PeoplePicker
                placeholder="Start typing a name..."
                onChange={this.onCommsManagersChange}
                selectedItems={this.state.event.commsManagers} />
            </div>
          </div>
        </div>
        <DialogFooter>
          <PrimaryButton text="Save"
            disabled={this.requiredFieldsHaveValues() ? false : true}
            onClick={this.save} />
          <DefaultButton text="Cancel"
            onClick={this.dismiss} />
        </DialogFooter>
      </div>
    )
  }

  @autobind
  private onExecSelectionChange(execs: IExecutive[]) {
    this.setState(state => ({ ...state, event: { ...state.event, execs } }))
  }

  @autobind
  private async onEventNameChange(eventName: string) {
    if (eventName.indexOf("/") >= 0) {
      this.setState({ eventNameErrorMessage: "/ (forward slash) is not allowed. How about using a - (hyphen) instead?" })
      return
    }
    
    this.setState(state => ({ ...state, event: { ...state.event, eventName }, eventNameErrorMessage: "" }))
  }

  @autobind
  private onStartDateChange(startDate: Date) {
    this.setState(state => ({ ...state, event: { ...state.event, startDate } }))
  }

  @autobind
  private onEndDateChange(endDate: Date) {
    this.setState(state => ({ ...state, event: { ...state.event, endDate } }))
  }

  @autobind
  private onLocationChange(location: string) {
    this.setState(state => ({ ...state, event: { ...state.event, location } }))
  }

  @autobind
  private onCommsManagersChange(commsManagers: IPerson[]) {
    this.setState(state => ({ ...state, event: { ...state.event, commsManagers } }))
  }

  @autobind
  private save() {
    this.props.onSave(this.state.event)
  }

  @autobind
  private dismiss() {
    if (this.props.onDismiss) this.props.onDismiss()
  }

  @autobind
  private requiredFieldsHaveValues(): boolean {
    // These are the required fields
    const { execs, eventName, startDate, endDate } = this.state.event

    const validEndDate:boolean = (endDate !== null && endDate >= startDate) || endDate === null ? true:false;

    return (!!(
      execs &&
      execs.length &&
      eventName &&
      startDate
    ) && (
      validEndDate
    ))
  }
}

//#endregion
//#region SavingEvent

export interface ISavingEventViewProps {
  message: string
}

export class SavingEventView extends React.PureComponent<ISavingEventViewProps> {
  public render(): JSX.Element {
    return (
      <div>
        <div className={styles.SavingEvent}>
          <div className={styles.moveUp}>
            <Spinner size={SpinnerSize.large} />
            <p className="ms-font-xl">{this.props.message}</p>
          </div>
        </div>
        <DialogFooter></DialogFooter>      
      </div>
    )
  }
}

//#endregion
//#region SavedEvent

export interface ISavedEventViewProps {
  message: string
  showGoToEvent: boolean
  onGoToEvent?: () => void
  onClose: () => void
}

export class SavedEventView extends React.PureComponent<ISavedEventViewProps> {
  public render(): JSX.Element {
    return (
      <div>
        <div className={styles.SavedEvent}>
          <div className={styles.moveUp}>
            <Icon iconName="EventAccepted" style={{ color: "rgba(0,255,0,0.7)" }} />
            <p className="ms-font-xl">{this.props.message}</p>
          </div>
        </div>
        <DialogFooter>
          {this.props.showGoToEvent && (
            <PrimaryButton text="Take me to my event"
            onClick={() => this.goToEvent()} />
          )}
          <DefaultButton text="Close"
          onClick={() => this.close()} />
            {/* onClick={() => this.close()} */}
        </DialogFooter>
      </div>
    )
  }

  private goToEvent() {
    if (this.props.onGoToEvent) {
      this.props.onGoToEvent()
    }
  }

  private close() {
    this.props.onClose()
  }
}

//#endregion
//#region ErrorView

export interface IErrorViewProps {
  message: string
  onRetry: () => void
  onClose: () => void
}

export class ErrorView extends React.PureComponent<IErrorViewProps> {
  public render(): JSX.Element {
    return (
      <div>
        <div className={styles.ErrorView}>
          <div className={styles.moveUp}>
            <Icon iconName="ErrorBadge" style={{ color: "rgba(255,0,0,0.7)" }} />
            <p className="ms-font-xl">{this.props.message}</p>
          </div>
        </div>
        <DialogFooter>
          {/*
          <PrimaryButton text="Retry"
            onClick={() => this.retry()} />
          */}
          <DefaultButton text="Close"
            onClick={() => this.close()} />          
        </DialogFooter>
      </div>      
    )
  }

  private retry() {
    this.props.onRetry()
  }

  private close() {
    this.props.onClose()
  }
}

//#endregion
//#region EventDetailsDialog

export enum View {
  loading,
  input,
  saving,
  saved,
  error
}

export interface IEventDetailsDialogProps {
  eventId?: number
  onSuccess?: (newEventId?: number) => void
  onDismiss?: () => void
}

export interface IEventDetailsDialogState {
  view: View
  title: string
  showCloseButton: boolean
  event: IEventFormValues
  message: string
  isBlocking: boolean
  newEventId: number
}

export default class EventDetailsDialog extends React.Component<IEventDetailsDialogProps, IEventDetailsDialogState> {
  private _execs: IExecutive[] = []

  private getInitialState(): IEventDetailsDialogState {
    return {
      view: View.loading,
      title: "Create a new event",
      showCloseButton: true,
      event: null,
      message: "",
      isBlocking: false,
      newEventId: -1
    }
  }

  constructor(props: IEventDetailsDialogProps) {
    super(props)
    this.state = this.getInitialState()
      //this.getTestStateForView(NewEventDialogView.saved)
  }

  public async componentDidMount() {
    this._execs = await getAllExecutives()

    if (this.props.eventId) {
      const event = await getEventById(this.props.eventId)

      this.setState(state => ({
        ...state,
        view: View.input,
        event: {
          id: event.id,
          execs: this._execs.filter(exec => event.principalIds.indexOf(exec.groupId) >= 0),
          eventName: event.eventName,
          startDate: event.startDate ? event.startDate.toDate() : null,
          endDate: event.endDate ? event.endDate.toDate() : null,
          location: event.location,
          commsManagers: event.commsManagers.map(x => ({
            id: x.Id,
            primaryText: x.Title,
            secondaryText: x.JobTitle,
            tertiaryText: x.Department,
            imageShouldFadeIn: true,
            imageUrl: "/_layouts/15/userphoto.aspx?size=S&accountname=" + x.EMail
          } as IPerson))
        } as IEventFormValues
      }))
    } else {
      this.setState(state => ({
        ...state,
        view: View.input,
        event: {
          id: -1,
          execs: [],
          eventName: "",
          startDate: null,
          endDate: null,
          location: "",
          commsManagers: []
        } as IEventFormValues
      }))
    }
  }

  public render(): JSX.Element {
    return (
      <Dialog hidden={false}
        modalProps={{
          isBlocking: this.state.isBlocking,
          className: styles.NewEventDialog
        }}
        dialogContentProps={{
          type: DialogType.normal,
          title: this.state.title,
          showCloseButton: this.state.showCloseButton
        }}
        onDismiss={() => this.dismiss()}
      >
        {this.renderContent.bind(this)()}
      </Dialog>
    )
  }

  private renderContent(): JSX.Element {
    switch (this.state.view) {
      case View.loading: return (
        <div className={styles.dialogLoading}>
          <Spinner size={SpinnerSize.large} />
        </div>
      )
      case View.input: return (
        <EventForm
          execs={this._execs}
          values={this.state.event}
          onDismiss={() => this.dismiss()}
          onSave={event => this.valuesSubmitted(event)} />
      )
      case View.saving: return (
        <SavingEventView
          message={this.state.message} />
      )
      case View.saved: return (
        <SavedEventView
          message={this.state.message}
          // onClose={() => this.success(this.state.newEventId)}
          onClose={() =>this.dismiss() }
          showGoToEvent={!this.props.eventId}
          onGoToEvent={() => this.success(this.state.newEventId)} />
      )
      case View.error: return (
        <ErrorView
          message={this.state.message}
          onClose={() => this.dismiss()}
          onRetry={() => {
            // TODO
          }} />
      )
      default: return null
    }
  }

  private valuesSubmitted(event: IEventFormValues) {
    this.setState(state => ({
      ...state,
      event,
      title: "",
      view: View.saving,
      showCloseButton: false,
      message: "...",
      isBlocking: true
    }) as IEventDetailsDialogState)

    if (this.props.eventId) {
      // Update
      updateEventMetadata({
        event,
        progress: message => {
          this.setState(state => ({ ...state, message }))
        }
      })
      .then(() => {
        this.setState(state => ({
          ...state,
          title: "Everything looks good.",
          view: View.saved,
          showCloseButton: true,
          message: "Your event has been updated.",
          isBlocking: false
        }))
      })
      .catch(() => {
        // TODO: Get more specific error message
        this.setState(state => ({
          ...state,
          title: "Error",
          view: View.error,
          showCloseButton: true,
          message: "Something went wrong.",
          isBlocking: false
        }))
      })

      // Steps
      /*
        - For each exec
          - If removing
            - Remove access to exec's folder, event folder, and shared folder
          - Else if adding
            - Check if exec's folder already exists
              (in the case where they were previously removed)
            - If exists
              - Add access back to exec's folder, event folder, and shared folder
            - Else
              - Create new folder for exec
              - Add access to new folder, event folder, and shared folder
        - Update calendar event metadata
      */

    } else {
      // Create
      createNewEvent({
        event,
        progress: message => {
          this.setState(state => ({ ...state, message }))
        }
      })
      .then(eventId => {
        this.setState(state => ({
          ...state,
          title: "Everything looks good.",
          view: View.saved,
          showCloseButton: true,
          message: "Your event has been created.",
          isBlocking: false,
          newEventId: eventId
        }))
      })
      .catch(() => {
        // TODO: Get more specific error message
        this.setState(state => ({
          ...state,
          title: "Error",
          view: View.error,
          showCloseButton: true,
          message: "Something went wrong.",
          isBlocking: false
        }))
      })
    }
  }

  /*
  private getTestStateForView(view: View): IEventDetailsDialogState {
    switch (view) {
      case View.input: return this.getInitialState()
      case View.saving: return {
        event: null,
        title: "",
        view: View.saving,
        showCloseButton: false,
        message: "Generating folders...",
        isBlocking: true,
        newEventId: -1
      }
      case View.saved: return {
        event: null,
        title: "",
        view: View.saved,
        showCloseButton: true,
        message: "Your event has been created.",
        isBlocking: false,
        newEventId: -1
      }
      case View.error: return {
        event: null,
        title: "",
        view: View.error,
        showCloseButton: true,
        message: "Something went wrong.",
        isBlocking: false,
        newEventId: -1
      }
    }
  }
  */

  private success(newEventId?: number) {
    if (this.props.onSuccess) this.props.onSuccess(newEventId)
  }

  private dismiss() {
    if (this.props.onDismiss) this.props.onDismiss()
  }
}

//#endregion
