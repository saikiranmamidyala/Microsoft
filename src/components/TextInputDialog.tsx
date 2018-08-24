import * as React from 'react'

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
  PrimaryButton,
  DefaultButton
} from 'office-ui-fabric-react/lib/Button'


import styles from './TextInputDialog.module.scss'

export interface ITextInputDialogForm {
  value: string
  title: string
  extension?: string
}

export interface ITextInputDialogProps {
  value: string
  title: string
  onSave: (value: string) => void
  onDismiss?: () => void
  extension?: string
}

export interface ITextInputDialogState extends ITextInputDialogForm { }

export default class TextInputDialog extends React.Component<ITextInputDialogProps, ITextInputDialogState> {
  constructor(props: ITextInputDialogProps) {
    super(props)

    this.state = {
      value: "",
      title: "",
      extension: ""
    }

    this.save = this.save.bind(this)
    this.dismiss = this.dismiss.bind(this)
    this.formValid = this.formValid.bind(this)

  }
  private save() {
    if (this.props.onSave) {
      this.props.onSave(this.state.value as string)
      this.dismiss()
    }
  }

  private dismiss() {
    if (this.props.onDismiss) {
      this.props.onDismiss()
    }
  }

  private formValid(): boolean {
    return !!(
      this.state.value
    )
  }

  public render() {

    return (
      <Dialog hidden={false}
        modalProps={{
          className: styles.TextInputDialog
        }}
        dialogContentProps={{
          type: DialogType.normal,
          title: this.props.title
        }}
        onDismiss={this.dismiss}>
        <div className="input-container">
          <TextField
            onChanged={text => this.setState({ value: text.trim() })}
            autoFocus={true}
            onKeyUp={evt => {
              if (evt.which === 13) {
                this.save()
              }
            }}
            
          />
          {this.props.extension && 
            <div className="extensionString">.{this.props.extension}</div>
          }
        </div>
        <DialogFooter>
          <PrimaryButton text="Save"
            disabled={this.formValid() ? false : true}
            onClick={this.save} />
          <DefaultButton text="Cancel"
            onClick={this.dismiss} />
        </DialogFooter>
      </Dialog>
    )
  }
} 
