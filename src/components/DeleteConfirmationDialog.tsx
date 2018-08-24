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
    PrimaryButton,
    DefaultButton
  } from 'office-ui-fabric-react/lib/Button'

import {
deleteItem,
IFSObject
} from '../shared/SharePoint'

import styles from './TextInputDialog.module.scss'

//export interface IDeleteConfirmationDialogForm {
export interface IDeleteConfirmationDialogState {
    item: IFSObject
}
  
export interface IDeleteConfirmationDialogProps {
    item?: IFSObject
    onDelete: () => void
    onDismiss: () => void
}

//export interface IDeleteConfirmationDialogState extends IDeleteConfirmationDialogForm {}

export default class DeleteConfirmationDialog extends React.Component<IDeleteConfirmationDialogProps, IDeleteConfirmationDialogState> {
    constructor(props: IDeleteConfirmationDialogProps) {
        super(props)

        this.state = {
            item: this.props.item
        } 

        this.deleteItem = this.deleteItem.bind(this)
        this.dismiss = this.dismiss.bind(this)

        
    }
    private deleteItem() {
        if (this.props.onDelete) {
            this.props.onDelete()
        }
        this.dismiss()
    }

    private dismiss() {
        if (this.props.onDismiss) {
            this.props.onDismiss()
        }
    }

    public render() {
        return (
            <Dialog 
                hidden={false}
                modalProps={{
                    className: styles.TextInputDialog
                }}
               
                dialogContentProps={{
                    type: DialogType.normal,
                    title: `Delete ` + this.state.item.name + `?`
                }}
                onDismiss={this.dismiss}>
                <div className="confirm-message">
                    <Label>Are you sure you want to sent the item(s) to the site Recycle Bin?</Label>
                </div>
                <DialogFooter>
                    <PrimaryButton text="Delete"
                        onClick={this.deleteItem} />
                    <DefaultButton text="Cancel"
                        onClick={this.dismiss} />
                </DialogFooter>
            </Dialog>
        )
    }
} 
