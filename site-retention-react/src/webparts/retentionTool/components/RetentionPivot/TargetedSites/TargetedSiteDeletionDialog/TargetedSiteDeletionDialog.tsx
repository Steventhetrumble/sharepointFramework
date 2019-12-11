import * as React from 'react';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { getId } from 'office-ui-fabric-react/lib/Utilities';
import { hiddenContentStyle, mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { ContextualMenu } from 'office-ui-fabric-react/lib/ContextualMenu';
import { ITargetedSiteDeletionDialogProps, ITargetedSiteDeletionDialogState} from "./TargetedSiteDeletionDialog.types";

const screenReaderOnly = mergeStyles(hiddenContentStyle);



export class TargetedSiteDeletionDialog extends React.Component<ITargetedSiteDeletionDialogProps, ITargetedSiteDeletionDialogState> {
    constructor(props: ITargetedSiteDeletionDialogProps) {
        super(props);
        this.state = {
            targetName: ""
        };
       
    }

    // Use getId() to ensure that the IDs are unique on the page.
    // (It's also okay to use plain strings without getId() and manually ensure uniqueness.)
    private _labelId: string = getId('dialogLabel');
    private _subTextId: string = getId('subTextLabel');
    private _dragOptions = {
        moveMenuItemText: 'Move',
        closeMenuItemText: 'Close',
        menu: ContextualMenu
    };

    public componentDidUpdate(previousProps: ITargetedSiteDeletionDialogProps){
        if(previousProps.siteSelected != this.props.siteSelected  && this.props.siteSelected != null){
            this.setState({
                targetName: this.props.siteSelected.Title 
            });
        }
    }
   
    public render() {
        return (
            <div>

                <Dialog
                    hidden={!this.props.showDeletionDialog}
                    onDismiss={this.props.onDialogClose}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Delete Target',
                        subText: 'Do you want to remove ' + this.state.targetName + ' from the target deletion list?'
                    }}
                    modalProps={{
                        titleAriaId: this._labelId,
                        subtitleAriaId: this._subTextId,
                        isBlocking: false,
                        styles: { main: { maxWidth: 450 } }
                    }}
                >
                    <DialogFooter>
                        <PrimaryButton onClick={this.props.onDeleteConfirmation} text="Delete" />
                        <DefaultButton onClick={this.props.onDialogClose} text="Cancel" />
                    </DialogFooter>
                </Dialog>
            </div>
        );
    }

}
