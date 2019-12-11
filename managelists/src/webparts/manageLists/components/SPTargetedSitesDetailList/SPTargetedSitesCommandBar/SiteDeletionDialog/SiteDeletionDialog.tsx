import * as React from 'react';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { getId } from 'office-ui-fabric-react/lib/Utilities';
import { hiddenContentStyle, mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { ContextualMenu } from 'office-ui-fabric-react/lib/ContextualMenu';
import { ListRow } from '../../SPTargetedSitesDetailList';
import { SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';

const screenReaderOnly = mergeStyles(hiddenContentStyle);

export interface ISiteDeletionDialogProps {
    callback(): void;
    spHttpClient: SPHttpClient;
    siteSelected: ListRow;
    showDialog: boolean;
}

export interface ISiteDeletionDialogState {
    hideDialog: boolean;
    isDraggable: boolean;
}

export class SiteDeletionDialog extends React.Component<ISiteDeletionDialogProps, ISiteDeletionDialogState> {
    constructor(props: ISiteDeletionDialogProps) {
        super(props);
        this.state = {
            hideDialog: true,
            isDraggable: false
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

    public componentDidUpdate(previousProps: ISiteDeletionDialogProps, previousState: ISiteDeletionDialogState) {
        if (!previousProps.showDialog && this.props.showDialog) {
            this._showDialog();
        }
    }

    public render() {
        const { hideDialog, isDraggable } = this.state;
        return (
            <div>

                <Dialog
                    hidden={hideDialog}
                    onDismiss={this._closeDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Delete Target',
                        subText: 'Do you want to remove this site from the target deletion list??'
                    }}
                    modalProps={{
                        titleAriaId: this._labelId,
                        subtitleAriaId: this._subTextId,
                        isBlocking: false,
                        styles: { main: { maxWidth: 450 } },
                        dragOptions: isDraggable ? this._dragOptions : undefined
                    }}
                >
                    <DialogFooter>
                        <PrimaryButton onClick={this._deleteSiteFromList} text="Delete" />
                        <DefaultButton onClick={this._closeDialog} text="Cancel" />
                    </DialogFooter>
                </Dialog>
            </div>
        );
    }

    private _showDialog = (): void => {
        this.setState({ hideDialog: false });
    }

    private _closeDialog = (): void => {
        this.setState({ hideDialog: true });
    }

    private _deleteSiteFromList = (): void => {
        console.log(this.props.siteSelected["@odata.editLink"]);
        this.props.spHttpClient.post(`https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/_api/${this.props.siteSelected['@odata.editLink']}`, SPHttpClient.configurations.v1, {
            headers: {
                'IF-MATCH': '*',
                'X-HTTP-Method': 'DELETE'
            }
        })
            .then((response: SPHttpClientResponse) => {
                if (response.ok) {
                    console.log("Delete response:", response.status);
                    this._closeDialog();
                    this.props.callback();
                }
            });
        this._closeDialog();
    }



}
