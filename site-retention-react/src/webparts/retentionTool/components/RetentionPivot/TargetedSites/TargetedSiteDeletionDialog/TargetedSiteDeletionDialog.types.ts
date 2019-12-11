import { ListRow } from "../../../common/IObjects";

export interface ITargetedSiteDeletionDialogProps {
    onDeleteConfirmation(): void;
    onDialogClose(): void;
    siteSelected: ListRow;
    showDeletionDialog: boolean;
}

export interface ITargetedSiteDeletionDialogState {
    targetName: string;
}