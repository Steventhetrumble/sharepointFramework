import { DefaultPalette, FontSizes, FontWeights } from "@uifabric/styling";

import {ILabelStyles, IStackStyles, IPivotItemProps} from 'office-ui-fabric-react';


export const labels: ILabelStyles = {
    root: {
        backgroundColor: DefaultPalette.yellowLight,
        fontSize: FontSizes.mediumPlus,
        fontWeight: FontWeights.bold
    }
}; 

export const contentContainer: IStackStyles = {
    
    root: {
        minHeight: "40vh",
        maxHeight: "40vh",
        position: "relative"
    }
};