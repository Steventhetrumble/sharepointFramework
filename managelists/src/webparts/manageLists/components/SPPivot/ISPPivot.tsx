import * as React from 'react';
import { Label, ILabelStyles } from 'office-ui-fabric-react/lib/Label';
import { Pivot, PivotItem } from 'office-ui-fabric-react/lib/Pivot';
import { IStyleSet } from 'office-ui-fabric-react/lib/Styling';
import { SPDetailList } from '../SPDetailList/SPDetailList';
import { SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
import {
    Result,
    RelevantResults,
    Row
} from '../IManageLists';
import {
    IList,
    ISPSite
} from '../../common/IObjects';
import {
    Version,
    Environment,
    EnvironmentType
} from '@microsoft/sp-core-library';
import { IDataProvider } from '../../dataproviders/IDataProvider';
import styles from '../ManageLists.module.scss';
import {SPTargetedSitesDetailList} from '../SPTargetedSitesDetailList/SPTargetedSitesDetailList';
import {FirstRoundEmail} from './FirstRoundEmail/FirstRoundEmail';
import {SecondRoundEmail} from './SecondRoundEmail/SecondRoundEmail';




const labelStyles: Partial<IStyleSet<ILabelStyles>> = {
    root: { marginTop: 10 }
};

export interface ISPPivotState {
    lists: IList[];
    rootSite: ISPSite;

}

export interface ISPPivotProps {
    provider: IDataProvider;
    site: ISPSite;
    spHttpClient: SPHttpClient;
}

export class SPPivot extends React.Component<ISPPivotProps, ISPPivotState> {
    

    constructor(props: ISPPivotProps) {
        super(props);
        
        this.state = {
            lists: [],
            rootSite: this.props.site

        };

    }

    public render() {

        return (<Pivot>
            <PivotItem
                headerText="Search List"
                headerButtonProps={{
                    'data-order': 1,
                    'data-title': 'Search List'
                }}
            >
                <SPDetailList provider={this.props.provider} site={this.props.site} spHttpClient={this.props.spHttpClient}></SPDetailList>
            </PivotItem>
            <PivotItem headerText="Targeted Sites">
                <SPTargetedSitesDetailList spHttpClient={this.props.spHttpClient}></SPTargetedSitesDetailList>               
            </PivotItem>
            <PivotItem headerText="1st Round Email">
                <FirstRoundEmail spHttpClient={this.props.spHttpClient} />
            </PivotItem>
            <PivotItem headerText="2nd Round Email">
                <SecondRoundEmail spHttpClient={this.props.spHttpClient} />
            </PivotItem>
            <PivotItem headerText="Developer">
                <Label styles={labelStyles}>name: Steven Trumble</Label>
                <Label styles={labelStyles}>email:StevenAndrew.Trumbl1@rci.rogers.com</Label>
                <Label styles={labelStyles}>Alternate email:Steven.a.trumble@gmail.ca</Label>
            </PivotItem>
        </Pivot>
        );
    }
}
