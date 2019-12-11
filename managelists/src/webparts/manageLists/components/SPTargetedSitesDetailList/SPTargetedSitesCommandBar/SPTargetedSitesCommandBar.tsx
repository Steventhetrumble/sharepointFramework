import * as React from 'react';
import { ListRow, ISPTargetedSitesDetailListState } from '../SPTargetedSitesDetailList';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { CommandBarButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { DirectionalHint } from 'office-ui-fabric-react/lib/Callout';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SiteDeletionDialog } from './SiteDeletionDialog/SiteDeletionDialog';
import { SiteEditPanel } from './SiteEditPanel/SiteEditPanel';
import { ListResult } from '../../SPPivot/FirstRoundEmail/FirstRoundEmail';
import { Modal } from 'office-ui-fabric-react';
import styles from '../../ManageLists.module.scss';


export interface ISPTargetedSitesCommandBarProps {
    spHttpClient: SPHttpClient;
    siteIsSelected: boolean;
    selectedSite: ListRow;
    callerback(): void;
}
export interface ISPTargetedSitesCommandBarState {
    showDeletionDialog: boolean;
    showEditPanel: boolean;
    firstEmail: string;
    emailRecipients: emailGroups;
    showEmail1Modal: boolean;
}

export interface emailGroups{
    [title: string] : string[];
}



export class SPTargetedSitesCommandBar extends React.Component<ISPTargetedSitesCommandBarProps, ISPTargetedSitesCommandBarState> {
    constructor(props: ISPTargetedSitesCommandBarProps) {
        super(props);
        this.state = {
            showDeletionDialog: false,
            showEditPanel: false,
            firstEmail: "",
            emailRecipients: null,
            showEmail1Modal: false
        };
    }

    public render(): JSX.Element {

        const customButton = (props: IButtonProps) => {
            return (
                <CommandBarButton
                    {...props}
                    styles={{
                        ...props.styles,
                        textContainer: { fontSize: 18 },
                        icon: { color: '#E20000' }
                    }}
                />
            );
        };

        return (
            <div>
                <CommandBar
                    overflowButtonProps={{
                        ariaLabel: 'More commands',
                        menuProps: {
                            items: [], // Items must be passed for typesafety, but commandBar will determine items rendered in overflow
                            isBeakVisible: true,
                            beakWidth: 20,
                            gapSpace: 10,
                            directionalHint: DirectionalHint.topCenter
                        }
                    }}
                    buttonAs={customButton}
                    overflowItems={this.getOverflowItems()}
                    items={this.getItems()}
                    farItems={this.getFarItems()}
                    ariaLabel={'Use left and right arrow keys to navigate between commands'}
                />

                <SiteDeletionDialog
                    callback={this.props.callerback}
                    spHttpClient={this.props.spHttpClient}
                    siteSelected={this.props.selectedSite}
                    showDialog={this.state.showDeletionDialog}
                />
                <SiteEditPanel
                    refreshCallback={this.props.callerback}
                    spHttpClient={this.props.spHttpClient}
                    siteSelected={this.props.selectedSite}
                    showPanel={this.state.showEditPanel} />

                <Modal
                    isOpen={this.state.showEmail1Modal}
                    onDismiss={this._closeEmail1Modal}
                    containerClassName={styles.container}
                >
                    <div className={styles.manageLists}>
                        <div className={styles.modal}>


                            <div className={styles.container}>
                                <div className={styles.header}> First Email Template</div>
                                <div className={styles.body}>
                                    
                                    
                                    <p>
                                        {this.state.firstEmail}
                                    </p>
                                </div>

                            </div>

                        </div>
                    </div>

                </Modal>
            </div>
        );
    }

    
    // Data for CommandBar
    private getItems = () => {
        return [
            {
                key: 'deleteItem',
                name: 'Delete',
                cacheKey: 'myCacheKey', // changing this key will invalidate this items cache
                iconProps: {
                    iconName: 'Delete'
                },
                ariaLabel: 'Delete',
                onClick: () => this.setState({
                    showDeletionDialog: true
                })
            },
            {
                key: 'edit',
                name: 'Edit',
                iconProps: {
                    iconName: 'Edit'
                },
                onClick: () => this.setState({
                    showEditPanel: true
                })
            }

        ];
    }

    private getOverflowItems = () => {
        return [
            {
                key: 'visitSiteList',
                name: 'Site List',
                iconProps: {
                    iconName: 'Emoji'
                },
                href: 'https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/Lists/TargetSitesForDeletionSteven1234/AllItems.aspx',
                ['data-automation-id']: 'visitSiteList',
                onClick: () => console.log('hey')
            },
            {
                key: 'visitEmailList',
                name: 'Email List',
                iconProps: {
                    iconName: 'Emoji2'
                },
                href: 'https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/Lists/TargetedEmailTemplates/AllItems.aspx',
                ['data-automation-id']: 'visitEmailList'
            }
        ];
    }

    private getFarItems = () => {
        return [
            {
                key: 'mail',
                name: 'Generate Mail',
                ariaLabel: 'Generate Mail',
                subMenuProps: {
                    items: [
                        {
                            key: 'email1',
                            name: 'email 1',
                            iconProps: {
                                iconName: 'Mail'
                            },
                            onClick: () => {
                                if (this.props.siteIsSelected) {
                                    this._getFirstEmail();
                                    this._parseGroups();
                                    this._showEmail1Modal();
                                }



                            },
                            ['data-automation-id']: 'newEmailButton1'
                        },
                        {
                            key: 'email2',
                            name: 'email 2',
                            iconProps: {
                                iconName: 'Mail'
                            },
                            onClick: () => console.log('email2'),
                            ['data-automation-id']: 'newEmailButton2'
                        }
                    ]
                },
                iconProps: {
                    iconName: 'Mail'
                }
            }
        ];
    }
    public componentDidUpdate(previousProps: ISPTargetedSitesCommandBarProps, previousState: ISPTargetedSitesCommandBarState) {
        if (!previousState.showDeletionDialog && this.state.showDeletionDialog) {
            this.setState({
                showDeletionDialog: false
            });
        }
        if (!previousState.showEditPanel && this.state.showEditPanel) {
            this.setState({
                showEditPanel: false
            });
        }

    }

    private _generateEmail1 = (): void => {
        let tempstring: string[] = this.state.firstEmail.split("<#>");
        let resultString: string = "";
        for( var i = 0; i < tempstring.length; i++){
            if( i == tempstring.length -1){
                resultString += tempstring[i];
            }else{
                resultString += tempstring[i] + this.props.selectedSite.Title;
            }
        }
        this.setState({
            firstEmail: resultString
        });
    }
    private _parseGroups = (): void => {
        let tempString: string[] = this.props.selectedSite.emails.split("&quot;");
        let currentGroupName: string= "";
        let currentEmailNames: string[] = [];
        let grpNameMode: boolean = false;
        let usrEmailMode: boolean = false;
        let _emailGroups: emailGroups = {};
        tempString.forEach((s: string) => {
            if(s === "GroupName"){
                usrEmailMode = false;
                grpNameMode = true;
                if(currentGroupName != ""){
                    _emailGroups[currentGroupName] = currentEmailNames;
                }
                currentGroupName="";
            }
            else if(s === "userEmails"){
                usrEmailMode = true;
                grpNameMode = false;
                currentEmailNames = [];
            }else{
                if(grpNameMode){
                    if(s != "&#58;" && s != ","){
                        currentGroupName = s;
                    }
                }
                if(usrEmailMode){
                    if(s != "&#58;[" && s != "," && s != "]&#125;]</div>" && s != "]&#125;,&#123;"){
                        currentEmailNames.push(s);
                    }
                }
            }
        });
        
        console.log(_emailGroups);
    }

    private _getFirstEmail(): void {
        this.props.spHttpClient.get(`https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/_api/web/lists/getbytitle('TargetedEmailTemplates')/Items?$top=1000`, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                if (response.ok) {
                    response.json().then((responseJSON: ListResult) => {
                        console.log(responseJSON.value[0].bodyTemplate);
                        this.setState({
                            firstEmail: responseJSON.value[0].bodyTemplate
                        });
                        this._generateEmail1();
                    });
                }
            });


    }

    

    private _closeEmail1Modal = (): void => {
        this.setState({
            showEmail1Modal: false,
            emailRecipients: null,
            firstEmail: ""
        });
    }

    private _showEmail1Modal = (): void => {
        this.setState({
            showEmail1Modal: true
        });
    }


}
