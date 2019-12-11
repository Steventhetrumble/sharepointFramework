import * as React from 'react';
import { Dropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { SPGroupList, SPGroup, ISPUsers, ISPUser, ISPSite, ISPGroupEmails, ISPGroupEmail } from '../../../../common/IObjects';
import styles from '../../../ManageLists.module.scss';
import { SPScrollableUserList } from '../SPScrollableUserList/SPScrollableUserList';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration, ISPHttpClientOptions } from '@microsoft/sp-http';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { getId } from '@uifabric/utilities';
import Modal from 'office-ui-fabric-react/lib/Modal';
import { Result } from '../../../IManageLists';
import { TargetedSites } from '../../SPDetailList';
import { Label } from 'office-ui-fabric-react';


export interface ISPModalState {
    selectedItems: string[];
    DisplayOptions: IDropdownOption[];
    updated: boolean;
    selectedGroupIds: number[];
    groupsFound: boolean;
    allGroups: { [id: number]: ISPUsers };
    selectedGroupEmails: ISPGroupEmails;
    _showModal: boolean;
    inputError: string;
    title: string;

}

export interface ISPModalProps {
    groups: SPGroupList;
    spHttpClient: SPHttpClient;
    site: ISPSite;
    showModal: boolean;
    sitesAlreadyTargeted: TargetedSites;
    siteAddedCallback(s: string): void; 

}

const warningStyle = {
    root: {
      color: 'red'
    }
  };

export class SPModal extends React.Component<ISPModalProps, ISPModalState> {
    private _selectedEmails: ISPGroupEmails;
    private _titleId: string = getId('title');
    private _subtitleId: string = getId('subText');

    constructor(props: ISPModalProps) {
        super(props);
        this._selectedEmails = { groups: [] };
        this.state = {
            selectedItems: [],
            DisplayOptions: [
                { key: 'fruitsHeader', text: 'Fruits', itemType: DropdownMenuItemType.Header },
                { key: 'apple', text: 'Apple' },
                { key: 'banana', text: 'Banana' },
                { key: 'orange', text: 'Orange', disabled: true },
                { key: 'grape', text: 'Grape' },
                { key: 'divider_1', text: '-', itemType: DropdownMenuItemType.Divider },
                { key: 'vegetablesHeader', text: 'Vegetables', itemType: DropdownMenuItemType.Header },
                { key: 'broccoli', text: 'Broccoli' },
                { key: 'carrot', text: 'Carrot' },
                { key: 'lettuce', text: 'Lettuce' }
            ],
            updated: false,
            selectedGroupIds: [],
            groupsFound: false,
            allGroups: {},
            selectedGroupEmails: { groups: [] },
            _showModal: false,
            inputError: '',
            title: ''
        };
    }

    public render() {
        const { selectedItems } = this.state;

        return (
            <Modal
                titleAriaId={this._titleId}
                subtitleAriaId={this._subtitleId}
                isOpen={this.state._showModal}
                onDismiss={this._closeModal}
                isBlocking={false}
                containerClassName={styles.manageLists}
            >
                <div className={styles.modal}>
                    <div className={styles.container}>

                        <div className={styles.titleBlock}>
                            <div className={styles.row}>
                                <span id={this._titleId} className={styles.title}>{this.state.title}</span>
                            </div>
                        </div>
                        

                        <div className={styles.row}>
                        <Label styles={warningStyle}>{(this.state._showModal && this.props.sitesAlreadyTargeted[this.props.site.Title])? `*${this.props.site.Title} is already targeted for deletion and can not be added`: ''}</Label>
                            <div className={styles.description}> Choose the most appropriate Owner Group</div>
                            <div className={styles.column}>
                                <h1 className={styles.description}> Users Emails</h1>
                                <SPScrollableUserList emailGroups={this.state.selectedGroupEmails} />
                            </div>
                            <div className={styles.column}>
                                <Dropdown
                                    placeholder="Select Group"
                                    label="Group Owners"
                                    selectedKeys={selectedItems}
                                    onChange={this._onChange}
                                    multiSelect
                                    options={this.state.DisplayOptions}
                                    styles={{ dropdown: { width: 300 } }}
                                    errorMessage={this.state.inputError}
                                    required
                                />
                            </div>
                        </div>
                        <div className={styles.footer}>
                            <div id={this._subtitleId} className={styles.row}>
                                <PrimaryButton onClick={this._addTargetedSite} style={{ marginRight: '8px' }}>Confirm</PrimaryButton>
                                <DefaultButton onClick={this._closeModal} text="Close" />
                            </div>
                        </div>
                    </div>
                </div>
            </Modal>
        );
    }

    public componentDidUpdate(previousProps: any, previousState: ISPModalState) {
        if (this.props.showModal && !this.state._showModal) {
            this._showModal();
            this.setState({
                title: this.props.site.Title
            });
        }
        if (!this.state.updated && this.props.groups != null) {
            this._makeOptions();
        }
        if (previousState.selectedGroupIds !== this.state.selectedGroupIds) {
            if (this.state.groupsFound) {
                let _groupEmails: ISPGroupEmails = { groups: [] };
                this.state.selectedGroupIds.forEach((id: number) => {
                    let _group: ISPGroupEmail = { GroupName: "", userEmails: [] };
                    if (this.state.allGroups[id] == null) {
                        console.log(this.state.allGroups);
                    } else {

                        _group.GroupName = this.state.allGroups[id].OwnerTitle;//this.state.allGroups[id].OwnerTitle
                        this.state.allGroups[id].value.forEach((user: ISPUser) => {
                            _group.userEmails.push(user.Email);
                        });
                        _groupEmails.groups.push(_group);
                    }
                });
                this._selectedEmails = _groupEmails;
                this.setState({
                    selectedGroupEmails: this._selectedEmails
                });
            }
        }
    }

    private _makeOptions(): void {
        let _options: IDropdownOption[] = [];
        let _all_users: { [id: number]: ISPUsers } = {};

        _options.push({ key: 'SuggestedOptionsHeader', text: 'SuggestedGroup', itemType: DropdownMenuItemType.Header });
        let bestOption: IDropdownOption = { key: 'NoOptions', text: 'No Suggestions', disabled: true };

        this.props.groups.value.forEach((group: SPGroup) => {
            if (group.Id == 5) {
                bestOption = { key: group.Id, text: `ID:${group.Id} ${group.OwnerTitle}` };
            } else if (group.Id == 6 && bestOption.key != 5) {
                bestOption = { key: group.Id, text: `ID:${group.Id} ${group.OwnerTitle}` };
            }
        });

        _options.push(bestOption);
        _options.push({ key: 'groups', text: 'Group Options', itemType: DropdownMenuItemType.Header });

        this.props.groups.value.forEach((group: SPGroup) => {
            if (group.Id != bestOption.key) {
                if (group.OwnerTitle === "System Account" || group.OwnerTitle === "serv_epmwss") {
                    _options.push({ key: group.Id, text: `ID:${group.Id} ${group.OwnerTitle}`, disabled: true });
                }
                else {
                    this.props.spHttpClient.get(this.props.site.Url + `/_api/Web/SiteGroups/GetById(${group.Id})/Users`, SPHttpClient.configurations.v1)
                        .then((response: SPHttpClientResponse) => {
                            response.json().then((responseJSON: ISPUsers) => {
                                if (responseJSON.error == null) {
                                    let _users: ISPUsers = responseJSON;
                                    _all_users[group.Id] = _users;
                                    _all_users[group.Id].OwnerTitle = group.OwnerTitle;
                                    console.log(group.OwnerTitle);
                                    _options.push({ key: group.Id, text: `ID:${group.Id} ${group.OwnerTitle}` });
                                }
                                else {
                                    console.log(responseJSON.error);
                                    _options.push({ key: group.Id, text: `ID:${group.Id} ${group.OwnerTitle}`, disabled: true });
                                }
                            });
                        });
                }
            }
            else {
                this.props.spHttpClient.get(this.props.site.Url + `/_api/Web/SiteGroups/GetById(${group.Id})/Users`, SPHttpClient.configurations.v1)
                    .then((response: SPHttpClientResponse) => {
                        response.json().then((responseJSON: ISPUsers) => {
                            if (responseJSON.error == null) {
                                let _users: ISPUsers = responseJSON;
                                _all_users[group.Id] = _users;
                                _all_users[group.Id].OwnerTitle = group.OwnerTitle;
                            }
                        });
                    });
            }
        });

        this.setState({
            DisplayOptions: _options,
            updated: true,
            groupsFound: true,
            allGroups: _all_users
        });

    }

    private _onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        const newSelectedItems = [...this.state.selectedItems];
        const newSelectedIds = [...this.state.selectedGroupIds];
        if (item.selected) {
            // add the option if it's checked
            newSelectedItems.push(item.key as string);
            newSelectedIds.push(Number(item.key));
        } else {
            // remove the option if it's unchecked
            const currIndex = newSelectedItems.indexOf(item.key as string);
            const numIndex = newSelectedIds.indexOf(Number(item.key));

            if (numIndex > -1) {
                newSelectedIds.splice(numIndex, 1);
            }
            if (currIndex > -1) {
                newSelectedItems.splice(currIndex, 1);
            }
        }
        this.setState({
            selectedItems: newSelectedItems,
            selectedGroupIds: newSelectedIds
        });

    }
    private _showModal = (): void => {
        this.setState({
            _showModal: true
        });
    }

    private _closeModal = (): void => {
        this.setState({
            selectedItems: [],
            DisplayOptions: [
                { key: 'fruitsHeader', text: 'Fruits', itemType: DropdownMenuItemType.Header },
                { key: 'apple', text: 'Apple' },
                { key: 'banana', text: 'Banana' },
                { key: 'orange', text: 'Orange', disabled: true },
                { key: 'grape', text: 'Grape' },
                { key: 'divider_1', text: '-', itemType: DropdownMenuItemType.Divider },
                { key: 'vegetablesHeader', text: 'Vegetables', itemType: DropdownMenuItemType.Header },
                { key: 'broccoli', text: 'Broccoli' },
                { key: 'carrot', text: 'Carrot' },
                { key: 'lettuce', text: 'Lettuce' }
            ],
            updated: false,
            selectedGroupIds: [],
            groupsFound: false,
            allGroups: {},
            selectedGroupEmails: { groups: [] },
            _showModal: false
        });
    }

    private _checkItems(): boolean {
        if (this.state.selectedGroupEmails.groups.length > 0) {
            return true;
        } else {
            return false;
        }
    }

    private _addTargetedSite = (): void => {
        if (!this.props.sitesAlreadyTargeted[this.props.site.Title]) {
            console.log("I Pressed the add site button");
            if (this._checkItems()) {
                this._postListData();
                this._closeModal();
                this.props.siteAddedCallback(this.props.site.Title);
            } else {
                this.setState({
                    inputError: "This field is required."
                });
            }
        }

    }

    // and example of a post requrest... perhaps will need this at some point
    private _postListData(): void {
        console.log(this.props.site.LastItemUserModifiedDateSharepoint);
        var itemType = getItemTypeForListName('TargetSitesForDeletionSteven1234');
        const body: string = JSON.stringify({
            '__metadata': {
                'type': itemType
            },
            'Title': this.props.site.Title,
            'URL': this.props.site.Url,
            'UserLastModified': this.props.site.LastItemUserModifiedDateSharepoint,
            'RecentViews': this.props.site.ViewsRecent,
            'emails': JSON.stringify(this.state.selectedGroupEmails.groups)

        });
        // make the actuall post request and handle the resulsts
        this.props.spHttpClient.post(`https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/_api/web/lists/getbytitle('TargetSitesForDeletionSteven1234')/Items`, SPHttpClient.configurations.v1, {
            headers: {
                'Accept': 'application/json;odata=verbose',
                'Content-type': 'application/json;odata=verbose',
                'odata-version': ''
            },
            body: body
        })
            .then((response: SPHttpClientResponse) => {
                response.json().then((responseJSON: Result) => {

                    console.log(responseJSON);
                }).catch(err => {
                    console.log("error on post", err);
                });
            });
    }
}

// Get List Item Type metadata
function getItemTypeForListName(name) {
    return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
}