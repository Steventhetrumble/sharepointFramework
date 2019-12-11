import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { lorem } from '@uifabric/example-data';
import { Stack, IStackProps } from 'office-ui-fabric-react/lib/Stack';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration, ISPHttpClientOptions } from '@microsoft/sp-http';
import { string } from 'prop-types';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';


export interface ISecondRoundEmailState {
  multiline: boolean;
  email: string;
}

export interface ISecondRoundEmailProps {
  spHttpClient: SPHttpClient;
}
export interface ListResult {
  value: ListRow[];

}

export interface ListRow {
  Title: string;
  bodyTemplate: string;

}

export class SecondRoundEmail extends React.Component<ISecondRoundEmailProps, ISecondRoundEmailState> {
  constructor(props: ISecondRoundEmailProps) {
    super(props);
    this.state = {
      multiline: true,
      email: ""
    };
    this._getSecondEmail();
  }

  public render(): JSX.Element {

    // TextFields don't have to be inside Stacks, we're just using Stacks for layout
    const columnProps: Partial<IStackProps> = {
      tokens: { childrenGap: 15 },
      styles: { root: { width: '70vh', height: '80vh' } }
    };

    return (
      <div>

        <Stack horizontal tokens={{ childrenGap: 50 }} styles={{ root: { width: '90vh', height: '60vh' } }}>

          <Stack {...columnProps}>
            <Stack.Item>
              <TextField label="Email Template" multiline defaultValue={this.state.email} autoAdjustHeight onChange={this._onChange} />

            </Stack.Item>
            <Stack.Item align="end">
              <PrimaryButton styles={{ root: { maxWidth: 60 } }} onClick={this._saveEmail}>Save </PrimaryButton>

            </Stack.Item>
          </Stack>
        </Stack>

      </div>

    );
  }

  private _onChange = (ev: any, newText: string): void => {
    const newEmail = newText;
    if (newEmail !== this.state.email) {
      this.setState({ email: newEmail });
    }
  }

  private _getSecondEmail(): void {
    this.props.spHttpClient.get(`https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/_api/web/lists/getbytitle('TargetedEmailTemplates')/Items?$top=1000`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          response.json().then((responseJSON: ListResult) => {
            this.setState({
              email: responseJSON.value[1].bodyTemplate
            });
          });
        }
      });
  }

  private _saveEmail = (): void => {
    var itemType = getItemTypeForListName('TargetedEmailTemplates');
    console.log(this.state.multiline);

    const body: string = JSON.stringify({
      '__metadata': {
        'type': itemType
      },
      bodyTemplate: this.state.email

    });
    this.props.spHttpClient.post(`https://rcirogers.sharepoint.com/teams/ITS-Sharepoint/_api/web/Lists(guid'01cb56d5-817b-4358-a1c8-b328342534e6')/Items(2)`, SPHttpClient.configurations.v1, {
      headers: {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
        'Accept': 'application/json;odata=verbose',
        'Content-type': 'application/json;odata=verbose',
        'odata-version': ''
      },
      body: body
    })
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          console.log("Save response:", response.status);
        }
      });
  }
}

// Get List Item Type metadata
function getItemTypeForListName(name) {
  return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
}