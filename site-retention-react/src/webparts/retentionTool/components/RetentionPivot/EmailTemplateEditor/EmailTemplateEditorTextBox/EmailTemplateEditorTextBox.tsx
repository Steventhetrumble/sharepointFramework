import * as React from "react";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { lorem } from "@uifabric/example-data";
import { Stack, IStackProps } from "office-ui-fabric-react/lib/Stack";
import { stackStyles } from "./EmailTemplateEditorTextBox.styles";

export interface IEmailTemplateEditorTextBoxState {
  multiline: boolean;
}

export interface IEmailTemplateEditorTextBoxProps {
  email: string;
}

export class EmailTemplateEditorTextBox extends React.Component<
  IEmailTemplateEditorTextBoxProps,
  IEmailTemplateEditorTextBoxState
> {
  public state: IEmailTemplateEditorTextBoxState = { multiline: false };
  private _lorem: string = lorem(100);

  public render(): JSX.Element {
    return (
      <Stack
        horizontalAlign="stretch"
        verticalAlign="center"
        tokens={{ childrenGap: 50 }}
        styles={stackStyles}
      >
        <TextField
          label="Standard"
          multiline
          defaultValue={this.props.email}
          autoAdjustHeight
        />
      </Stack>
    );
  }

  private _onChange = (ev: any, newText: string): void => {
    const newMultiline = newText.length > 50;
    if (newMultiline !== this.state.multiline) {
      this.setState({ multiline: newMultiline });
    }
  };
}
