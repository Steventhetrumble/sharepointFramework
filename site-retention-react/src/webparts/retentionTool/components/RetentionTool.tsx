import * as React from "react";
import styles from "./RetentionTool.module.scss";
import {
  IRetentionToolProps,
  IRetentionToolState
} from "./IRetentionTool.types";
import { LandingPage } from "./LandingPage/LandingPage";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  Stack,
  Text,
  IStackItemStyles,
  DefaultPalette,
  IStackStyles,
  IStackTokens
} from "office-ui-fabric-react";
import {
  stackItemStyles,
  stackStyles,
  itemAlignmentsStackTokens,
  pivotStackItemStyles,
  row,
  container,
  title,
  subTitle
} from "./RetentionTool.styles";
import { RetentionPivot } from "./RetentionPivot/RetentionPivot";
import { ISPSite } from "./common/IObjects";
import { SharepointRestProvider } from "./dataproviders/SharepointRestProvider";

export default class RetentionTool extends React.Component<
  IRetentionToolProps,
  IRetentionToolState
> {
  constructor(props: IRetentionToolProps) {
    super(props);
    this.state = {
      showPivot: false,
      sitesPromise: null,
      provider: new SharepointRestProvider(this.props.context)
    };

    this.state.provider.getRootSiteData().then((site: ISPSite) => {
      console.log(site);
    });
  }

  public render(): React.ReactElement<IRetentionToolProps> {
    if (!this.state.showPivot) {
      return (
        <Stack styles={container}>
          <Stack styles={stackStyles} tokens={itemAlignmentsStackTokens}>
            <Stack.Item align="center" styles={stackItemStyles}>
              <Text styles={title}>Site Retention Tool</Text>
            </Stack.Item>
            <Stack.Item styles={stackItemStyles}>
              <LandingPage onClick={this._continueForward}></LandingPage>
            </Stack.Item>
          </Stack>
        </Stack>
      );
    } else {
      return (
        <Stack styles={container}>
          <Stack styles={stackStyles} tokens={itemAlignmentsStackTokens}>
            <Stack.Item align="center" styles={stackItemStyles}>
              <Text styles={title}>Site Retention Tool</Text>
            </Stack.Item>
            <Stack.Item styles={pivotStackItemStyles} align="stretch">
              <RetentionPivot
                provider={this.state.provider}
                context={this.props.context}
              ></RetentionPivot>
            </Stack.Item>
          </Stack>
        </Stack>
      );
    }
  }

  private _continueForward = (): void => {
    this.setState({
      showPivot: true
    });
  };
}
