import * as React from "react";
import { CompoundButton, Stack, Separator, Icon } from "office-ui-fabric-react";
import {
  landingPageStackStyles,
  landingPageStackItemStyles,
  landingPageItemAlignmentsStackTokens,
  iconStyles,
  landingPageCarousel,
  compoundButtonStyle
} from "./LandingPage.styles";
import { ILandingPageProps, ILandingPageState } from "./LandingPage.types";
import { Carousel } from "./Carousel/Carousel";

export class LandingPage extends React.Component<
  ILandingPageProps,
  ILandingPageState
> {
  constructor(props: ILandingPageProps) {
    super(props);

    this.state = {};
  }
  public render(): JSX.Element {
    return (
      <Stack
        styles={landingPageStackStyles}
        tokens={landingPageItemAlignmentsStackTokens}
      >
        <Stack.Item styles={landingPageStackItemStyles}>
          <Stack horizontalAlign="center">
            <Carousel />
          </Stack>
        </Stack.Item>
        <Stack.Item align="center" styles={landingPageStackItemStyles}>
          <Stack styles={{ root: { paddingTop: "10vh" } }}>
            <CompoundButton
              styles={compoundButtonStyle}
              secondaryText="This will be a nice looking landing page at some point"
              disabled={this.props.disabled}
              checked={this.props.checked}
              onClick={this.props.onClick}
            >
              Begin use of app
            </CompoundButton>
          </Stack>
        </Stack.Item>
      </Stack>
    );
  }
}
