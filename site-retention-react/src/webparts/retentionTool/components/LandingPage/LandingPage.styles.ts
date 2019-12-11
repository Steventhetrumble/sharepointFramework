import {
  IStackStyles,
  IStackItemStyles,
  IStackTokens,
  IIconStyles,
  IButtonStyles
} from "office-ui-fabric-react";
import { DefaultPalette } from "@uifabric/styling";

// Styles definition
export const landingPageStackStyles: IStackStyles = {
  root: {
    background: DefaultPalette.themeTertiary,
    minHeight: "40vh"
  }
};
export const landingPageStackItemStyles: IStackItemStyles = {
  root: {
    height: "20vh",
    overflow: "visible"
  }
};
export const landingPageCarousel: IStackItemStyles = {
  root: {
    height: "40vh",
    width: "40vw",
    position: "absolute",
    overflow: "hidden"
  }
};

// Example formatting
export const landingPageItemAlignmentsStackTokens: IStackTokens = {
  childrenGap: 12
};

export const iconStyles: IIconStyles = {
  root: {
    fontSize: "24px",
    height: "24px",
    width: "24px"
  }
};

export const compoundButtonStyle: IButtonStyles = {
  root: {
    backgroundColor: DefaultPalette.blackTranslucent40,
    color: DefaultPalette.magentaLight
  },
  description: {
    color: DefaultPalette.greenLight
  },
  rootHovered: {
    backgroundColor: DefaultPalette.blackTranslucent40,
    color: DefaultPalette.magentaLight
  },
  rootPressed: {
    backgroundColor: DefaultPalette.whiteTranslucent40,
    color: DefaultPalette.red
  },
  descriptionHovered: {
    color: DefaultPalette.greenLight,
    fontWeight: "bold"
  },
  descriptionPressed: {
    color: DefaultPalette.yellowLight,
    fontStretch: "ultra-expanded"
  },
  textContainer: {
    textAlign: "center"
  }
};
