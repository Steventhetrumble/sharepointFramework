import {
  IStackItemStyles,
  IStackTokens,
  IStackStyles
} from "office-ui-fabric-react";
import { DefaultPalette, FontSizes, FontWeights } from "@uifabric/styling";

export const stackItemStyles: IStackItemStyles = {
  root: {
    background: DefaultPalette.themePrimary,
    color: DefaultPalette.white,
    padding: 5
  }
};
export const pivotStackItemStyles: IStackItemStyles = {
  root: {
    background: DefaultPalette.white,
    padding: 10,
    overflow: "hidden"
  }
};

export const itemAlignmentsStackTokens: IStackTokens = {
  childrenGap: 5
};

export const stackStyles: IStackStyles = {
  root: {
    background: DefaultPalette.themeSecondary,
    paddingLeft: 10,
    paddingRight: 10,
    paddingBottom: 10
  }
};

export const container: IStackStyles = {
  root: {
    maxWidth: "70vw",
    width: "700px",
    margin: "0px auto",
    maxHeight: "70vh",
    boxShadow:
      "0 2px 4px 0 rgba(0, 0, 0, 0.2), 0 25px 50px 0 rgba(0, 0, 0, 0.1)"
  }
};

export const row: IStackStyles = {
  root: {
    backgroundColor: DefaultPalette.whiteTranslucent40,
    marginTop: "5px",
    marginBottom: "5px"
  }
};

export const title: IStackItemStyles = {
  root: {
    fontSize: FontSizes.xLarge,
    color: DefaultPalette.white,
    backgroundColor: DefaultPalette.themePrimary
  }
};

export const subTitle: IStackItemStyles = {
  root: {
    fontSize: FontSizes.large,
    color: DefaultPalette.whiteTranslucent40
  }
};
