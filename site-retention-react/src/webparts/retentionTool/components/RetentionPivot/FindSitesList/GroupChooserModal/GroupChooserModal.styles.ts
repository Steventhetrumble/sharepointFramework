import { IStackSlots, IStackStyles, IStackItemStyles, IStackTokens } from "office-ui-fabric-react";

export const modalContainer: IStackStyles = {
  root: {
    minWidth: "80vw",

    minHeight: "90vh",
    margin: "0px auto",
    boxShadow:
      "0 2px 4px 0 rgba(0, 0, 0, 0.2), 0 25px 50px 0 rgba(0, 0, 0, 0.1)"
  }
};
export const modalBody: IStackStyles = {
  root: {
    minWidth: "40vw",

    minHeight: "80vh",
  }
};

export const modalComponents: IStackItemStyles = {
    root: {
        width: "30vw"
    }
};

export const modalStackTokens: IStackTokens = {
    childrenGap: 12
};
