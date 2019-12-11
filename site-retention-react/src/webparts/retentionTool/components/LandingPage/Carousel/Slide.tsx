import * as React from "react";
import { Stack, IStackStyles } from "office-ui-fabric-react";

export const Slide = ({ image }) => {
  const styles: IStackStyles = {
    root: {
      height: "41vh",
      width: "650px",
      minWidth: "50vw",
      minHeight: "40vh",
      backgroundImage: `URL(${image})`,
      backgroundSize: "cover",
      backgroundRepeat: "no-repeat",
      backgroundPosition: "50% 60%"
    }
  };
  return (
    <Stack className="slide" styles={styles} horizontalAlign="stretch"></Stack>
  );
};
