import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Progress from "./Progress";
import { testScript } from "../scripts/testScript";

/* global require  */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const App: React.FC<AppProps> = ({ title, isOfficeInitialized }) => {
  return !isOfficeInitialized ? (
    <Progress
      title={title}
      logo={require("./../../../assets/logo-filled.png")}
      message="Please sideload your addin to see app body."
    />
  ) : (
    <div className="ms-welcome">
      <DefaultButton
        className="ms-welcome__action"
        iconProps={{ iconName: "ChevronRight" }}
        onClick={() => {
          testScript();
        }}
      >
        Run
      </DefaultButton>
      <DefaultButton
        className="ms-welcome__action"
        iconProps={{ iconName: "ChevronRight" }}
        onClick={() => {
          testScript();
        }}
      >
        Kita
      </DefaultButton>
    </div>
  );
};

export default App;
