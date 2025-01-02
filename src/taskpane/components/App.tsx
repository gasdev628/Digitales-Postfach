import * as React from "react";
import { Divider, makeStyles } from "@fluentui/react-components";
import { TabPanes } from "./TabPanes";

/* global */

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    maxWidth: "370px",
  },
  divider: {
    marginTop: "10px",
  },
});

const App = () => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <TabPanes />
      <Divider className={styles.divider} inset />
    </div>
  );
};

export default App;
