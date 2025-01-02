import * as React from "react";
import { makeStyles, Tab, TabList } from "@fluentui/react-components";
import type { SelectTabData, SelectTabEvent, TabValue } from "@fluentui/react-components";
import TabAnswer from "./TabAnswer";
import TabSettings from "./TabSettings";

/* global localStorage */

const useStyles = makeStyles({
  root: {
    alignItems: "flex-start",
    display: "flex",
    flexDirection: "column",
    justifyContent: "flex-start",
    paddingTop: "10px",
    paddingLeft: "1px",
    paddingRight: "1px",
  },
  tab_menu: {
    paddingBottom: "20px",
  },
  tab_cont: {
    width: "100%",
  },
});

export const TabPanes = () => {
  const styles = useStyles();

  var apiKey = localStorage.getItem("apiKey"); // load apiKey from storage

  const [selectedTab, setSelectedTab] = React.useState<TabValue>(apiKey == "" ? "settings" : "answer");

  const onTabSelect = (_event: SelectTabEvent, data: SelectTabData) => {
    setSelectedTab(data.value);
  };

  return (
    <div className={styles.root}>
      <TabList className={styles.tab_menu} selectedValue={selectedTab} onTabSelect={onTabSelect}>
        <Tab id="Answer" value="answer">
          Antwort
        </Tab>

        <Tab id="Settings" value="settings">
          Einstellungen
        </Tab>
      </TabList>

      <div className={styles.tab_cont}>
        {selectedTab === "answer" && <TabAnswer />}
        {selectedTab === "settings" && <TabSettings />}
      </div>
    </div>
  );
};

export default TabPanes;
