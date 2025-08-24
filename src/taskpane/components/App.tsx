import * as React from "react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import { insertText } from "../taskpane";
import { callGraphMe, initializeMsal } from "../../msal";
import { Button } from "@fluentui/react-components";
import { useState } from "react";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
  graphContainer: {
  padding: "12px",
  },
  graphOutput: {
  marginTop: "12px",
  whiteSpace: "pre-wrap",
  fontSize: "12px",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  const listItems: HeroListItem[] = [
    {
      icon: <Ribbon24Regular />,
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Unlock features and functionality",
    },
    {
      icon: <DesignIdeas24Regular />,
      primaryText: "Create and visualize like a pro",
    },
  ];

  const [graphResult, setGraphResult] = useState<any | null>(null);
  const [loadingGraph, setLoadingGraph] = useState(false);

  async function onTestGraph() {
    setLoadingGraph(true);
    setGraphResult(null);
    try {
      await initializeMsal();
      const data = await callGraphMe();
      setGraphResult(data);
    } catch (err: any) {
      setGraphResult({ error: err?.message ?? String(err) });
    } finally {
      setLoadingGraph(false);
    }
  }

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} message="Welcome" />
      <HeroList message="Discover what this add-in can do for you today!" items={listItems} />
      <TextInsertion insertText={insertText} />
      <div className={styles.graphContainer}>
        <Button appearance="primary" onClick={onTestGraph} disabled={loadingGraph}>
          {loadingGraph ? "Testing…" : "Test Graph (/me)"}
        </Button>
        <div className={styles.graphOutput}>
          {graphResult ? JSON.stringify(graphResult, null, 2) : "No result yet."}
        </div>
      </div>
    </div>
  );
};

export default App;
