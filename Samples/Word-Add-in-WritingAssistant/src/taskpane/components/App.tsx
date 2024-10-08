import * as React from "react";
import Header from "./Header";
import AnnotationComponents from "./Annotations";
import { Field, makeStyles } from "@fluentui/react-components";
import "bootstrap/dist/css/bootstrap.min.css";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "10vh",
  },
  welcome__header: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
  },
  field_label: {
    fontWeight: "bold",
    maxWidth: "80%",
  },
});

const App = (props: AppProps) => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} message="Writing Assistant Add-in" />
      <div className={styles.welcome__header}>
        <Field
          size="large"
          className={styles.field_label}
          label="The sample add-in showcases capabilities for error checking, rephrasing content and improving writing. "
        ></Field>
      </div>
      <AnnotationComponents />
    </div>
  );
};

export default App;
