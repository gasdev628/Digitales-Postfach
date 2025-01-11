import * as React from "react";
import { Button, Field, Input, makeStyles } from "@fluentui/react-components";
import DialogForm from "./DialogForm";
import { DialogInfo, FieldInfo } from "../../../helpers";

/* global HTMLInputElement, localStorage */

const useStyles = makeStyles({
  root: {
    flexDirection: "column",
    display: "flex",
    paddingLeft: "10px",
    paddingRight: "10px",
  },
  saving_dialog: {
    flexDirection: "column",
    display: "flex",
  },
  dialog_field: {
    fontWeight: 600,
    marginBottom: "15px",
  },
  dialog_add: {
    width: "100%",
    marginBottom: "10px",
  },
});

const TabSettings = () => {
  const styles = useStyles();

  const [showDialog, setShowDialog] = React.useState<DialogInfo>({ show: false, text: "" }); // dialog form
  const [keyValue, setKeyValue] = React.useState<FieldInfo>({ current: "", state: "none" }); // apiKey field value

  React.useEffect(() => {
    const getStartData = async () => {
      var apiKey = localStorage.getItem("apiKey"); // load apiKey from storage
      setKeyValue({ current: apiKey, state: "none" });
    };

    getStartData(); // get Start Data
  }, []);

  // validate field
  const ValidateField = (): boolean => {
    let valid = true;
    var apiKey = keyValue.current.trim();

    if (!/^[a-zA-Z0-9]{90,100}$/gm.test(apiKey) && apiKey != "") {
      setKeyValue({ ...keyValue, state: "error" });
      valid = false;
    } else {
      setKeyValue({ ...keyValue, state: "none" });
    }

    return valid;
  };

  // change input value
  const handleChangeValue = (event: React.ChangeEvent<HTMLInputElement>) => {
    setKeyValue({ ...keyValue, current: event.target.value });
  };

  // button - save api key
  const onButtonSaveClick = () => {
    if (!ValidateField()) {
      return;
    }

    localStorage.setItem("apiKey", keyValue.current); // save

    setShowDialog({ show: true, text: "Änderungen gespeichert" });
  };

  // result dialog event
  const handleResultDialog = () => {
    setShowDialog({ ...showDialog, show: false });
  };

  return (
    <div className={styles.root} role="tabpanel" aria-labelledby="Settings">
      {showDialog.show && (
        <DialogForm
          type="modal"
          title="Info"
          content={showDialog.text}
          closeName="OK"
          showPrimary={false}
          onResultDialog={handleResultDialog}
        />
      )}

      <div className={styles.saving_dialog}>
        <Field
          className={styles.dialog_field}
          label="API-Schlüsselwert"
          validationState={keyValue.state}
          validationMessage="Kopieren Sie Ihren API-Schlüsselwert und fügen Sie ihn ein"
          required
        >
          <Input value={keyValue.current} placeholder="z.b. outlookapi47eda17f38081e..." onChange={handleChangeValue} />
        </Field>

        <Button className={styles.dialog_add} appearance="primary" onClick={onButtonSaveClick}>
          Speichern
        </Button>
      </div>
    </div>
  );
};

export default TabSettings;
