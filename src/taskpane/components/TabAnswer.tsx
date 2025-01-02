import * as React from "react";
import { Button, Dialog, DialogSurface, Field, Spinner, Textarea, makeStyles } from "@fluentui/react-components";
import DialogForm from "./DialogForm";
import { DialogInfo, FieldInfo, sendRequest } from "../../../helpers";
import { getMailDetails, insertText } from "../taskpane";

/* global console, HTMLTextAreaElement, localStorage */

const useStyles = makeStyles({
  root: {
    flexDirection: "column",
    display: "flex",
    paddingLeft: "10px",
    paddingRight: "10px",
  },
  answer_dialog: {
    flexDirection: "column",
    display: "flex",
  },
  dialog_field: {
    marginBottom: "15px",
  },
  button_send: {
    width: "100%",
    marginBottom: "10px",
  },
  spinner: {
    width: "75%",
  },
});

const TabAnswer = () => {
  const styles = useStyles();

  const [showDialog, setShowDialog] = React.useState<DialogInfo>({ show: false, text: "" }); // dialog form
  const [showSpinner, setShowSpinner] = React.useState<boolean>(false); // spinner

  const [answerValue, setAnswerValue] = React.useState<FieldInfo>({ current: "", state: "none" }); // answer field value

  const def_answer =
    "Bitte antworte höflich dass ich morgen leider keine Zeit habe, aber übermorgen gerne zwischen 14-16 Uhr kann.";

  React.useEffect(() => {
    const getStartData = async () => {
      var answer = localStorage.getItem("answer"); // load last answer from storage

      // set default value
      if (!answer || answer == "") {
        answer = def_answer;
      }

      setAnswerValue({ current: answer, state: "none" });
    };

    getStartData(); // get Start Data
  }, []);

  // validate answer field
  const ValidateField = (): boolean => {
    let valid = true;
    var answer = answerValue.current ?? "";

    // set default value
    if (answer == " ") {
      answer = def_answer;
      localStorage.setItem("answer", answer); // save
    }

    answer = answer.trim();

    if (answer == "") {
      setAnswerValue({ current: answer, state: "error" });
      valid = false;
    } else {
      setAnswerValue({ current: answer, state: "none" });
    }

    return valid;
  };

  // change answer value
  const handleChangeValue = (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setAnswerValue({ ...answerValue, current: event.target.value });
  };

  // button - get full answer
  const onButtonSaveClick = () => {
    if (!ValidateField()) {
      return;
    }

    localStorage.setItem("answer", answerValue.current); // save

    var apiKey: string = localStorage.getItem("apiKey"); // load apiKey from storage

    if (apiKey == "") {
      setShowDialog({ show: true, text: "API-Schlüssel nicht angegeben" });
      return;
    }

    // get Mail Details
    getMailDetails((data) => {
      setShowSpinner(true);

      var user_input: string = answerValue.current;

      console.log("subject = " + data.subject);
      console.log("sender = " + data.sender);
      console.log("body = " + data.body);
      console.log("user_input = " + user_input);
      console.log("user_email = " + data.user_email);

      // send Request
      sendRequest("/outlook/generate-email", apiKey, data.subject, data.sender, data.body, user_input, data.user_email)
        .then(async (response) => {
          setShowSpinner(false);

          console.log("response = " + JSON.stringify(response, null, 2));

          // show error dialog
          if (response.status != "success") {
            setShowDialog({ show: true, text: response.detail });
            return;
          }

          insertText(response.email_text); // Write text to the cursor point in the compose surface.
        })
        .catch((error) => {
          setShowSpinner(false);
          setShowDialog({ show: true, text: `Request error: ${error}` }); // show error dialog
        });
    });
  };

  // result dialog event
  const handleResultDialog = () => {
    setShowSpinner(false);
    setShowDialog({ ...showDialog, show: false });
  };

  return (
    <div className={styles.root} role="tabpanel" aria-labelledby="Settings">
      {showSpinner && (
        <Dialog defaultOpen={true}>
          <DialogSurface className={styles.spinner}>
            <Spinner autoFocus labelPosition="after" label="Loading..." />
          </DialogSurface>
        </Dialog>
      )}

      {showDialog.show && (
        <DialogForm
          type="alert"
          title="Fehler"
          content={showDialog.text}
          closeName="OK"
          showPrimary={false}
          onResultDialog={handleResultDialog}
        />
      )}

      <div className={styles.answer_dialog}>
        <Field
          className={styles.dialog_field}
          label="Kurze Antwort"
          validationState={answerValue.state}
          validationMessage="Schreiben Sie eine kurze Antwort"
          required
        >
          <Textarea
            value={answerValue.current}
            rows={10}
            placeholder="type here..."
            size="large"
            resize="vertical"
            onChange={handleChangeValue}
          />
        </Field>

        <Button className={styles.button_send} appearance="primary" onClick={onButtonSaveClick}>
          Erhalten vollständige Antwort
        </Button>
      </div>
    </div>
  );
};

export default TabAnswer;
