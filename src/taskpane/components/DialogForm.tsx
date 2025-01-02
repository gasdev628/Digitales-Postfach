import * as React from "react";
import Parser from "html-react-parser";
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogTrigger,
  makeStyles,
} from "@fluentui/react-components";

export interface DialogFormProps {
  type: "non-modal" | "modal" | "alert";
  title: string;
  content: string;
  closeName: string;
  showPrimary: boolean;
  onResultDialog: (result: "ok" | "cancel") => void;
}

const useStyles = makeStyles({
  dialog: {
    marginTop: "10px",
    marginRight: "10px",
    marginLeft: "10px",
  },
  dialog_apply: {
    backgroundColor: "red",
  },
});

const DialogForm = (props: DialogFormProps) => {
  const styles = useStyles();

  // button - confirm
  const onButtonApplyClick = () => {
    props.onResultDialog("ok");
  };

  // button - cancel
  const onOpenChange = () => {
    props.onResultDialog("cancel");
  };

  return (
    <Dialog modalType={props.type} defaultOpen={true} onOpenChange={onOpenChange}>
      <DialogSurface className={styles.dialog}>
        <DialogBody>
          <DialogTitle>{props.title}</DialogTitle>

          <DialogContent>{Parser(props.content)}</DialogContent>

          <DialogActions>
            <DialogTrigger disableButtonEnhancement>
              <Button appearance="secondary" autoFocus>
                {props.closeName}
              </Button>
            </DialogTrigger>
            {props.showPrimary && (
              <Button appearance="primary" className={styles.dialog_apply} onClick={onButtonApplyClick}>
                Löschen
              </Button>
            )}
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};

export default DialogForm;
