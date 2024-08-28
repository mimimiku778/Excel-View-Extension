import * as React from "react";
import Snackbar, { SnackbarCloseReason } from "@mui/material/Snackbar";
import Alert, { AlertColor, AlertPropsColorOverrides } from "@mui/material/Alert";
import { OverridableStringUnion } from "@mui/types";
import { SetError } from "../hooks/useErrorSnackbar";

type Props = {
  messages: string[];
  setMessage: SetError;
  severity?: OverridableStringUnion<AlertColor, AlertPropsColorOverrides>;
  autoHideDuration?: number;
  clickAway?: boolean;
};

export default function SnackbarUi({ messages, setMessage, severity, autoHideDuration = 6000, clickAway }: Props) {
  const handleClose = (_: React.SyntheticEvent | Event, reason?: SnackbarCloseReason) => {
    if (!clickAway && reason === "clickaway") return;
    setMessage([]);
  };

  return (
    <Snackbar
      open={!!messages.length}
      onClose={handleClose}
      autoHideDuration={autoHideDuration}
    >
      <Alert onClose={handleClose} severity={severity} variant="filled" sx={{ width: "100%" }}>
        {messages.map((message, i) => (
          <div key={i}>{message}</div>
        ))}
      </Alert>
    </Snackbar>
  );
}
