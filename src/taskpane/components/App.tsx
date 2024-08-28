import * as React from "react";
import { Divider, FormGroup, Stack } from "@mui/material";
import useErrorSnackbar from "../hooks/useErrorSnackbar";
import SnackbarUi from "./SnackBarUi";
import HighlightRow from "./HighlightRow";
import ColumnProfileList from "./ColumnProfileList";

interface AppProps {
  title: string;
}

export default function App({}: AppProps) {
  const [errors, setError] = useErrorSnackbar();

  return (
    <>
      <FormGroup>
        <Stack direction="column" spacing={2} p={1} divider={<Divider orientation="horizontal" flexItem />}>
          <HighlightRow setError={setError} />
          <ColumnProfileList setError={setError} />
        </Stack>
      </FormGroup>
      <SnackbarUi messages={errors} setMessage={setError} severity={"warning"} autoHideDuration={6000} clickAway />
    </>
  );
}
