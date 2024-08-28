/* global HTMLInputElement */
import * as React from "react";
import { useRef } from "react";
import { Button, FormControlLabel, Stack, Switch } from "@mui/material";
import { SetError } from "../hooks/useErrorSnackbar";
import useHighlightRow from "../utils/highlightRow/useHighlightRow";

export default function HighlightRow({ setError }: { setError: SetError }) {
  const { isEnabled, setIsEnabled, xColor, setXColor, keepSelection, setKeepSelection } = useHighlightRow(setError);

  const inputColorRef = useRef<HTMLInputElement | undefined>();

  return (
    <Stack direction="column" p={1}>
      <FormControlLabel
        sx={{ width: "fit-content", userSelect: "none" }}
        control={<Switch onChange={(e) => setIsEnabled(e.target.checked)} defaultChecked={isEnabled} />}
        label="Highlight Active Row"
      />
      <Button
        sx={{ width: "100%", height: "3rem", p: 0.5 }}
        onClick={() => inputColorRef.current?.click()}
        disabled={!isEnabled}
      >
        <input
          style={{ width: "100%", height: "100%", cursor: "pointer", opacity: isEnabled ? 1 : 0.25 }}
          type="color"
          value={xColor}
          onChange={(e: React.FormEvent<any>) => {
            e.currentTarget.value && setXColor(e.currentTarget.value);
          }}
          ref={inputColorRef}
          disabled={!isEnabled}
        ></input>
      </Button>
      {/* <FormControlLabel
        sx={{ width: "fit-content", userSelect: "none" }}
        control={
          <Switch
            onChange={(e) => setKeepSelection(e.target.checked)}
            color="secondary"
            defaultChecked={keepSelection}
          />
        }
        label="Keep Selection"
      /> */}
    </Stack>
  );
}
