import * as React from "react";
import UndoIcon from "@mui/icons-material/Undo";
import List from "@mui/material/List";
import ListItemButton from "@mui/material/ListItemButton";
import ListItemIcon from "@mui/material/ListItemIcon";
import ListItemText from "@mui/material/ListItemText";
import ViewColumnIcon from "@mui/icons-material/ViewColumn";
import DeleteIcon from "@mui/icons-material/Delete";
import { Box, Button, IconButton, ListItem, Stack, SxProps, Typography } from "@mui/material";
import AddCircleOutlineIcon from "@mui/icons-material/AddCircleOutline";
import { Theme } from "@emotion/react";
import { SetError } from "../hooks/useErrorSnackbar";
import useColumnProfile from "../utils/columnProfile/useColumnProfile";

const typegraphySx: SxProps<Theme> = {
  userSelect: "none",
  cursor: "default",
};

export default function ColumnProfileList({ setError }: { setError: SetError }) {
  const { add, restore, remove, profilesCount, undo, hasPrevious } = useColumnProfile(setError);

  return (
    <Box p={1}>
      <Typography variant="subtitle1" sx={typegraphySx}>
        Column Width Profile
      </Typography>
      <div>
        <Button variant="contained" startIcon={<AddCircleOutlineIcon />} fullWidth onClick={() => add()}>
          Add
        </Button>
        <Typography variant="caption" sx={typegraphySx}>
          Add new profile from columns in current worksheet.
        </Typography>
      </div>
      <Stack direction="column" spacing={1}>
        <List component="nav" aria-label="Column width profiles">
          {[...Array(profilesCount).keys()].map((index) => {
            return (
              <ListItem
                secondaryAction={
                  <IconButton edge="end" aria-label="delete" onClick={() => remove(index)}>
                    <DeleteIcon />
                  </IconButton>
                }
                disablePadding
              >
                <ListItemButton onClick={() => restore(index)}>
                  <ListItemIcon>
                    <ViewColumnIcon />
                  </ListItemIcon>
                  <ListItemText primary={`Profile ${index + 1}`} />
                </ListItemButton>
              </ListItem>
            );
          })}
        </List>
        <Stack direction="row" spacing={2}>
          {hasPrevious && (
            <Button
              onClick={() => undo()}
              size="small"
              variant="outlined"
              startIcon={<UndoIcon />}
              sx={{ width: "fit-content" }}
            >
              Undo Columns
            </Button>
          )}
        </Stack>
      </Stack>
    </Box>
  );
}
