import React from "react";
import {
  ListItem,
  ListItemIcon,
  ListItemText,
  Paper,
  TextField,
  Tooltip,
  Typography,
} from "@mui/material";

import { DragDropContext, Droppable, Draggable } from "react-beautiful-dnd";
import DirectionsRunIcon from "@mui/icons-material/DirectionsRun";
import PlayArrowIcon from "@mui/icons-material/PlayArrow";
import StopIcon from "@mui/icons-material/Stop";
import CheckIcon from "@mui/icons-material/Check";
import DownloadIcon from "@mui/icons-material/Download";
import MenuIcon from "@mui/icons-material/Menu";
import ThumbUpIcon from "@mui/icons-material/ThumbUp";
import LockIcon from "@mui/icons-material/Lock";
import LockOpenIcon from "@mui/icons-material/LockOpen";

const ProcedureRow = (props) => {
  const { step, processClassPhase, index } = props;

  return (
    <Paper elevation={5} sx={{ m: 1 }}>
      <Tooltip title={"Test tooltip"}>
        <ListItem
          role="listitem"
          sx={{
            cursor: "pointer",
            height: "35px",
          }}
        >
          <ListItemIcon>
            <MenuIcon />
          </ListItemIcon>
          <Typography
            component="h5"
            variant="h5"
            color="inherit"
            noWrap
            sx={{ mr: 2, width: "50px" }}
          >
            {index + 1}
          </Typography>
          {/* <ListItemIcon>
            {step.stepType === "Allocate" ? (
              <LockIcon />
            ) : step.stepType === "Run" ? (
              <DirectionsRunIcon />
            ) : step.stepType === "Start" ? (
              <PlayArrowIcon />
            ) : step.stepType === "Stop" ? (
              <StopIcon />
            ) : step.stepType === "Deallocate" ? (
              <LockOpenIcon />
            ) : (
              <ThumbUpIcon />
            )}
          </ListItemIcon> */}
          <ListItemText id={step.ID} primary={processClassPhase.Name} />
        </ListItem>
      </Tooltip>
    </Paper>
  );
};

export default ProcedureRow;
