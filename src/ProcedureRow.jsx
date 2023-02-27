import React from 'react';
import {
  ListItemButton,
  ListItemIcon,
  ListItemText,
  TextField,
  Typography,
} from '@mui/material';

import { DragDropContext, Droppable, Draggable } from 'react-beautiful-dnd';
import DirectionsRunIcon from '@mui/icons-material/DirectionsRun';
import PlayArrowIcon from '@mui/icons-material/PlayArrow';
import StopIcon from '@mui/icons-material/Stop';
import CheckIcon from '@mui/icons-material/Check';
import DownloadIcon from '@mui/icons-material/Download';
import MenuIcon from '@mui/icons-material/Menu';
import ThumbUpIcon from '@mui/icons-material/ThumbUp';
import LockIcon from '@mui/icons-material/Lock';
import LockOpenIcon from '@mui/icons-material/LockOpen';

const ProcedureRow = props => {
  const { step, index } = props;

  return (
    <ListItemButton
      role='listitem'
      button
      sx={{
        cursor: 'pointer',
        border: '1px solid gray',
        mx: 3,
        height: '50px',
      }}
    >
      <ListItemIcon>
        <MenuIcon />
      </ListItemIcon>
      <Typography
        component='h5'
        variant='h5'
        color='inherit'
        noWrap
        sx={{ mr: 2, width: '50px' }}
      >
        {index + 1}
      </Typography>
      <ListItemIcon>
        {step.stepType === 'Allocate' ? (
          <LockIcon />
        ) : step.stepType === 'Run' ? (
          <DirectionsRunIcon />
        ) : step.stepType === 'Start' ? (
          <PlayArrowIcon />
        ) : step.stepType === 'Stop' ? (
          <StopIcon />
        ) : step.stepType === 'Deallocate' ? (
          <LockOpenIcon />
        ) : (
          <ThumbUpIcon />
        )}
      </ListItemIcon>
      <ListItemText id={index} primary={`${step.stepName}`} />
      <TextField size='small' label={'Phase Parameter'}  />
    </ListItemButton>
  );
};

export default ProcedureRow;
