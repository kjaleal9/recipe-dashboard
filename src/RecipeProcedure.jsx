import React, { Fragment } from 'react';
import {
  Box,
  Divider,
  FormControl,
  FormHelperText,
  InputLabel,
  List,
  ListItem,
  ListItemButton,
  ListItemIcon,
  ListItemText,
  MenuItem,
  Select,
  Stack,
  Paper,
  Typography,
  TextField,
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

// 20230224185204
// http://localhost:5000/recipes/latest

const procedure = {
  steps: [
    { stepName: 'T151 Allocate', stepType: 'Allocate' },
    { stepName: 'T151 Agitation', stepType: 'Run' },
    { stepName: 'R100 Allocate', stepType: 'Allocate' },
    { stepName: 'R100 Empty', stepType: 'Start' },
    { stepName: 'T151 Fill', stepType: 'Start' },
    { stepName: 'R100 Empty', stepType: 'Stop' },
    { stepName: 'T151 Fill', stepType: 'Stop' },
    { stepName: 'R100 Deallocate', stepType: 'Deallocate' },
    { stepName: 'T151 Deallocate', stepType: 'Deallocate' },
    { stepName: 'T151 Allocate', stepType: 'Allocate' },
    { stepName: 'T151 Agitation', stepType: 'Run' },
    { stepName: 'R100 Allocate', stepType: 'Allocate' },
    { stepName: 'R100 Empty', stepType: 'Start' },
    { stepName: 'T151 Fill', stepType: 'Start' },
    { stepName: 'R100 Empty', stepType: 'Stop' },
    { stepName: 'T151 Fill', stepType: 'Stop' },
    { stepName: 'R100 Deallocate', stepType: 'Deallocate' },
    { stepName: 'T151 Deallocate', stepType: 'Deallocate' },
    { stepName: 'T151 Allocate', stepType: 'Allocate' },
    { stepName: 'T151 Agitation', stepType: 'Run' },
    { stepName: 'R100 Allocate', stepType: 'Allocate' },
    { stepName: 'R100 Empty', stepType: 'Start' },
    { stepName: 'T151 Fill', stepType: 'Start' },
    { stepName: 'R100 Empty', stepType: 'Stop' },
    { stepName: 'T151 Fill', stepType: 'Stop' },
    { stepName: 'R100 Deallocate', stepType: 'Deallocate' },
    { stepName: 'T151 Deallocate', stepType: 'Deallocate' },
    { stepName: 'T151 Allocate', stepType: 'Allocate' },
    { stepName: 'T151 Agitation', stepType: 'Run' },
    { stepName: 'R100 Allocate', stepType: 'Allocate' },
    { stepName: 'R100 Empty', stepType: 'Start' },
    { stepName: 'T151 Fill', stepType: 'Start' },
    { stepName: 'R100 Empty', stepType: 'Stop' },
    { stepName: 'T151 Fill', stepType: 'Stop' },
    { stepName: 'R100 Deallocate', stepType: 'Deallocate' },
    { stepName: 'T151 Deallocate', stepType: 'Deallocate' },
  ],
};

const RecipeProcedure = props => {
  const { materials, materialClasses } = props;
  return (
    <Box
      sx={{ display: 'flex', justifyContent: 'space-between', height: '100%' }}
    >
      <Box sx={{ flexGrow: 1, display: 'flex', flexDirection: 'column' }}>
        <Typography
          component='h3'
          variant='h6'
          color='inherit'
          noWrap
          sx={{ ml: 4, mt: 2 }}
        >
          Procedure
        </Typography>
        {/* <Divider /> */}
        <Box
          sx={{
            height: '100%',
            overflowY: 'scroll',
            my: 1,
          }}
        >
          {procedure.steps.map((step, index) => (
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
              <Typography
                component='h5'
                variant='h5'
                color='inherit'
                noWrap
                sx={{ mr: 2 }}
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
              <TextField size='small' />
            </ListItemButton>
          ))}
        </Box>
      </Box>
      {/* <Divider orientation='vertical' variant='middle' flexItem /> */}
      <Box Box sx={{ width: '40%', p: 1, m: 2 }}>
        Test
        <FormControl sx={{ mb: 2, width: 350, alignSelf: 'center' }}>
          <InputLabel id='demo-simple-select-helper-label'>
            Material Class
          </InputLabel>
          <Select
            labelId='material-class-select-helper-label'
            id='material-class-select-helper'
            label='Material Class'
          >
            <MenuItem value=''>
              <em>None</em>
            </MenuItem>
            {materialClasses.map(materialClass => {
              return (
                <MenuItem
                  key={materialClass.ID}
                  value={materialClass}
                >{`${materialClass.Name}`}</MenuItem>
              );
            })}
          </Select>
          <FormHelperText></FormHelperText>
        </FormControl>
        <FormControl sx={{ mb: 2, width: 350, alignSelf: 'center' }}>
          <InputLabel id='demo-simple-select-helper-label'>Material</InputLabel>
          <Select
            labelId='material-select-helper-label'
            id='material-select-helper'
            label='Material'
          >
            <MenuItem value=''>
              <em>None</em>
            </MenuItem>
            {materials.map(material => {
              return (
                <MenuItem
                  key={material.ID}
                  value={material}
                >{`${material.SiteMaterialAlias} - ${material.Name}`}</MenuItem>
              );
            })}
          </Select>
          <FormHelperText></FormHelperText>
        </FormControl>
      </Box>
    </Box>
  );
};

export default RecipeProcedure;
