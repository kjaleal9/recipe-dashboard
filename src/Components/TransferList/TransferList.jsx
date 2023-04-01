import * as React from 'react';
import Grid from '@mui/material/Grid';
import List from '@mui/material/List';
import ListItemButton from '@mui/material/ListItem';

import ListItemText from '@mui/material/ListItemText';

import Button from '@mui/material/Button';
import Paper from '@mui/material/Paper';
import {
  Tooltip,
  Typography,
  FormControl,
  InputLabel,
  Select,
  Box,
  MenuItem,
} from '@mui/material';

function not(a, b) {
  return a.filter(value => b.indexOf(value) === -1);
}

function intersection(a, b) {
  return a.filter(value => b.indexOf(value) !== -1);
}

export default function TransferList(props) {
  const {
    setChecked,
    setLeft,
    setRight,
    handleToggle,
    left,
    right,
    checked,
    processClasses,
    requiredProcessClasses,
    mode,
    selected,
    mainProcessClass,
    setMainProcessClass,
  } = props;

  const leftChecked = intersection(checked, left);
  const rightChecked = intersection(checked, right);

  React.useEffect(() => {
    if (mode === 'New') {
      setLeft(processClasses.map(item => item.Name));
      setRight([]);
    }

    if (mode === 'View') {
      const selectedPCs = requiredProcessClasses.filter(
        reqPC =>
          reqPC.Recipe_RID === selected.RID &&
          reqPC.Recipe_Version === selected.Version
      );
      const selectedPCNames = selectedPCs.map(PC => PC.ProcessClass_Name);

      setRight(
        processClasses
          .filter(item => selectedPCNames.includes(item.Name))
          .map(item => item.Name)
      );
      setLeft(
        processClasses
          .filter(item => !selectedPCNames.includes(item.Name))
          .map(item => item.Name)
      );

      const mainPCSelected = selectedPCs.find(PC => PC.IsMainBatchUnit === 1);

      if (mainPCSelected) {
        setMainProcessClass(mainPCSelected.ProcessClass_Name);
      } else {
        setMainProcessClass('');
      }
    }
  }, []);

  const handleCheckedRight = () => {
    setRight(right.concat(leftChecked));
    setLeft(not(left, leftChecked));
    setChecked(not(checked, leftChecked));
  };

  const handleCheckedLeft = () => {
    setLeft(left.concat(rightChecked));
    setRight(not(right, rightChecked));
    setChecked(not(checked, rightChecked));
  };

  const customList = (items, fullList) => (
    <Paper sx={{ width: 200, height: 300, overflow: 'auto' }}>
      <List dense component='div' role='list'>
        {items.map(value => {
          const labelId = `transfer-list-item-${value}-label`;
          const toolTipText = fullList.find(
            processClass => processClass.Name === value
          ).Description;
          return (
            <Tooltip key={value} title={`${toolTipText}`}>
              <ListItemButton
                role='listitem'
                button
                selected={checked.includes(value)}
                disabled={mode === 'View'}
                onClick={handleToggle(value)}
                sx={{ cursor: 'pointer' }}
              >
                <ListItemText id={labelId} primary={`${value}`} />
              </ListItemButton>
            </Tooltip>
          );
        })}
      </List>
    </Paper>
  );

  return (
    <Grid container spacing={2} justifyContent='center' alignItems='center'>
      <Grid item>
        <Grid container direction='column' alignItems='center'>
          <Grid>
            <Typography
              component='h5'
              variant='subtitle'
              color='inherit'
              noWrap
              sx={{ mb: 2 }}
            >
              Available Process Classes
            </Typography>
          </Grid>
          <Grid item>{customList(left, processClasses)}</Grid>
        </Grid>
      </Grid>
      <Grid item>
        <Grid container direction='column' alignItems='center'>
          <Button
            sx={{ my: 0.5 }}
            variant='outlined'
            size='small'
            onClick={handleCheckedRight}
            disabled={leftChecked.length === 0}
            aria-label='move selected right'
          >
            &gt;
          </Button>
          <Button
            sx={{ my: 0.5 }}
            variant='outlined'
            size='small'
            onClick={handleCheckedLeft}
            disabled={rightChecked.length === 0}
            aria-label='move selected left'
          >
            &lt;
          </Button>
        </Grid>
      </Grid>
      <Grid item>
        <Grid container direction='column' alignItems='center'>
          <Grid>
            <Typography
              component='h5'
              variant='subtitle'
              color='inherit'
              noWrap
              sx={{ mb: 2 }}
            >
              Selected Process Classes
            </Typography>
          </Grid>
          <Grid item>{customList(right, processClasses)}</Grid>
        </Grid>
      </Grid>
      <Grid item sx={{ mt: 2 }}>
        <FormControl sx={{ width: 200 }}>
          <InputLabel id='main-process-class-label'>
            Main Process Class
          </InputLabel>
          <Select
            labelId='main-process-class-label'
            id='main-process-class'
            value={mainProcessClass}
            label='Main Process Class'
            disabled={mode === 'View'}
            onChange={event => setMainProcessClass(event.target.value)}
          >
            <MenuItem value=''>
              <em>None</em>
            </MenuItem>
            {right.map(value => {
              return (
                <MenuItem key={value} value={value}>{`${value}`}</MenuItem>
              );
            })}
          </Select>
        </FormControl>
      </Grid>
    </Grid>
  );
}
