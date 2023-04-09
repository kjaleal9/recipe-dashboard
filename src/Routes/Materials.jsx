import React, { useState } from 'react';
import { TextField, Grid, Paper } from '@mui/material';
import MaterialArray from '../Components/MaterialArray/MaterialArray';

const Materials = ({ label, value, onChange }) => {
  const [inputValue, setInputValue] = useState(value);

  const handleInputChange = event => {
    const newValue = event.target.value;
    setInputValue(newValue);
    onChange(newValue);
  };

  return (
    <Grid container spacing={2} sx={{ height: '88vh' }}>
      <Grid item xs={12} md={8} lg={6}>
        <Paper>
          <MaterialArray />
        </Paper>
      </Grid>
    </Grid>
  );
};

export default Materials;
