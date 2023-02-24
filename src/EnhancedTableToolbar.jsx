import React from 'react';
import PropTypes from 'prop-types';

import Button from '@mui/material/Button';
import ButtonGroup from '@mui/material/ButtonGroup';
import IconButton from '@mui/material/IconButton';
import InputAdornment from '@mui/material/InputAdornment';
import SearchIcon from '@mui/icons-material/Search';
import Toolbar from '@mui/material/Toolbar';
import {
  Box,
  Input,
  FormControl,
  InputLabel,
  Switch,
  FormGroup,
  FormControlLabel,
} from '@mui/material';
import Typography from '@mui/material/Typography';

const EnhancedTableToolbar = props => {
  const {
    anyRowSelected,
    isChecked,
    setMode,
    setOpen,
    selected,
    setSelected,
    rows,
    handleDelete,
  } = props;
  const [checked, setChecked] = React.useState(true);

  const handleChange = event => {
    setChecked(event.target.checked);
    isChecked(event.target.checked, checked);
  };

  const handleOpen = mode => {
    setMode(mode);
    setOpen(true);
  };

  return (
    <Toolbar
      sx={{
        pl: { sm: 2 },
        pr: { xs: 1, sm: 1 },
        display: 'flex',
        justifyContent: 'space-between',
        // width: '100%',
      }}
    >
      <FormControl sx={{ m: 1, width: '50ch' }} variant='standard'>
        <InputLabel htmlFor='standard-adornment-password'>
          Recipe Search
        </InputLabel>
        <Input
          id='standard-adornment-password'
          type='text'
          endAdornment={
            <InputAdornment position='end'>
              <IconButton
                aria-label='toggle password visibility'
                onClick={() => console.log('click')}
                onMouseDown={() => console.log('mousedown')}
              >
                <SearchIcon />
                {/* {showPassword ? <VisibilityOff /> : <Visibility />} */}
              </IconButton>
            </InputAdornment>
          }
        />
      </FormControl>

      <Box sx={{ display: 'flex' }}>
        <Box>
          <FormGroup>
            <FormControlLabel
              control={
                <Switch
                  onChange={handleChange}
                  inputProps={{ 'aria-label': 'controlled' }}
                />
              }
              label='Show All'
            />
          </FormGroup>
        </Box>
        <ButtonGroup color='primary' sx={{ mx: 2 }}>
          <Button variant='contained' onClick={() => handleOpen('New')}>
            New
          </Button>
          <Button
            variant='contained'
            onClick={() => handleOpen('Copy')}
            disabled={!anyRowSelected}
          >
            Copy
          </Button>
          <Button
            variant='contained'
            disabled={!anyRowSelected}
            onClick={handleDelete}
          >
            Delete
          </Button>
        </ButtonGroup>
        <ButtonGroup color='secondary' sx={{ mx: 2 }}>
          <Button disabled={!anyRowSelected}>Import</Button>
          <Button disabled={!anyRowSelected}>Export</Button>
        </ButtonGroup>
      </Box>
    </Toolbar>
  );
};

EnhancedTableToolbar.propTypes = {
  anyRowSelected: PropTypes.bool.isRequired,
};

export default EnhancedTableToolbar;
