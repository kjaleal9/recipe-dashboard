import React from 'react';

import { Badge, IconButton, Toolbar, Typography } from '@mui/material';

import MuiAppBar from '@mui/material/AppBar';
import { styled } from '@mui/material/styles';
import { createTheme, ThemeProvider } from '@mui/material/styles';

import NotificationsIcon from '@mui/icons-material/Notifications';
import SettingsIcon from '@mui/icons-material/Settings';
import MenuIcon from '@mui/icons-material/Menu';

const AppBar = styled(MuiAppBar, {
  shouldForwardProp: prop => prop !== 'open',
})(({ theme, open }) => ({
  zIndex: theme.zIndex.drawer + 1,
  transition: theme.transitions.create(['width', 'margin'], {
    easing: theme.transitions.easing.sharp,
    duration: theme.transitions.duration.leavingScreen,
  }),
  ...(open && {
    marginLeft: +process.env.REACT_APP_DRAWER_WIDTH,
    width: `calc(100% - ${+process.env.REACT_APP_DRAWER_WIDTH}px)`,
    transition: theme.transitions.create(['width', 'margin'], {
      easing: theme.transitions.easing.sharp,
      duration: theme.transitions.duration.enteringScreen,
    }),
  }),
}));

const mdTheme = createTheme({
  palette: {
    // mode: 'light',
    // mode: 'dark',
    // primary: {
    //   main: '#90caf9',
    // },
    // secondary: {
    //   main: '#ce93d8',
    // },
    // background: {
    //   default: '#121212',
    //   paper: '#121212',
    // },
  },
});

const CustomAppBar = ({ open, toggleDrawer }) => {
  return (
    
      <AppBar position='absolute' open={open}>
        <Toolbar
          sx={{
            pr: '24px', // keep right padding when drawer closed
          }}
        >
          <IconButton
            edge='start'
            color='inherit'
            aria-label='open drawer'
            onClick={toggleDrawer}
            sx={{
              marginRight: '36px',
              ...(open && { display: 'none' }),
            }}
          >
            <MenuIcon />
          </IconButton>
          <Typography
            component='h1'
            variant='h6'
            color='inherit'
            noWrap
            sx={{ flexGrow: 1 }}
          >
            Recipe Dashboard
          </Typography>
          <IconButton color='inherit'>
            <Badge badgeContent={4} color='secondary'>
              <NotificationsIcon />
            </Badge>
          </IconButton>
          <IconButton color='inherit'>
            <SettingsIcon />
          </IconButton>
        </Toolbar>
      </AppBar>
  
  );
};

export default CustomAppBar;
