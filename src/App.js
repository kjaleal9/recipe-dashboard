import * as React from 'react';
import { Outlet } from 'react-router-dom';

import { createTheme, ThemeProvider } from '@mui/material/styles';
import CssBaseline from '@mui/material/CssBaseline';

import { Box, Container, Grid, Toolbar } from '@mui/material';

import CustomAppBar from './Components/CustomAppBar/CustomAppBar';
import CustomDrawer from './Components/CustomDrawer/CustomDrawer';

// Custom styling for app bar

// Change theme for app. "light" or "dark" mode
const mdTheme = createTheme({
  palette: {
    mode: 'dark',
    appBar: { main: '#3D7C98',contrastText: '#fff', },
  },
});

console.log(mdTheme)

function App() {
  const [open, setOpen] = React.useState(false);

  const toggleDrawer = () => {
    setOpen(!open);
  };


  return (
    <ThemeProvider theme={mdTheme}>
      <Box sx={{ display: 'flex' }}>
        <CssBaseline />
        <CustomAppBar open={open} toggleDrawer={toggleDrawer} />
        <CustomDrawer open={open} toggleDrawer={toggleDrawer} />
        <Box
          component='main'
          sx={{
            backgroundColor: theme =>
              theme.palette.mode === 'light'
                ? theme.palette.grey[200]
                : theme.palette.grey[900],
            flexGrow: 1,
            height: '100vh',
            overflow: 'auto',
          }}
        >
          <Toolbar />
          <Container maxWidth='xlg' sx={{ mt: 2, mb: 2 }}>
            <Grid container spacing={3}>
              <Grid item xs={12}>
                <Outlet /> {/* !!!This is where the routes will render!!! */}
              </Grid>
            </Grid>
          </Container>
        </Box>
      </Box>
    </ThemeProvider>
  );
}

export default App;
