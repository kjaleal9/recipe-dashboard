import * as React from 'react';
import Button from '@mui/material/Button';
import Dialog from '@mui/material/Dialog';
import DialogActions from '@mui/material/DialogActions';
import DialogContent from '@mui/material/DialogContent';
import DialogContentText from '@mui/material/DialogContentText';
import DialogTitle from '@mui/material/DialogTitle';

const AlertDialog = props => {
  const { open, setOpen, selected, setSnackPack,refreshTable } = props;

  const handleClose = choice => {
    if (choice === 'No') {
      setOpen(false);
    } else if (choice === 'Yes') {
      fetch('http://localhost:5000/recipes', {
        method: 'DELETE', // *GET, POST, PUT, DELETE, etc.
        mode: 'cors', // no-cors, *cors, same-origin
        cache: 'no-cache', // *default, no-cache, reload, force-cache, only-if-cached
        credentials: 'same-origin', // include, *same-origin, omit
        headers: {
          'Content-Type': 'application/json',
          // 'Content-Type': 'application/x-www-form-urlencoded',
        },
        redirect: 'follow', // manual, *follow, error
        referrerPolicy: 'no-referrer', // no-referrer, *no-referrer-when-downgrade, origin, origin-when-cross-origin, same-origin, strict-origin, strict-origin-when-cross-origin, unsafe-url
        body: JSON.stringify({
          RID: selected.RID,
          Version: selected.Version,
        }),
      })
        .then(response => response.json())
        .then(data => {
          refreshTable()
        });

      setSnackPack(prev => [
        ...prev,
        { message: 'Recipe Deleted', key: new Date().getTime() },
      ]);
      setOpen(false);
    }
  };

  return (
    <Dialog
      open={open}
      onClose={handleClose}
      aria-labelledby='alert-dialog-title'
      aria-describedby='alert-dialog-description'
    >
      <DialogTitle id='alert-dialog-title'>{'Delete Recipe?'}</DialogTitle>
      <DialogContent>
        <DialogContentText id='alert-dialog-description'>
          Are you sure you want to delete this recipe?
        </DialogContentText>
      </DialogContent>
      <DialogActions>
        <Button onClick={() => handleClose('No')}>Disagree</Button>
        <Button onClick={() => handleClose('Yes')} autoFocus>
          Agree
        </Button>
      </DialogActions>
    </Dialog>
  );
};

export default AlertDialog;
