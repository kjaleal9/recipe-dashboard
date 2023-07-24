import React, { useState, useEffect } from 'react';
import {
  TextField,
  Grid,
  Paper,
  Table,
  TableHead,
  TableRow,
  TableCell,
  TableBody,
} from '@mui/material';

const Materials = ({ label, value, onChange }) => {
  const [errorLog, setErrorLog] = useState([]);

  const loadData = () => {
    const getErrors = () => fetch('/errors').then(response => response.json());

    function getAllData() {
      return Promise.all([getErrors()]);
    }

    getAllData()
      .then(([errors]) => {
        setErrorLog(errors);
      })
      .catch(err => console.log(err, 'uih'));
  };
  useEffect(() => {
    loadData();

  }, []);

  return (
    <Grid container spacing={2} sx={{ height: '88vh' }}>
      <Grid item xs={12} md={8} lg={12}>
        <Paper>
          <Table sx={{ }} aria-label='simple table'>
            <TableHead>
              <TableRow>
                {errorLog[0].keys.map((key)=>{
                  console.log(errorLog[0])
                  return (<TableCell>{key}</TableCell>)

                })}
                
              </TableRow>
            </TableHead>
            <TableBody>
              {errorLog.map(row => (
                <TableRow
                  key={row.LogID}
                  // sx={{ '&:last-child td, &:last-child th': { border: 0 } }}
                >
                  <TableCell>{row.LogID}</TableCell>
                  <TableCell>{row.ErrorValue}</TableCell>
                  <TableCell>{row.ErrorMessage}</TableCell>
                  <TableCell>{row.ErrorTime}</TableCell>
                  <TableCell>{row.ObjectName}</TableCell>
                  <TableCell>{row.UserName}</TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </Paper>
      </Grid>
    </Grid>
  );
};

export default Materials;
