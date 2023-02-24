import React, { useState, useEffect } from 'react';

import {
  Box,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TablePagination,
  TableRow,
  Paper,
  Skeleton,
  Typography,
} from '@mui/material';

import EnhancedTableHead from './EnhancedTableHead';
import EnhancedTableToolbar from './EnhancedTableToolbar';
import EnhancedModal from './EnhancedModal';
import DeleteDialog from './DeleteDialog';
import Snackbar from './SnackBar';
import RecipeProcedure from './RecipeProcedure';

import { getComparator, stableSort } from './utilities';

const RecipeSearch = props => {
  const [order, setOrder] = useState('asc');
  const [orderBy, setOrderBy] = useState('RID');
  const [selected, setSelected] = useState({});
  const [page, setPage] = useState(0);
  const [rowsPerPage, setRowsPerPage] = useState(20);
  const [rows, setRows] = useState([]);
  const [fullDatabase, setFullDatabase] = useState([]);
  const [latestVersion, setLatestVersionRecipes] = useState([]);
  const [materials, setMaterials] = useState([]);
  const [materialClasses, setMaterialClasses] = useState([]);
  const [processClasses, setProcessClasses] = useState([]);
  const [requiredProcessClasses, setRequiredProcessClasses] = useState([]);
  const [openNewModal, setOpenNewModal] = useState(false);
  const [openDeleteDialog, setOpenDeleteDialog] = useState(false);
  const [openSnackBar, setOpenSnackbar] = useState(false);
  const [snackPack, setSnackPack] = React.useState([]);
  const [messageInfo, setMessageInfo] = React.useState(undefined);
  const [mode, setMode] = useState('Search');

  const refreshTable = () => {
    if (mode === 'Search') {
      console.time('Get all data');
      const getRecipes = () =>
        fetch('http://localhost:5000/recipes').then(response =>
          response.json()
        );

      const getLatestRecipes = () =>
        fetch('http://localhost:5000/recipes/latest').then(response =>
          response.json()
        );

      const getMaterials = () =>
        fetch('http://localhost:5000/materials').then(response =>
          response.json()
        );

      const getMaterialClasses = () =>
        fetch('http://localhost:5000/material-classes').then(response =>
          response.json()
        );
      const getProcessClasses = () =>
        fetch('http://localhost:5000/process-classes').then(response =>
          response.json()
        );
      const getRequiredProcessClasses = () =>
        fetch('http://localhost:5000/process-classes/required').then(response =>
          response.json()
        );

      function getAllData() {
        return Promise.all([
          getRecipes(),
          getLatestRecipes(),
          getMaterials(),
          getMaterialClasses(),
          getProcessClasses(),
          getRequiredProcessClasses(),
        ]);
      }

      getAllData().then(
        ([
          allRecipes,
          latestRecipes,
          allMaterials,
          allMaterialClasses,
          allProcessClasses,
          allRequiredProcessClasses,
        ]) => {
          setFullDatabase(allRecipes);
          setLatestVersionRecipes(latestRecipes);
          setRows(latestRecipes);
          setMaterials(allMaterials);
          setMaterialClasses(allMaterialClasses);
          setProcessClasses(allProcessClasses);
          setRequiredProcessClasses(allRequiredProcessClasses);
        }
      );
      console.timeEnd('Get all data');
    }
  };

  useEffect(() => {
    refreshTable();
  }, [mode]);

  const handleRequestSort = (event, property) => {
    const isAsc = orderBy === property && order === 'asc';
    setOrder(isAsc ? 'desc' : 'asc');
    setOrderBy(property);
  };

  const handleClick = (event, name, version, row) => {
    const selectedRow = selected.RID === name && selected.Version === version;
    let newSelected = '';

    if (selectedRow === false) {
      newSelected = row;
    }

    setSelected(newSelected);
  };

  const handleDoubleClick = (event, row) => {
    setSelected(row);
    setMode('View');
    setOpenNewModal(true);
  };

  const handleDelete = () => {
    setOpenDeleteDialog(true);
  };

  const handleChangePage = (event, newPage) => {
    setPage(newPage);
  };

  const handleChangeRowsPerPage = event => {
    setRowsPerPage(parseInt(event.target.value, 10));
    setPage(0);
  };

  const isChecked = checked => {
    checked ? setRows(fullDatabase) : setRows(latestVersion);
  };

  // const isSelected = row => selected.indexOf(row) !== -1;
  function isSelected(RID, version) {
    return selected.RID === RID && selected.Version === version;
  }

  // Avoid a layout jump when reaching the last page with empty rows.
  const emptyRows =
    page > 0 ? Math.max(0, (1 + page) * rowsPerPage - rows.length) : 0;

  return (
    <Box sx={{ width: '100%' }}>
      <Paper sx={{ width: '100%', mb: 2 }}>
        {latestVersion.length > 0 ? (
          mode === 'Procedure' ? (
            <RecipeProcedure />
          ) : (
            <Box>
              <EnhancedTableToolbar
                anyRowSelected={selected !== ''}
                isChecked={isChecked}
                setOpen={setOpenNewModal}
                setMode={setMode}
                selected={selected}
                setSelected={setSelected}
                handleDelete={handleDelete}
              />
              <EnhancedModal
                open={openNewModal}
                setOpen={setOpenNewModal}
                mode={mode}
                setMode={setMode}
                rows={rows}
                selected={selected}
                setSelected={setSelected}
                materials={materials}
                materialClasses={materialClasses}
                processClasses={processClasses}
                requiredProcessClasses={requiredProcessClasses}
                refreshTable={refreshTable}
              />
              <DeleteDialog
                open={openDeleteDialog}
                setOpen={setOpenDeleteDialog}
                selected={selected}
                setSnackPack={setSnackPack}
                refreshTable={refreshTable}
              />
              <Snackbar
                open={openSnackBar}
                setOpen={setOpenSnackbar}
                snackPack={snackPack}
                setSnackPack={setSnackPack}
                messageInfo={messageInfo}
                setMessageInfo={setMessageInfo}
              />
              <TableContainer>
                <Table
                  sx={{ minWidth: 750 }}
                  aria-labelledby='tableTitle'
                  size='small'
                >
                  <EnhancedTableHead
                    order={order}
                    orderBy={orderBy}
                    onRequestSort={handleRequestSort}
                  />
                  <TableBody>
                    {stableSort(rows, getComparator(order, orderBy))
                      .slice(
                        page * rowsPerPage,
                        page * rowsPerPage + rowsPerPage
                      )
                      .map((row, index) => {
                        const isItemSelected = isSelected(row.RID, row.Version);
                        const labelId = `enhanced-table-checkbox-${index}`;

                        return (
                          <TableRow
                            hover
                            onClick={event =>
                              handleClick(event, row.RID, row.Version, row)
                            }
                            onDoubleClick={event =>
                              handleDoubleClick(event, row)
                            }
                            role='checkbox'
                            aria-checked={isItemSelected}
                            tabIndex={-1}
                            key={row.RID + row.Version}
                            selected={isItemSelected}
                            sx={{ cursor: 'pointer' }}
                          >
                            <TableCell
                              component='th'
                              id={labelId}
                              scope='row'
                              sx={{ width: '200px' }}
                            >
                              {row.RID}
                            </TableCell>
                            <TableCell align='right' sx={{ width: '80px' }}>
                              {row.Version}
                            </TableCell>
                            <TableCell align='left' sx={{ width: '200px' }}>
                              {new Date(row.VersionDate).toDateString()}
                            </TableCell>
                            <TableCell align='left' sx={{ width: '100px' }}>
                              {row.RecipeType}
                            </TableCell>
                            <TableCell align='left' sx={{ width: '300px' }}>
                              {row.Description}
                            </TableCell>
                            <TableCell align='left' sx={{ width: '100px' }}>
                              {row.Status}
                            </TableCell>
                            <TableCell align='right' sx={{ width: '100px' }}>
                              {row.ProductID}
                            </TableCell>
                            <TableCell align='left' sx={{ width: '300px' }}>
                              {row.Name}
                            </TableCell>
                          </TableRow>
                        );
                      })}
                    {emptyRows > 0 && (
                      <TableRow
                        style={{
                          height: 33 * emptyRows,
                        }}
                      >
                        <TableCell colSpan={6} />
                      </TableRow>
                    )}
                  </TableBody>
                </Table>
              </TableContainer>
              <TablePagination
                rowsPerPageOptions={[5, 10, 15, 20]}
                component='div'
                count={rows.length}
                rowsPerPage={rowsPerPage}
                page={page}
                onPageChange={handleChangePage}
                onRowsPerPageChange={handleChangeRowsPerPage}
              />
            </Box>
          )
        ) : (
          <Box>
            <Skeleton variant='rectangular' width={210} height={118} />
            <Typography>LOADING</Typography>
          </Box>
        )}
      </Paper>
    </Box>
  );
};

export default RecipeSearch;
