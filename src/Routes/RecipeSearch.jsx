import React, { useState, useEffect } from "react";

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
} from "@mui/material";

import EnhancedTableHead from "../Components/Table/EnhancedTableHead/EnhancedTableHead";
import EnhancedTableToolbar from "../Components/Table/EnhancedTableToolbar/EnhancedTableToolbar";
import EnhancedModal from "../Components/EnhancedModal/EnhancedModal";
import DeleteDialog from "../Components/DeleteDialog/DeleteDialog";
import Snackbar from "../Components/SnackBar/SnackBar";
import RecipeView from "../Components/RecipeView/RecipeView";

import { getComparator, stableSort } from "../utilities";

const RecipeSearch = (props) => {
  const [order, setOrder] = useState("asc");
  const [orderBy, setOrderBy] = useState("RID");
  const [selected, setSelected] = useState("");
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
  const [mode, setMode] = useState("Search");
  const [checked, setShowAllChecked] = React.useState(false);
  const [filter, setFilter] = useState({
    showAll: false,
    approved: false,
    registered: false,
    obsolete: false,
    search: "",
  });

  function groupRecipes(rows) {
    function groupBy(objectArray, property) {
      return objectArray.reduce((acc, obj) => {
        const key = obj[property];
        if (!acc[key]) {
          acc[key] = [];
        }
        // Add object to list for given key's value
        acc[key].push(obj);
        return acc;
      }, {});
    }

    const groupedRecipes = groupBy(rows, "RID");

    return Object.keys(groupedRecipes).map((key) =>
      groupedRecipes[key].reduce((max, recipe) =>
        max["Version"] > recipe["Version"] ? max : recipe
      )
    );
  }

  const refreshTable = () => {
    if (mode === "Search" || mode === "Procedure") {
      console.time("Get all data");
      const getRecipes = () =>
        fetch("/recipes").then((response) => response.json());

      // const getLatestRecipes = () =>
      //   fetch("/recipes/latest").then((response) => response.json());

      const getMaterials = () =>
        fetch("/materials").then((response) => response.json());
      const getMaterialClasses = () =>
        fetch("/material-classes").then((response) => response.json());
      const getProcessClasses = () =>
        fetch("/process-classes").then((response) => response.json());
      const getRequiredProcessClasses = () =>
        fetch("/process-classes/required").then((response) => response.json());

      function getAllData() {
        return Promise.all([
          getRecipes(),
          // getLatestRecipes(),
          getMaterials(),
          getMaterialClasses(),
          getProcessClasses(),
          getRequiredProcessClasses(),
        ]);
      }

      getAllData().then(
        ([
          allRecipes,
          // latestRecipes,
          allMaterials,
          allMaterialClasses,
          allProcessClasses,
          allRequiredProcessClasses,
        ]) => {
          setFullDatabase(allRecipes);
          setLatestVersionRecipes(groupRecipes(allRecipes));
          setRows(groupRecipes(allRecipes));
          setMaterials(allMaterials);
          setMaterialClasses(allMaterialClasses);
          setProcessClasses(allProcessClasses);
          setRequiredProcessClasses(allRequiredProcessClasses);
        }
      );
      console.timeEnd("Get all data");
      console.log(selected);
    }
  };

  useEffect(() => {
    refreshTable();
    console.log("Test");
  }, [mode]);

  useEffect(() => {
    const { showAll, approved, registered, obsolete } = filter;

    let filteredRows;

    showAll ? (filteredRows = fullDatabase) : (filteredRows = latestVersion);

    if (approved) {
      filteredRows = filteredRows.filter((row) => row.Status === "Approved");
    }
    if (registered) {
      filteredRows = filteredRows.filter((row) => row.Status === "Registered");
    }
    if (obsolete) {
      filteredRows = filteredRows.filter((row) => row.Status === "Obsolete");
    }

    filteredRows = filteredRows.filter((row) =>
      row.RID.toLowerCase().includes(filter.search)
    );

    setRows(filteredRows);
    console.log(filter);
  }, [filter]);

  const handleRequestSort = (event, property) => {
    const isAsc = orderBy === property && order === "asc";
    setOrder(isAsc ? "desc" : "asc");
    setOrderBy(property);
  };

  const handleClick = (event, name, version, row) => {
    const selectedRow = selected.RID === name && selected.Version === version;
    let newSelected = "";

    if (selectedRow === false) {
      newSelected = row;
    }

    setSelected(newSelected);
  };

  const handleDoubleClick = (event, row) => {
    setSelected(row);
    setMode("View");
    setOpenNewModal(true);
  };

  const handleDelete = () => {
    setOpenDeleteDialog(true);
  };

  const handleChangePage = (event, newPage) => {
    setPage(newPage);
  };

  const handleChangeRowsPerPage = (event) => {
    setRowsPerPage(parseInt(event.target.value, 10));
    setPage(0);
  };

  // const isSelected = row => selected.indexOf(row) !== -1;
  function isSelected(RID, version) {
    return selected.RID === RID && selected.Version === version;
  }

  // Avoid a layout jump when reaching the last page with empty rows.
  const emptyRows =
    page > 0 ? Math.max(0, (1 + page) * rowsPerPage - rows.length) : 0;

  const handleSearch = (event) => {
    console.log(event.target.value);
    setFilter({ ...filter, search: event.target.value.toLowerCase() });
  };

  return (
    <Box
      sx={{
        width: "100%",
        display: "flex",
        flexDirection: "row",
        justifyContent: "space-around",
      }}
    >
      <Paper sx={{ height: "88vh", width: "54vw" }}>
        {fullDatabase.length > 0 ? (
          <Box>
            <EnhancedTableToolbar
              anyRowSelected={selected !== ""}
              setOpen={setOpenNewModal}
              setMode={setMode}
              selected={selected}
              setSelected={setSelected}
              handleDelete={handleDelete}
              handleSearch={handleSearch}
              checked={checked}
              setShowAllChecked={setShowAllChecked}
              setPage={setPage}
              filter={filter}
              setFilter={setFilter}
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
                aria-labelledby="tableTitle"
                size="small"
              >
                <EnhancedTableHead
                  order={order}
                  orderBy={orderBy}
                  onRequestSort={handleRequestSort}
                />
                <TableBody>
                  {stableSort(rows, getComparator(order, orderBy))
                    .slice(page * rowsPerPage, page * rowsPerPage + rowsPerPage)
                    .map((row, index) => {
                      const isItemSelected = isSelected(row.RID, row.Version);
                      const labelId = `enhanced-table-checkbox-${index}`;

                      return (
                        <TableRow
                          hover
                          onClick={(event) =>
                            handleClick(event, row.RID, row.Version, row)
                          }
                          onDoubleClick={(event) =>
                            handleDoubleClick(event, row)
                          }
                          role="checkbox"
                          aria-checked={isItemSelected}
                          tabIndex={-1}
                          key={row.RID + row.Version}
                          selected={isItemSelected}
                          sx={{ cursor: "pointer" }}
                        >
                          <TableCell
                            component="th"
                            id={labelId}
                            scope="row"
                            sx={{ width: "200px" }}
                          >
                            {row.RID}
                          </TableCell>
                          <TableCell align="right" sx={{ width: "80px" }}>
                            {row.Version}
                          </TableCell>
                          <TableCell align="left" sx={{ width: "80px" }}>
                            {new Date(row.VersionDate).toLocaleDateString()}
                          </TableCell>
                          {/* <TableCell align="left" sx={{ width: "100px" }}>
                              {row.RecipeType}
                            </TableCell> */}
                          <TableCell align="left" sx={{ width: "200px" }}>
                            {row.Description}
                          </TableCell>
                          <TableCell align="left" sx={{ width: "100px" }}>
                            {row.Status}
                          </TableCell>
                          <TableCell align="right" sx={{ width: "100px" }}>
                            {row.ProductID}
                          </TableCell>
                          <TableCell align="left" sx={{ width: "200px" }}>
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
              component="div"
              count={rows.length}
              rowsPerPage={rowsPerPage}
              page={page}
              onPageChange={handleChangePage}
              onRowsPerPageChange={handleChangeRowsPerPage}
            />
          </Box>
        ) : (
          <Box>
            <Skeleton variant="rectangular" width={210} height={118} />
            <Typography>LOADING</Typography>
          </Box>
        )}{" "}
      </Paper>
      <Paper sx={{ width: "30vw", flexGrow: 1, ml:2 }}>
        {selected ? (
          <RecipeView selected={selected} />
        ) : (
          <Box
            sx={{
              display: "flex",
              justifyContent: "center",
              alignItems: "center",
              height: "100%",
            }}
          >
            <Typography
              component="h2"
              variant="p"
              color="inherit"
              sx={{ alignSelf: "center", pb: 0.25 }}
            >
              Select a Recipe to view details
            </Typography>
          </Box>
        )}
      </Paper>
    </Box>
  );
};

export default RecipeSearch;
