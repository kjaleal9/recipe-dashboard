import React, { useState, useEffect } from "react";

import { Box, Paper, Skeleton, Typography, Grid } from "@mui/material";

// TODO rename components
import EnhancedTableToolbar from "../Components/Table/EnhancedTableToolbar/EnhancedTableToolbar";
import EnhancedModal from "../Components/Modals/EnhancedModal/EnhancedModal";
import DeleteDialog from "../Components/DeleteDialog/DeleteDialog";
import Snackbar from "../Components/SnackBar/SnackBar";
import RecipeView from "../Components/RecipeView/RecipeView";
import CustomTableBody from "../Components/Table/CustomTableBody/CustomTableBody";

const RecipeSearch = () => {
  const [fullDatabase, setFullDatabase] = useState([]);
  const [latestVersion, setLatestVersionRecipes] = useState([]);
  const [selected, setSelected] = useState("");
  const [page, setPage] = useState(0);
  const [rows, setRows] = useState([]);
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
        +max["Version"] > +recipe["Version"] ? max : recipe
      )
    );
  }
  const getRecipes = () =>
    fetch("/recipes").then((response) => response.json());
  const getMaterials = () =>
    fetch("/materials").then((response) => response.json());
  const getMaterialClasses = () =>
    fetch("/materials/classes").then((response) => response.json());
  const getProcessClasses = () =>
    fetch("/process-classes").then((response) => response.json());
  const getRequiredProcessClasses = () =>
    fetch("/process-classes/required").then((response) => response.json());

  const refreshTable = () => {
    console.time("Get all data");

    function getAllData() {
      return Promise.all([
        getRecipes(),
        getMaterials(),
        getMaterialClasses(),
        getProcessClasses(),
        getRequiredProcessClasses(),
      ]);
    }

    getAllData()
      .then(
        ([
          allRecipes,
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
      )
      .catch((err) => console.log(err));
    console.timeEnd("Get all data");
  };

  const filterRows = () => {
    console.time("Filter");
    const { showAll, status } = filter;
    let filteredRows;
    showAll ? (filteredRows = fullDatabase) : (filteredRows = latestVersion);
    filteredRows = filteredRows.filter(
      (row) =>
        (!status || row.Status === status) &&
        row.RID.toLowerCase().includes(filter.search)
    );
    setRows(filteredRows);
    console.timeEnd("Filter");
  };

  useEffect(() => {
    refreshTable();
  }, []);

  // Table filter useEffect. Runs everytime the filter changes
  useEffect(() => {
    filterRows();
  }, [filter]);

  const handleDelete = () => {
    setOpenDeleteDialog(true);
  };

  return (
    <Grid container spacing={2} sx={{ height: "88vh" }}>
      <Grid item xs={12} md={12} lg={9}>
        <Paper sx={{ height: "100%" }}>
          {fullDatabase.length > 0 ? (
            <Box>
              <EnhancedTableToolbar
                selected={selected}
                checked={checked}
                filter={filter}
                anyRowSelected={selected !== ""}
                setOpen={setOpenNewModal}
                setMode={setMode}
                setSelected={setSelected}
                setShowAllChecked={setShowAllChecked}
                setPage={setPage}
                setFilter={setFilter}
                handleDelete={handleDelete}
              />
              <EnhancedModal
                selected={selected}
                mode={mode}
                open={openNewModal}
                rows={rows}
                materials={materials}
                materialClasses={materialClasses}
                processClasses={processClasses}
                requiredProcessClasses={requiredProcessClasses}
                setOpen={setOpenNewModal}
                setMode={setMode}
                setSelected={setSelected}
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
              <CustomTableBody
                setOpen={setOpenNewModal}
                mode={mode}
                setMode={setMode}
                rows={rows}
                selected={selected}
                setSelected={setSelected}
                page={page}
                setPage={setPage}
              />
            </Box>
          ) : (
            <Box>
              <Skeleton variant="rectangular" width={210} height={118} />
              <Typography>LOADING</Typography>
            </Box>
          )}
        </Paper>
      </Grid>
      <Grid item xs={12} md={12} lg={3}>
        <Paper sx={{ height: "100%" }}>
          {selected ? (
            <RecipeView
              selected={selected}
              setMode={setMode}
              setOpen={setOpenNewModal}
            />
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
      </Grid>
    </Grid>
  );
};

export default RecipeSearch;
