import React, { useState, useEffect } from "react";
import { useParams } from "react-router-dom";
import {
  Grid,
  Paper,
  Box,
  InputLabel,
  FormControl,
  FormHelperText,
  Typography,
  Divider,
  Select,
  MenuItem,
  Button,
  TextField,
  Chip,
} from "@mui/material";
import SaveIcon from "@mui/icons-material/Save";
import EditIcon from "@mui/icons-material/Edit";
import CancelIcon from "@mui/icons-material/Cancel";

import TransferList from "../Components/TransferList/TransferList";

const RecipeEdit = (props) => {
  const { rows, setOpen, mode, setMode, refreshTable } = props;

  const { ...params } = useParams();

  console.log(params);

  const [material, setMaterial] = useState("");
  const [materialClass, setMaterialClass] = useState("");
  const [RID, setRID] = useState("");
  const [description, setDescription] = useState("");
  const [batchSizeMin, setBSMin] = useState("");
  const [batchSizeMax, setBSMax] = useState("");
  const [batchSizeNom, setBSNom] = useState("");
  const [mainProcessClass, setMainProcessClass] = useState("");
  const [checked, setChecked] = useState([]);
  const [left, setLeft] = useState([]);
  const [right, setRight] = useState([]);
  const [version, setVersion] = useState("");
  const [date, setDate] = useState("");
  const [invalidNames, setInvalidNames] = useState([]);
  const [selected, setSelected] = useState("");
  const [materials, setMaterials] = useState([]);
  const [materialClasses, setMaterialClasses] = useState([]);
  const [processClasses, setProcessClasses] = useState([]);
  const [requiredProcessClasses, setRequiredProcessClasses] = useState([]);

  function initializeView() {
    console.time("Initialize View");
    const getRecipe = () =>
      fetch(`/recipes/${params.RID}/${params.Version}`).then((response) =>
        response.json()
      );
    const getMaterials = () =>
      fetch("/materials").then((response) => response.json());
    const getMaterialClasses = () =>
      fetch("/material/classes").then((response) => response.json());
    const getProcessClasses = () =>
      fetch("/process-classes").then((response) => response.json());
    const getRequiredProcessClasses = () =>
      fetch("/process-classes/required").then((response) => response.json());

    console.timeEnd("Initialize View");

    function getAllData() {
      return Promise.all([
        getRecipe(),
        getMaterials(),
        getMaterialClasses(),
        getProcessClasses(),
        getRequiredProcessClasses(),
      ]);
    }

    getAllData().then(
      ([
        selectedRecipe,
        allMaterials,
        allMaterialClasses,
        allProcessClasses,
        allRequiredProcessClasses,
      ]) => {
        setSelected(selectedRecipe[0]);
        setMaterials(allMaterials);
        setMaterialClasses(allMaterialClasses);
        setProcessClasses(allProcessClasses);
        setRequiredProcessClasses(allRequiredProcessClasses);
      }
    );
  }

  useEffect(() => {
    initializeView();
  }, []);

  const handleClose = () => {
    setRID("");
    setDescription("");
    setMaterial("");
    setMaterialClass("");
    setBSNom("");
    setBSMin("");
    setBSMax("");
    setOpen(false);
    setMode("Search");
    setInvalidNames([]);
  };

  const handleToggle = (value) => () => {
    setChecked([value]);
  };

  const handleButtonSubmit = () => {
    fetch("http://localhost:5000/recipes", {
      method: "POST", // *GET, POST, PUT, DELETE, etc.
      mode: "cors", // no-cors, *cors, same-origin
      cache: "no-cache", // *default, no-cache, reload, force-cache, only-if-cached
      credentials: "same-origin", // include, *same-origin, omit
      headers: {
        "Content-Type": "application/json",
        // 'Content-Type': 'application/x-www-form-urlencoded',
      },
      redirect: "follow", // manual, *follow, error
      referrerPolicy: "no-referrer", // no-referrer, *no-referrer-when-downgrade, origin, origin-when-cross-origin, same-origin, strict-origin, strict-origin-when-cross-origin, unsafe-url
      body: JSON.stringify({
        RID: RID,
        Description: description,
        ProductID: material.SiteMaterialAlias,
        BatchSizeNominal: batchSizeNom,
        BatchSizeMin: batchSizeMin,
        BatchSizeMax: batchSizeMax,
        Status: "Registered",
        Version: mode === "Edit" ? version + 1 : 1,
        VersionDate: new Date(Date.now()),
        RecipeType: "Master",
        UseBatchKernel: 1,
        RunMode: 0,
        IsPackagingRecipeType: 0,
      }), // body data type must match "Content-Type" header
      if(right) {
        console.log("This ran");
      },
    })
      .then((response) => response.json())
      .then((data) => {
        refreshTable();
        setSelected(data);
        setMode("View");
      });
  };

  return (
    <Grid container spacing={2} sx={{ height: "88vh" }}>
      <Grid item xs={12}>
        <Paper sx={{ height: "100%" }}>
          <Box sx={{ display: "flex", justifyContent: "space-between" }}>
            <Typography
              component="h1"
              variant="h5"
              color="inherit"
              noWrap
              sx={{ flexGrow: 1, alignSelf: "center" }}
            >
              Edit Recipe{`- ${selected.RID}`}
            </Typography>
            <Box sx={{}}>
              <Box
                sx={{
                  display: "flex",
                  justifyContent: "space-between",
                  mb: 1,
                }}
              >
                <Typography
                  component="h5"
                  variant="subtitle"
                  color="inherit"
                  noWrap
                  sx={{ alignSelf: "center", pb: 0.25 }}
                >
                  Version: {selected.Version}
                </Typography>
                <Typography
                  component="h5"
                  variant="subtitle"
                  color="inherit"
                  noWrap
                  sx={{ alignSelf: "center" }}
                >
                  Status: {``}
                  <Chip
                    label={selected.Status}
                    color={
                      selected.Status === "Approved"
                        ? "success"
                        : selected.Status === "Valid"
                        ? "info"
                        : selected.Status === "Registered"
                        ? "info"
                        : selected.Status === "Obsolete"
                        ? "error"
                        : "info"
                    }
                    // variant='outlined'
                  />
                </Typography>
              </Box>
              <Typography
                component="h5"
                variant="subtitle"
                color="inherit"
                noWrap
                sx={{ alignSelf: "flex-start" }}
              >
                Date: {new Date(selected.VersionDate).toUTCString()}
              </Typography>
            </Box>
          </Box>
          <Divider sx={{ mb: 2, mt: 2 }} />
          <Box>
            <Box sx={{ display: "flex" }}>
              <Box
                sx={{
                  width: 375,
                  mr: 4,
                  display: "flex",
                  flexDirection: "column",
                  alignItems: "space-around",
                }}
              >
                {mode === "Copy" ? (
                  <Box>
                    <TextField
                      id="RID"
                      label="New Recipe ID"
                      value={RID}
                      error={
                        !RID || rows.map((recipe) => recipe.RID).includes(RID)
                      }
                      type="string"
                      onChange={(event) => {
                        setRID(event.target.value);
                      }}
                      sx={{ mb: 2, width: 350, alignSelf: "center" }}
                    />
                    <TextField
                      id="Description"
                      label="New Description"
                      value={description}
                      type="string"
                      InputProps={{
                        readOnly: mode === "View" ? true : false,
                      }}
                      onChange={(event) => {
                        setDescription(event.target.value);
                      }}
                      sx={{ mb: 2, width: 350, alignSelf: "center" }}
                    />
                  </Box>
                ) : (
                  <Box>
                    <Typography
                      component="h5"
                      variant="h6"
                      color="inherit"
                      noWrap
                      sx={{ mb: 2, alignSelf: "flex-start", ml: 2 }}
                    >
                      Recipe Information
                    </Typography>
                    <TextField
                      id="RID"
                      label="Recipe ID"
                      value={RID}
                      error={!RID}
                      type="string"
                      // helperText={!RID && 'Please enter a valid name'}
                      InputProps={{
                        readOnly: mode === "View" ? true : false,
                      }}
                      onChange={(event) => {
                        setRID(event.target.value);
                      }}
                      sx={{ mb: 2, width: 350, alignSelf: "center" }}
                    />
                    <TextField
                      id="Description"
                      label="Description"
                      value={description}
                      type="string"
                      InputProps={{
                        readOnly: mode === "View" ? true : false,
                      }}
                      onChange={(event) => {
                        setDescription(event.target.value);
                      }}
                      sx={{ mb: 2, width: 350, alignSelf: "center" }}
                    />
                    <FormControl
                      sx={{ mb: 2, width: 350, alignSelf: "center" }}
                    >
                      <InputLabel id="demo-simple-select-helper-label">
                        Material Class
                      </InputLabel>
                      <Select
                        labelId="material-class-select-helper-label"
                        id="material-class-select-helper"
                        value={materialClass}
                        label="Material Class"
                        disabled={mode === "View"}
                        error={!materialClass}
                        onChange={(event) => {
                          setMaterialClass(event.target.value);
                        }}
                      >
                        <MenuItem value="">
                          <em>None</em>
                        </MenuItem>
                        {materialClasses.map((materialClass) => {
                          return (
                            <MenuItem
                              key={materialClass.ID}
                              value={materialClass}
                            >{`${materialClass.Name}`}</MenuItem>
                          );
                        })}
                      </Select>
                      <FormHelperText></FormHelperText>
                    </FormControl>
                    <FormControl
                      sx={{ mb: 2, width: 350, alignSelf: "center" }}
                    >
                      <InputLabel id="demo-simple-select-helper-label">
                        Material
                      </InputLabel>
                      <Select
                        labelId="material-select-helper-label"
                        id="material-select-helper"
                        value={material}
                        label="Material"
                        disabled={mode === "View" || !materialClass}
                        error={!material}
                        onChange={(event) => {
                          setMaterial(event.target.value);
                        }}
                      >
                        <MenuItem value="">
                          <em>None</em>
                        </MenuItem>
                        {materials
                          .filter(
                            (material) =>
                              material.MaterialClass_ID === materialClass.ID
                          )
                          .map((material) => {
                            return (
                              <MenuItem
                                key={material.ID}
                                value={material}
                              >{`${material.SiteMaterialAlias} - ${material.Name}`}</MenuItem>
                            );
                          })}
                      </Select>
                      <FormHelperText></FormHelperText>
                    </FormControl>
                    <Divider sx={{ mb: 2, mt: 2, width: 300 }} />
                    <Typography
                      component="h1"
                      variant="h6"
                      color="inherit"
                      noWrap
                      sx={{ mb: 2, alignSelf: "flex-start", ml: 2 }}
                    >
                      Batch Size
                    </Typography>
                    <Box
                      sx={{
                        display: "flex",
                        justifyContent: "space-around",
                        mb: 2,
                      }}
                    >
                      <TextField
                        id="BatchSizeNominal"
                        label="Nominal"
                        value={batchSizeNom}
                        type="number"
                        InputProps={{
                          readOnly: mode === "View" ? true : false,
                        }}
                        error={
                          batchSizeNom > batchSizeMax ||
                          batchSizeNom < batchSizeMin ||
                          batchSizeNom === ""
                        }
                        onChange={(event) => {
                          setBSNom(event.target.value);
                        }}
                        sx={{ mb: 2, width: 100 }}
                      />
                      <TextField
                        id="BatchSizeMin"
                        label="Min"
                        value={batchSizeMin}
                        type="number"
                        error={
                          batchSizeMin > batchSizeMax || batchSizeMin === ""
                        }
                        InputProps={{
                          readOnly: mode === "View" ? true : false,
                        }}
                        onChange={(event) => {
                          setBSMin(event.target.value);
                        }}
                        sx={{ mb: 2, width: 100 }}
                      />
                      <TextField
                        id="BatchSizeMax"
                        label="Max"
                        value={batchSizeMax}
                        type="number"
                        error={batchSizeMax === 0 || batchSizeMax === ""}
                        InputProps={{
                          readOnly: mode === "View" ? true : false,
                        }}
                        onChange={(event) => {
                          setBSMax(event.target.value);
                        }}
                        sx={{ mb: 2, width: 100 }}
                      />
                    </Box>
                  </Box>
                )}
                <Box>
                  <Box
                    sx={{ display: "flex", justifyContent: "space-between" }}
                  >
                    <Button
                      variant="contained"
                      endIcon={<SaveIcon />}
                      onClick={handleButtonSubmit}
                      sx={{ width: "45%" }}
                    >
                      Save
                    </Button>
                  </Box>
                </Box>
              </Box>
              {mode !== "Copy" && (
                <Box>
                  <Divider orientation="vertical" flexItem sx={{ mr: 2 }} />
                  <Box
                    sx={{
                      display: "flex",
                      flexDirection: "column",
                    }}
                  >
                    <Box>
                      <Typography
                        component="h1"
                        variant="h6"
                        color="inherit"
                        noWrap
                        sx={{ ml: 4, mb: 2.8 }}
                      >
                        Process Class Requirement
                      </Typography>
                      <Box sx={{ mb: 3 }}>
                        <TransferList
                          setChecked={setChecked}
                          setLeft={setLeft}
                          setRight={setRight}
                          handleToggle={handleToggle}
                          processClasses={processClasses}
                          requiredProcessClasses={requiredProcessClasses}
                          left={left}
                          right={right}
                          checked={checked}
                          mode={mode}
                          selected={selected}
                          setMainProcessClass={setMainProcessClass}
                          mainProcessClass={mainProcessClass}
                        />
                      </Box>
                    </Box>
                  </Box>
                </Box>
              )}
            </Box>
          </Box>
        </Paper>
      </Grid>
    </Grid>
  );
};

export default React.memo(RecipeEdit);
