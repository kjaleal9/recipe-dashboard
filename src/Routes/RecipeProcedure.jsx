import React, { useEffect, useState } from "react";
import { DragDropContext, Droppable, Draggable } from "react-beautiful-dnd";

// TODO: Convert to react-beautiful-dnd to dndkit. rbdnd is no longer maintained.

import {
  Box,
  Button,
  ButtonGroup,
  Divider,
  Grid,
  Paper,
  Typography,
  Toolbar,
  Tooltip,
  FormControl,
  InputLabel,
  Select,
  MenuItem,
  FormHelperText,
  selectClasses,
} from "@mui/material";

import ProcedureRow from "../Components/ProcudureRow/ProcedureRow";
import StepView from "../Components/StepView/StepView";
import RecipeProcedureTable from "../Components/Table/CustomTableBody/RecipeProcedureTable";

import AddBoxIcon from "@mui/icons-material/AddBox";
import EditIcon from "@mui/icons-material/Edit";
import DeleteIcon from "@mui/icons-material/Delete";



// 20230224185204
// http://localhost:5000/recipes/latest
const RecipeProcedure = () => {
  // Fetch steps from database
  // Pull Recipe_RID Recipe_Version TPIBK_StepType_ID ProcessClassPhase_ID, Step UserString Material_ID ProcessClass_ID
  // Assign a temporary ID

  // Temporary
  const [recipes, setRecipes] = useState([]);
  const [recipeSelect, setRecipeSelect] = useState("");
  const [versionSelect, setVersionSelect] = useState("");

  const [phases, setPCPhases] = useState([]);
  const [materials, setMaterials] = useState([]);
  const [materialClasses, setMaterialClasses] = useState([]);
  const [selected, setSelected] = useState("");
  const [steps, setSteps] = useState("");
  const [stepTypes, setStepTypes] = useState("");
  const [equipment, setEquipment] = useState([]);

  const [selectedRecipe, setSelectedRecipe] = useState("");
  const [selectedRER, setSelectedRER] = useState("");
  const [isProcedureEditable, setIsProcedureEditable] = useState(false);

  // Move to route loader
  const loadData = () => {
    const getRecipes = () =>
      fetch("/recipes").then((response) => response.json());
    const getPCPhases = () =>
      fetch("/process-classes/phases").then((response) => response.json());
    const getStepTypes = () =>
      fetch("/recipes/step-types").then((response) => response.json());
    const getMaterials = () =>
      fetch("/materials").then((response) => response.json());
    const getMaterialClasses = () =>
      fetch("/materials/classes").then((response) => response.json());
    const getEquipment = () =>
      fetch("/equipment").then((response) => response.json());

    function getAllData() {
      return Promise.all([
        getRecipes(),
        getPCPhases(),
        getStepTypes(),
        getMaterials(),
        getMaterialClasses(),
        getEquipment(),
      ]);
    }

    getAllData()
      .then(
        ([
          recipes,
          pcPhases,
          stepTypes,
          allMaterials,
          allMaterialClasses,
          requiredEquipment,
        ]) => {
          setRecipes(recipes);
          setPCPhases(pcPhases);
          setStepTypes(stepTypes);
          setMaterials(allMaterials);
          setMaterialClasses(allMaterialClasses);
          setEquipment(requiredEquipment);
        }
      )
      .catch((err) => console.log(err, "uih"));
  };

  useEffect(() => {
    loadData();
  }, []);

  useEffect(() => {
    setVersionSelect("");
  }, [recipeSelect]);

  useEffect(() => {
    setSteps(selectedRecipe);
  }, [selectedRecipe]);

  const handleNewStep = () => {};
  const handleDeleteStep = () => {};
  const handleEditStep = () => {};

  const procedureSearchButton = () => {
    fetch(`/recipes/procedure/${recipeSelect}/${versionSelect}`).then(
      (response) =>
        response.json().then((data) => {
          setSelectedRecipe(data);
        })
    );
    fetch(`/process-classes/required/${recipeSelect}/${versionSelect}`)
      .then((response) => response.json())
      .then((data) => setSelectedRER(data));
  };

  const handleClick = (event, step, row) => {
    const selectedRow = `${selected.ID}-${selected.Message}` === step;
    let newSelected = "";

    if (selectedRow === false) {
      newSelected = row;
    }

    setSelected(newSelected);
  };

  const onDragEnd = (result) => {
    if (!result.destination) {
      return;
    }
    const newItems = Array.from(steps);
    const [removed] = newItems.splice(result.source.index, 1);
    newItems.splice(result.destination.index, 0, removed);
    setSteps(newItems);
  };

  function isSelected(step) {
    return selected.Step === step;
  }

  const handleCancelButton = () => {
    setIsProcedureEditable(false);
    setSteps(selectedRecipe);
  };

  const handleEditProcedureButton = () => {
    setIsProcedureEditable(true);
  };

  return (
    <Grid container spacing={2}>
      <Grid item xs={12} md={6} lg={4}>
        <Paper>
          <Box sx={{ height: "88vh" }}>
            <Toolbar
              sx={{
                pl: { sm: 2 },
                pr: { xs: 1, sm: 1 },
                display: "flex",
                justifyContent: "space-between",
              }}
            >
              <Typography
                component="h1"
                variant="h5"
                color="inherit"
                noWrap
                sx={{ flexGrow: 1, alignSelf: "center" }}
              >
                Recipe Procedure
              </Typography>
              <Box
                sx={{
                  display: "flex",
                  flexDirection: "column",
                  justifyContent: "space-around",
                }}
              >
                {isProcedureEditable ? (
                  <Button
                    variant="contained"
                    alignSelf="center"
                    sx={{ m: 1, alignSelf: "center", width: "75%" }}
                    onClick={handleCancelButton}
                  >
                    Cancel
                  </Button>
                ) : (
                  <Button
                    variant="contained"
                    alignSelf="center"
                    sx={{ m: 1, alignSelf: "center", width: "75%" }}
                    onClick={handleEditProcedureButton}
                  >
                    Edit Procedure
                  </Button>
                )}
              </Box>
              {isProcedureEditable && (
                <Box sx={{ display: "flex", justifySelf: "flex-end" }}>
                  <ButtonGroup
                    color="primary"
                    sx={{ justifySelf: "flex-end", mx: 2, height: 50 }}
                  >
                    <Tooltip title="New">
                      <Box>
                        <Button
                          variant="contained"
                          onClick={handleNewStep}
                          sx={{ height: "100%" }}
                        >
                          <AddBoxIcon />
                        </Button>
                      </Box>
                    </Tooltip>
                    <Tooltip title="Edit">
                      <Box>
                        <Button
                          variant="contained"
                          onClick={handleEditStep}
                          sx={{ height: "100%" }}
                        >
                          <EditIcon />
                        </Button>
                      </Box>
                    </Tooltip>
                    <Tooltip title="Delete">
                      <Box>
                        <Button
                          variant="contained"
                          disabled={selected === ""}
                          onClick={handleDeleteStep}
                          sx={{ height: "100%" }}
                        >
                          <DeleteIcon />
                        </Button>
                      </Box>
                    </Tooltip>
                  </ButtonGroup>
                </Box>
              )}
            </Toolbar>
            <Divider />
            {steps && (
              <RecipeProcedureTable
                rows={steps}
                onDragEnd={onDragEnd}
                handleClick={handleClick}
                isSelected={isSelected}
                isProcedureEditable={isProcedureEditable}
              />
            )}
          </Box>
        </Paper>
      </Grid>
      <Grid item xs={12} md={6} lg={6}>
        <Paper>
          <Box height={"88vh"}>
            {selectedRecipe && (
              <StepView
                isProcedureEditable={isProcedureEditable}
                setIsProcedureEditable={setIsProcedureEditable}
                setSteps={setSteps}
                selectedRecipe={selectedRecipe}
                selectedRER={selectedRER}
                selectedStep={selected}
                equipment={equipment}
                phases={phases}
              />
            )}
          </Box>
        </Paper>
      </Grid>
      <Grid item xs={12} md={4} lg={2}>
        <Paper>
          <Box
            p={2}
            display="flex"
            flexDirection={"column"}
            justifyContent={"space-around"}
            gap={"6px"}
          >
            <Typography component="h1" variant="h6" color="inherit" noWrap>
              Change Recipe
            </Typography>

            <FormControl width={"100%"}>
              <InputLabel id="demo-simple-select-helper-label">
                Recipe
              </InputLabel>
              <Select
                labelId="recipe-select-helper-label"
                id="recipe-select-helper"
                value={recipeSelect}
                label="Recipe"
                onChange={(event) => {
                  setRecipeSelect(event.target.value);
                }}
                MenuProps={{ PaperProps: { style: { maxHeight: 400 } } }}
              >
                <MenuItem value="">
                  <em>None</em>
                </MenuItem>

                {[...new Set(recipes.map((recipe) => recipe.RID))].map(
                  (name) => {
                    return (
                      <MenuItem key={name} value={name}>{`${name}`}</MenuItem>
                    );
                  }
                )}
              </Select>
              <FormHelperText></FormHelperText>
            </FormControl>

            <FormControl width={"50%"}>
              <InputLabel id="demo-simple-select-helper-label">
                Version
              </InputLabel>
              <Select
                labelId="version-select-helper-label"
                id="version-select-helper"
                value={versionSelect}
                label="Recipe"
                disabled={!recipeSelect}
                onChange={(event) => {
                  setVersionSelect(event.target.value);
                }}
                MenuProps={{ PaperProps: { style: { maxHeight: 400 } } }}
              >
                <MenuItem value="">
                  <em>None</em>
                </MenuItem>

                {[
                  ...recipes
                    .filter((recipe) => recipe.RID === recipeSelect)
                    .map((recipe) => recipe.Version)
                    .sort((a, b) => a - b),
                ].map((version) => {
                  return (
                    <MenuItem
                      key={version}
                      value={version}
                    >{`${version}`}</MenuItem>
                  );
                })}
              </Select>
              <FormHelperText></FormHelperText>
            </FormControl>
            <Button
              variant="contained"
              sx={{ width: "100%" }}
              onClick={procedureSearchButton}
            >
              Confirm
            </Button>
          </Box>
        </Paper>
      </Grid>
    </Grid>
  );
};

export default RecipeProcedure;
