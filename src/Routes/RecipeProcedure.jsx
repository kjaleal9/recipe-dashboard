import React, { Fragment, useEffect, useState } from "react";
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
} from "@mui/material";

import ProcedureRow from "../Components/ProcudureRow/ProcedureRow";
import StepView from "../Components/StepView/StepView";

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
  const [recipeNames, setRecipeNames] = useState([]);
  const [versionSelect, setVersionSelect] = useState("");

  const [phases, setPhases] = useState([]);
  const [materials, setMaterials] = useState([]);
  const [materialClasses, setMaterialClasses] = useState([]);
  const [selected, setSelected] = useState("");
  const [steps, setSteps] = useState("");
  const [selectedRecipe, setSelectedRecipe] = useState("");
  const [isProcedureEditable, setIsProcedureEditable] = useState(false);

  function handleOnDragEnd(result) {
    if (!result.destination) return;

    const items = Array.from(steps);
    const [reorderedItem] = items.splice(result.source.index, 1);
    items.splice(result.destination.index, 0, reorderedItem);

    setSteps(items);
  }

  // Move to route loader
  const refreshMaterials = () => {
    const getRecipes = () =>
      fetch("/recipes").then((response) => response.json());
    const getPhases = () =>
      fetch("/phases").then((response) => response.json());
    const getMaterials = () =>
      fetch("/materials").then((response) => response.json());
    const getMaterialClasses = () =>
      fetch("/materials/classes").then((response) => response.json());

    function getAllData() {
      return Promise.all([getRecipes(), getMaterials(), getMaterialClasses()]);
    }

    getAllData()
      .then(([recipes, phases, allMaterials, allMaterialClasses]) => {
        setRecipes(recipes);
        setPhases(phases);
        setMaterials(allMaterials);
        setMaterialClasses(allMaterialClasses);
      })
      .catch((err) => console.log(err, "uih"));
  };

  useEffect(() => {
    refreshMaterials();
  }, []);

  useEffect(() => {
    setVersionSelect("");
  }, [recipeSelect]);

  useEffect(() => {
    setSteps(selectedRecipe);
    console.log(selectedRecipe);
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
  };

  const findPCP = (phaseID) => {
    if (phaseID === "NULL") {
      return "NULL";
    } else {
      const phase = selectedRecipe.processClassPhases.find(
        (pcp) => pcp.ID === phaseID
      );
      return phase;
    }
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
                          disabled={selected !== ""}
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
              <Box sx={{ height: "80vh" }}>
                <DragDropContext onDragEnd={handleOnDragEnd}>
                  <Droppable
                    droppableId="characters"
                    isDropDisabled={!isProcedureEditable}
                  >
                    {(provided) => (
                      <div
                        className="characters"
                        {...provided.droppableProps}
                        ref={provided.innerRef}
                        style={{ overflowY: "scroll", maxHeight: "100%" }}
                      >
                        {steps.map((step, index) => {
                          return (
                            <Draggable
                              key={step.ID}
                              draggableId={step.ID.toString()}
                              index={index}
                              isDragDisabled={!isProcedureEditable}
                            >
                              {(provided) => (
                                <Box
                                  ref={provided.innerRef}
                                  {...provided.draggableProps}
                                  {...provided.dragHandleProps}
                                >
                                  <ProcedureRow step={step} index={index} />
                                </Box>
                              )}
                            </Draggable>
                          );
                        })}
                        {provided.placeholder}
                      </div>
                    )}
                  </Droppable>
                </DragDropContext>
              </Box>
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
