import React, { Fragment, useEffect, useState } from 'react';
import { DragDropContext, Droppable, Draggable } from 'react-beautiful-dnd';

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
} from '@mui/material';

import ProcedureRow from '../Components/ProcudureRow/ProcedureRow';

import AddBoxIcon from '@mui/icons-material/AddBox';
import DeleteIcon from '@mui/icons-material/Delete';

// 20230224185204
// http://localhost:5000/recipes/latest
const RecipeProcedure = () => {
  const procedure = {
    steps: [
      { stepName: 'T151 Allocate', stepType: 'Allocate' },
      { stepName: 'T151 Agitation', stepType: 'Run' },
      { stepName: 'R100 Allocate', stepType: 'Allocate' },
      { stepName: 'R100 Empty', stepType: 'Start' },
      { stepName: 'T151 Fill', stepType: 'Start' },
    ],
  };

  const localDatabase = [
    {
      ID: '783',
      Recipe_RID: 'BT_FullRec_Test',
      Recipe_Version: '2',
      TPIBK_StepType_ID: '8',
      ProcessClassPhase_ID: null,
      Step: '5',
      UserString: null,
      RecipeEquipmentTransition_Data_ID: null,
      NextStep: '0',
      Allocation_Type_ID: '0',
      LateBinding: null,
      Material_ID: null,
      ProcessClass_ID: '377',
    },
    {
      ID: '784',
      Recipe_RID: 'BT_FullRec_Test',
      Recipe_Version: '2',
      TPIBK_StepType_ID: '2',
      ProcessClassPhase_ID: '26',
      Step: '10',
      UserString: null,
      RecipeEquipmentTransition_Data_ID: null,
      NextStep: '0',
      Allocation_Type_ID: '0',
      LateBinding: null,
      Material_ID: null,
      ProcessClass_ID: '377',
    },
    {
      ID: '785',
      Recipe_RID: 'BT_FullRec_Test',
      Recipe_Version: '2',
      TPIBK_StepType_ID: '5',
      ProcessClassPhase_ID: '28',
      Step: '15',
      UserString: null,
      RecipeEquipmentTransition_Data_ID: null,
      NextStep: '0',
      Allocation_Type_ID: '0',
      LateBinding: null,
      Material_ID: null,
      ProcessClass_ID: '377',
    },
    {
      ID: '786',
      Recipe_RID: 'BT_FullRec_Test',
      Recipe_Version: '2',
      TPIBK_StepType_ID: '8',
      ProcessClassPhase_ID: '20',
      Step: null,
      UserString: null,
      RecipeEquipmentTransition_Data_ID: null,
      NextStep: '0',
      Allocation_Type_ID: '0',
      LateBinding: null,
      Material_ID: null,
      ProcessClass_ID: '388',
    },
    {
      ID: '787',
      Recipe_RID: 'BT_FullRec_Test',
      Recipe_Version: '2',
      TPIBK_StepType_ID: '2',
      ProcessClassPhase_ID: '68',
      Step: '25',
      UserString: null,
      RecipeEquipmentTransition_Data_ID: null,
      NextStep: '0',
      Allocation_Type_ID: '52',
      LateBinding: null,
      Material_ID: null,
      ProcessClass_ID: '388',
    },
    {
      ID: '788',
      Recipe_RID: 'BT_FullRec_Test',
      Recipe_Version: '2',
      TPIBK_StepType_ID: '8',
      ProcessClassPhase_ID: '35',
      Step: null,
      UserString: null,
      RecipeEquipmentTransition_Data_ID: null,
      NextStep: '0',
      Allocation_Type_ID: '0',
      LateBinding: null,
      Material_ID: null,
      ProcessClass_ID: '386',
    },
  ];

  // Fetch steps from database
  // Pull Recipe_RID Recipe_Version TPIBK_StepType_ID ProcessClassPhase_ID, Step UserString Material_ID ProcessClass_ID
  // Assign a temporary ID 

  const [materials, setMaterials] = useState([]);
  const [materialClasses, setMaterialClasses] = useState([]);
  const [selected, setSelected] = useState('');
  const [steps, setSteps] = useState(procedure.steps);

  function handleOnDragEnd(result) {
    if (!result.destination) return;

    const items = Array.from(steps);
    const [reorderedItem] = items.splice(result.source.index, 1);
    items.splice(result.destination.index, 0, reorderedItem);

    setSteps(items);
  }

  const refreshMaterials = () => {
    const getMaterials = () =>
      fetch('/materials').then(response => response.json());
    const getMaterialClasses = () =>
      fetch('/materials/classes').then(response => response.json());

    function getAllData() {
      return Promise.all([getMaterials(), getMaterialClasses()]);
    }

    getAllData()
      .then(([allMaterials, allMaterialClasses]) => {
        setMaterials(allMaterials);
        setMaterialClasses(allMaterialClasses);
      })
      .catch(err => console.log(err, 'uih'));
  };
  useEffect(() => {
    refreshMaterials();
  }, []);

  const handleNewStep = () => {};
  const handleDeleteStep = () => {};

  return (
    <Grid container spacing={2} sx={{ height: '88vh' }}>
      <Grid item xs={12} md={4} lg={3}>
        <Paper sx={{ height: '100%' }}>
          <Box>
            <Toolbar sx={{ display: 'flex', justifyContent: 'space-between' }}>
              <Typography
                component='h1'
                variant='h5'
                color='inherit'
                noWrap
                sx={{ flexGrow: 1, alignSelf: 'center' }}
              >
                Recipe Procedure
              </Typography>

              <Box sx={{ display: 'flex', justifySelf: 'flex-end' }}>
                <ButtonGroup color='primary' sx={{ justifySelf: 'flex-end' }}>
                  <Tooltip title='New'>
                    <Box>
                      <Button
                        variant='contained'
                        onClick={handleNewStep}
                        sx={{ height: '100%' }}
                      >
                        <AddBoxIcon />
                      </Button>
                    </Box>
                  </Tooltip>
                  <Tooltip title='Delete'>
                    <Box>
                      <Button
                        variant='contained'
                        disabled={selected !== ''}
                        onClick={handleDeleteStep}
                        sx={{ height: '100%' }}
                      >
                        <DeleteIcon />
                      </Button>
                    </Box>
                  </Tooltip>
                </ButtonGroup>
              </Box>
            </Toolbar>
            <Divider />
            <DragDropContext onDragEnd={handleOnDragEnd}>
              <Droppable droppableId='characters'>
                {provided => (
                  <Box
                    className='characters'
                    {...provided.droppableProps}
                    ref={provided.innerRef}
                  >
                    {steps.map((step, index) => {
                      return (
                        <Draggable
                          key={step.ID}
                          draggableId={step.stepName}
                          index={index}
                        >
                          {provided => (
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
                  </Box>
                )}
              </Droppable>
            </DragDropContext>
          </Box>
        </Paper>
      </Grid>
      <Grid item xs={12} md={4} lg={2}>
        <Paper sx={{ height: '100%' }}></Paper>
      </Grid>
    </Grid>
  );
};

export default RecipeProcedure;
