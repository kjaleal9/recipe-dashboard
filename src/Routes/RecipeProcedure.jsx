import React, { Fragment } from "react";
import {
  Box,
  Divider,
  FormControl,
  FormHelperText,
  InputLabel,
  List,
  ListItem,
  ListItemButton,
  ListItemIcon,
  ListItemText,
  MenuItem,
  Select,
  Stack,
  Paper,
  Typography,
  TextField,
} from "@mui/material";

import ProcedureRow from "../Components/ProcudureRow/ProcedureRow";

// 20230224185204
// http://localhost:5000/recipes/latest

const procedure = {
  steps: [
    { stepName: "T151 Allocate", stepType: "Allocate" },
    { stepName: "T151 Agitation", stepType: "Run" },
    { stepName: "R100 Allocate", stepType: "Allocate" },
    { stepName: "R100 Empty", stepType: "Start" },
    { stepName: "T151 Fill", stepType: "Start" },
    { stepName: "R100 Empty", stepType: "Stop" },
    { stepName: "T151 Fill", stepType: "Stop" },
    { stepName: "R100 Deallocate", stepType: "Deallocate" },
    { stepName: "T151 Deallocate", stepType: "Deallocate" },
    { stepName: "T151 Allocate", stepType: "Allocate" },
    { stepName: "T151 Agitation", stepType: "Run" },
    { stepName: "R100 Allocate", stepType: "Allocate" },
    { stepName: "R100 Empty", stepType: "Start" },
    { stepName: "T151 Fill", stepType: "Start" },
    { stepName: "R100 Empty", stepType: "Stop" },
    { stepName: "T151 Fill", stepType: "Stop" },
    { stepName: "R100 Deallocate", stepType: "Deallocate" },
    { stepName: "T151 Deallocate", stepType: "Deallocate" },
    { stepName: "T151 Allocate", stepType: "Allocate" },
    { stepName: "T151 Agitation", stepType: "Run" },
    { stepName: "R100 Allocate", stepType: "Allocate" },
    { stepName: "R100 Empty", stepType: "Start" },
    { stepName: "T151 Fill", stepType: "Start" },
    { stepName: "R100 Empty", stepType: "Stop" },
    { stepName: "T151 Fill", stepType: "Stop" },
    { stepName: "R100 Deallocate", stepType: "Deallocate" },
    { stepName: "T151 Deallocate", stepType: "Deallocate" },
    { stepName: "T151 Allocate", stepType: "Allocate" },
    { stepName: "T151 Agitation", stepType: "Run" },
    { stepName: "R100 Allocate", stepType: "Allocate" },
    { stepName: "R100 Empty", stepType: "Start" },
    { stepName: "T151 Fill", stepType: "Start" },
    { stepName: "R100 Empty", stepType: "Stop" },
    { stepName: "T151 Fill", stepType: "Stop" },
    { stepName: "R100 Deallocate", stepType: "Deallocate" },
    { stepName: "T151 Deallocate", stepType: "Deallocate" },
  ],
};

const RecipeProcedure = (props) => {
  const { materials, materialClasses } = props;
  return (
    <Box
      sx={{
        display: "flex",
        flexDirection: "column",
        alignItems: "space-between",
        height: "90vh",
      }}
    >
      {" "}
      <Typography
        component="h3"
        variant="h6"
        color="inherit"
        noWrap
        sx={{ m: 2 }}
      >
        Procedure
      </Typography>
      <Box sx={{ display: "flex", flexDirection: "column", height: "80%" }}>
        {/* <Divider /> */}
        <Box sx={{ overflowY: "scroll" }}>
          {procedure.steps.map((step, index) => (
            <ProcedureRow step={step} index={index} />
          ))}
        </Box>
      </Box>
      {/* <Divider orientation='vertical' variant='middle' flexItem /> */}
      <Box Box sx={{ p: 1, m: 2 }}>
        Test
        <FormControl sx={{ mb: 2, width: 350, alignSelf: "center" }}>
          <InputLabel id="demo-simple-select-helper-label">
            Material Class
          </InputLabel>
          <Select
            labelId="material-class-select-helper-label"
            id="material-class-select-helper"
            label="Material Class"
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
        <FormControl sx={{ mb: 2, width: 350, alignSelf: "center" }}>
          <InputLabel id="demo-simple-select-helper-label">Material</InputLabel>
          <Select
            labelId="material-select-helper-label"
            id="material-select-helper"
            label="Material"
          >
            <MenuItem value="">
              <em>None</em>
            </MenuItem>
            {materials.map((material) => {
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
      </Box>
    </Box>
  );
};

export default RecipeProcedure;
