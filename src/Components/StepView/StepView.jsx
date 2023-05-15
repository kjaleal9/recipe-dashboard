import React, { useState } from "react";
import { Box, Typography, Divider, Chip, Button } from "@mui/material";

const RecipeView = (props) => {
  const {
    isProcedureEditable,
    setIsProcedureEditable,
    setSteps,
    selectedRecipe,
    selectedRER,
    equipment,
    phases,
    selectedStep,
  } = props;

  console.log(selectedStep);

  const testJSX = [
    { ID: 1, jsxObject: <Typography>This is a test for: ID 1</Typography> },
    { ID: 2, jsxObject: <Typography>This is a test for: ID 2</Typography> },
    { ID: 3, jsxObject: <Typography>This is a test for: ID 3</Typography> },
    { ID: 4, jsxObject: <Typography>This is a test for: ID 4</Typography> },
    { ID: 5, jsxObject: <Typography>This is a test for: ID 5</Typography> },
    { ID: 6, jsxObject: <Typography>This is a test for: ID 6</Typography> },
    { ID: 7, jsxObject: <Typography>This is a test for: ID 7</Typography> },
    { ID: 8, jsxObject: <Typography>This is a test for: ID 8</Typography> },
    { ID: 9, jsxObject: <Typography>This is a test for: ID 9</Typography> },
    { ID: 10, jsxObject: <Typography>This is a test for: ID 10</Typography> },
  ];

  return (
    <Box sx={{ display: "flex", flexDirection: "column", p: 3 }}>
      {selectedStep && (
        <Box>
          {
            testJSX.find((item) => item.ID === selectedStep.TPIBK_StepType_ID)
              .jsxObject
          }
        </Box>
      )}
    </Box>
  );
};

export default RecipeView;
