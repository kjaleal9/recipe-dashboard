import React, { useState } from "react";
import { Box, Typography, Divider, Chip, Button } from "@mui/material";

const RecipeView = (props) => {
  const { isProcedureEditable, setIsProcedureEditable,setSteps, selectedRecipe } = props;

  const handleCancelButton = () => {
    setIsProcedureEditable(false);
    setSteps(selectedRecipe);
  };

  const handleEditProcedureButton = () => {
    setIsProcedureEditable(true);
  };

  return (
    <Box sx={{ display: "flex", flexDirection: "column", p: 3 }}>
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
    </Box>
  );
};

export default RecipeView;
