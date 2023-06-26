import React, { useState, useEffect } from "react";
import { Box, Typography, Divider, Chip, Button } from "@mui/material";

const StepView = (props) => {
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

  const [parameters, setParameters] = useState([]);

  console.log(selectedStep);

  useEffect(() => {
    console.time("get parameters");
    setParameters([]);
    if ([1, 2, 3, 4, 5].includes(selectedStep.TPIBK_StepType_ID)) {
      fetch(
        `/parameters/${selectedStep.ID}/${selectedStep.ProcessClassPhase_ID}`
      )
        .then((response) => response.json())
        .then((data) => {
          setParameters(data.recordset);
          console.log(data.recordset);
        });
    }
    console.timeEnd("get parameters");
  }, [selectedStep]);

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
          <Box>
            {
              testJSX.find((item) => item.ID === selectedStep.TPIBK_StepType_ID)
                .jsxObject
            }
          </Box>
          <Box>
            {parameters &&
              parameters.map((parameter) => (
                <Typography>
                  {parameter.Name} - {parameter.Value}
                </Typography>
              ))}
          </Box>
        </Box>
      )}
    </Box>
  );
};

export default StepView;
