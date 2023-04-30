import React, { useState } from "react";
import { Box, Typography, Divider, Chip, Button } from "@mui/material";

import ProcessClassModal from "../Modals/ProcessClassModal/ProcessClassModal";
import ProcedureModal from "../Modals/ProcedureModal/ProcedureModal";

const RecipeView = (props) => {
  const { selected, setMode, setOpen } = props;
  const [openProcessClassModal, setOpenProcessClassModal] = useState(false);
  const [openProcedureModal, setOpenProcedureModal] = useState(false);

  const handleButtonProcessClasses = () => {
    setOpenProcessClassModal(true);
  };
  const handleButtonProcedure = () => {
    setOpenProcedureModal(true);
  };

  return (
    <Box sx={{ display: "flex", flexDirection: "column", p: 3 }}>
      <ProcessClassModal
        selected={selected}
        recipe={selected.RID}
        version={selected.Version}
        open={openProcessClassModal}
        setOpen={setOpen}
        setMode={setMode}
        setOpenProcessClassModal={setOpenProcessClassModal}
      />
      <ProcedureModal
        selected={selected}
        recipe={selected.RID}
        version={selected.Version}
        open={openProcedureModal}
        setOpen={setOpen}
        setMode={setMode}
        setOpenProcedureModal={setOpenProcedureModal}
      />
      <Typography component="h1" variant="h5" align="center">
        {selected.RID}
      </Typography>

      <Box
        flexGrow={1}
        display="flex"
        flexDirection="column"
        alignItems="center"
      >
        <Typography
          component="p"
          variant="p"
          sx={{ alignSelf: "center", pb: 0.25 }}
        >
          Finished Product: {selected.Name}
        </Typography>
        <Typography
          component="p"
          variant="p"
          sx={{ alignSelf: "center", pb: 0.25 }}
        >
          ID: {selected.ProductID}
        </Typography>
      </Box>
      <Typography component="p" variant="p" align="center">
        {selected.Description}
      </Typography>
      <Divider sx={{ m: 2 }} light />
      <Box sx={{ display: "flex", justifyContent: "space-around" }}>
        <Box display="flex" flexDirection="column" alignItems={"flex-start"}>
          <Typography component="p" variant="p">
            Version: {selected.Version}
          </Typography>
          <Typography component="p" variant="p" noWrap>
            Status: {``}
            <Chip
              label={selected.Status}
              size="small"
              color={
                selected.Status === "Approved"
                  ? "success"
                  : selected.Status === "Valid"
                  ? "info"
                  : selected.Status === "Registered"
                  ? "warning"
                  : selected.Status === "Obsolete"
                  ? "error"
                  : "info"
              }
            />
          </Typography>

          <Typography component="p" variant="p" color="inherit" noWrap>
            Date:{" "}
            {new Date(
              (selected.VersionDate - 25569) * 86400 * 1000
            ).toLocaleDateString()}
          </Typography>
        </Box>
        <Box display="flex" flexDirection="column" alignItems={"flex-start"}>
          <Typography component="p" variant="p" color="inherit" noWrap>
            Max: {selected.BatchSizeMax} gal
          </Typography>
          <Typography component="p" variant="p" color="inherit" noWrap>
            Min: {selected.BatchSizeMin} gal
          </Typography>
          <Typography component="p" variant="p" color="inherit" noWrap>
            Nominal: {selected.BatchSizeMin} gal
          </Typography>
        </Box>
      </Box>

      <Divider sx={{ m: 2 }} light />
      <Box
        sx={{
          display: "flex",
          flexDirection: "column",
          justifyContent: "space-around",
        }}
      >
        <Button
          variant="contained"
          alignSelf="center"
          sx={{ m: 1, alignSelf: "center", width: "75%" }}
        >
          BOM
        </Button>
        <Button
          variant="contained"
          alignSelf="center"
          sx={{ m: 1, alignSelf: "center", width: "75%" }}
          onClick={handleButtonProcessClasses}
        >
          Process Classes
        </Button>
        <Button
          variant="contained"
          alignSelf="center"
          sx={{ m: 1, alignSelf: "center", width: "75%" }}
          onClick={handleButtonProcedure}
        >
          Procedure
        </Button>
      </Box>
      <Divider sx={{ m: 2 }} light />
      <Box display="flex" flexDirection="column">
        <Typography
          component="h5"
          variant="h5"
          color="inherit"
          sx={{ alignSelf: "center", pb: 0.25 }}
        >
          User Comment
        </Typography>
        <Typography
          component="p"
          variant="p"
          color="inherit"
          sx={{ alignSelf: "center", pb: 0.25 }}
        >
          Lorem ipsum dolor sit amet consectetur, adipisicing elit. Excepturi,
          quod optio doloribus exercitationem libero aut quis laborum illum ipsa
          provident, aperiam doloremque ex esse deleniti temporibus harum
          voluptatibus. Deserunt minima laboriosam aperiam, ipsum est
          perferendis suscipit! Vel, iusto modi! Consequuntur ratione ex
          voluptates sunt dolore veniam nulla modi qui ea eligendi libero fuga,
          minima nam? Officiis cum nihil, laborum facilis dignissimos culpa
          natus consequuntur sapiente nobis, ratione quisquam minima minus
          delectus similique praesentium aperiam obcaecati eligendi accusantium
          officia harum veritatis dolorum at odit neque?
        </Typography>
      </Box>
    </Box>
  );
};

export default RecipeView;
