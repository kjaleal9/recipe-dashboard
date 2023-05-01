import React, { useState } from "react";
import {
  Box,
  Typography,
  Divider,
  Chip,
  Button,
  InputLabel,
} from "@mui/material";

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

      <Typography
        component="h1"
        variant="h6"
        sx={{ mb: 1, alignSelf: "center" }}
      >
        {selected.RID.toUpperCase()}
      </Typography>

      <Box display="flex" justifyContent={"space-between"} sx={{ mb: 1 }}>
        <Typography component="p" variant="overline">
          Version {selected.Version}
        </Typography>

        <Typography
          component="p"
          variant="overline"
          color="inherit"
          noWrap
          mr={1}
        >
          {new Date(selected.VersionDate).toLocaleDateString()}
        </Typography>

        <Chip
          label={selected.Status}
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
          sx={{ width: "auto" }}
        />
      </Box>

      <Box sx={{ display: "flex", justifyContent: "space-around", mb: 2 }}>
        <Box display="flex" width="100%" justifyContent={"space-around"}>
          <Box>
            <Typography
              variant="overline"
              style={{ borderBottom: "1px solid #bdbdbd" }}
            >
              Max
            </Typography>
            <Typography component="p" variant="button" color="inherit" noWrap>
              {selected.BatchSizeMax} gal
            </Typography>
          </Box>
          <Divider orientation="vertical" variant="middle" light />
          <Box>
            {" "}
            <Typography
              variant="overline"
              style={{ borderBottom: "1px solid #bdbdbd" }}
            >
              Min
            </Typography>
            <Typography component="p" variant="button" color="inherit" noWrap>
              {selected.BatchSizeMin} gal
            </Typography>
          </Box>
          <Divider orientation="vertical" variant="middle" light />
          <Box>
            <Typography
              variant="overline"
              style={{ borderBottom: "1px solid #bdbdbd" }}
            >
              Nominal
            </Typography>
            <Typography component="p" variant="button" color="inherit" noWrap>
              {selected.BatchSizeMin} gal
            </Typography>
          </Box>
        </Box>
      </Box>

      <Box
        sx={{
          display: "flex",
          flexDirection: "column",
          justifyContent: "space-around",
          alignItems: "center",
        }}
      >
        <Button variant="contained" sx={{ m: 1, width: "75%" }}>
          BOM
        </Button>
        <Button
          variant="contained"
          sx={{ m: 1, width: "75%" }}
          onClick={handleButtonProcessClasses}
        >
          Process Classes
        </Button>
        <Button
          variant="contained"
          sx={{ m: 1, width: "75%" }}
          onClick={handleButtonProcedure}
        >
          Procedure
        </Button>
      </Box>
      <Divider sx={{ m: 2 }} light />
      <Box display="flex" flexDirection="column">
        <Typography component="h1" variant="h6" sx={{ mb: 1, alignSelf:'center' }}>
          USER COMMENT
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
