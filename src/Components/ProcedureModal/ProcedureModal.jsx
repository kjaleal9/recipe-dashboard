import React, { useState, useEffect } from "react";
import { Modal, Box, Typography } from "@mui/material";

const ProcedureModal = (props) => {
  const { recipe, version, open, setOpenProcedureModal, selected } = props;

  const [procedure, setProcedure] = useState([]);

  const getProcedure = () =>
    fetch(`/recipes/procedure/${recipe}/${version}`).then((response) =>
      response.json()
    );

  useEffect(() => {
    getProcedure().then((data) => {
      setProcedure(data);
      console.log(data);
    });
  }, [selected]);

  const style = {
    position: "absolute",
    top: "50%",
    left: "50%",
    transform: "translate(-50%, -50%)",
    width: 500,
    bgcolor: "background.paper",
    border: "2px solid #000",
    borderRadius: "20px",
    boxShadow: 24,
    p: 4,
    display: "flex",
    flexDirection: "column",
    height: 500,
  };

  const handleClose = () => {
    setOpenProcedureModal(false);
  };

  return (
    <Modal
      open={open}
      onClose={handleClose}
      aria-labelledby="modal-modal-title"
      aria-describedby="modal-modal-description"
    >
      <Box sx={style}>
        <Box
          sx={{
            display: "flex",
            flexDirection: "column",
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
            Procedure
          </Typography>

          <Box sx={{ overflowY: "scroll", height: 400 }}>
            {procedure.map((step) => (
              <Typography>
                {step.Step} - {step.Message}
              </Typography>
            ))}
          </Box>
        </Box>
      </Box>
    </Modal>
  );
};

export default React.memo(ProcedureModal);
