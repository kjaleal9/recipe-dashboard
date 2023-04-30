import React, { useState, useEffect } from "react";
import {
  Modal,
  Box,
  Typography,
  Table,
  TableHead,
  TableCell,
  TableRow,
  TableBody,
  Tooltip,
  Button,
} from "@mui/material";

import EditIcon from "@mui/icons-material/Edit";

const ProcedureModal = (props) => {
  const {
    recipe,
    version,
    open,
    setOpenProcedureModal,
    selected,
    setOpen,
    setMode,
  } = props;

  const [procedure, setProcedure] = useState([]);

  const getProcedure = () =>
    fetch(`/recipes/procedure/${recipe}/${version}`).then((response) =>
      response.json()
    );

  useEffect(() => {
    getProcedure().then((data) => {
      setProcedure(data);
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
    height: 600,
  };

  const handleClose = () => {
    setOpenProcedureModal(false);
  };
  const handleEdit = (mode) => {
    setMode(mode);
    handleClose();
    setOpen(true);
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
          <Box display={"flex"} sx={{ mb: 2 }}>
            <Typography
              component="h1"
              variant="h5"
              color="inherit"
              noWrap
              sx={{ flexGrow: 1 }}
            >
              Procedure
            </Typography>
            <Tooltip title="Edit">
              <Box>
                <Button
                  variant="contained"
                  onClick={() => handleEdit("View")}
                  sx={{ height: "100%" }}
                >
                  <EditIcon />
                </Button>
              </Box>
            </Tooltip>
          </Box>
          {/* 
          <Box sx={{ overflowY: "scroll", height: 500 }}>
            {procedure.map((step) => (
              <Typography>
                {step.Step} - {step.Message}
              </Typography>
            ))}
          </Box> */}
          <Box sx={{ overflowY: "scroll", height: 500 }}>
            <Table stickyHeader aria-label="sticky table" size="small">
              <TableHead>
                <TableRow>
                  <TableCell>Step</TableCell>
                  <TableCell>Message</TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {procedure.map((row) => (
                  <TableRow
                    key={row.Step}
                    sx={{ "&:last-child td, &:last-child th": { border: 0 } }}
                  >
                    <TableCell component="th" scope="row">
                      {row.Step}
                    </TableCell>
                    <TableCell>{row.Message}</TableCell>
                  </TableRow>
                ))}
              </TableBody>{" "}
            </Table>
          </Box>
        </Box>
      </Box>
    </Modal>
  );
};

export default React.memo(ProcedureModal);
