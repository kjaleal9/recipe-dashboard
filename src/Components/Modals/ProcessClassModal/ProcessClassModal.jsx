import React, { useState, useEffect } from "react";
import {
  Modal,
  Box,
  Typography,
  Accordion,
  AccordionSummary,
  AccordionDetails,
  Chip,
  Button,
  Tooltip,
} from "@mui/material";

import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import EditIcon from "@mui/icons-material/Edit";

const ProcessClassModal = (props) => {
  const {
    recipe,
    version,
    open,
    setOpenProcessClassModal,
    selected,
    setMode,
    setOpen,
  } = props;

  const [equipment, setEquipment] = useState([]);
  const [RPC, setRPC] = useState([]);
  const [expanded, setExpanded] = React.useState(false);

  const handleExpand = (panel) => (event, isExpanded) => {
    setExpanded(isExpanded ? panel : false);
  };

  const getRPC = () =>
    fetch(`/process-classes/required/${recipe}/${version}`).then((response) =>
      response.json()
    );

  const getEquipment = () =>
    fetch("/equipment").then((response) => response.json());

  function getAllData() {
    return Promise.all([getRPC(), getEquipment()]);
  }

  useEffect(() => {
    getAllData()
      .then(([requiredProcessClasses, requiredEquipment]) => {
        setRPC(requiredProcessClasses);
        setEquipment(requiredEquipment);
      })
      .catch((err) => console.log(err));
  }, []);

  useEffect(() => {
    getRPC().then((data) => setRPC(data));
  }, [selected]);

  const filterEquipment = (processClassID) => {
    return equipment.filter(
      (equipment) => equipment.ProcessClass_ID === processClassID
    );
  };

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
  };

  const handleClose = () => {
    setOpenProcessClassModal(false);
    setExpanded(false);
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
      style={{ overflowY: "scroll" }}
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
              Required Process Classes
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
          {RPC.map((processClass, index) => (
            <Accordion
              expanded={expanded === `panel${index}`}
              onChange={handleExpand(`panel${index}`)}
              key={`panel${index}`}
            >
              <AccordionSummary
                expandIcon={<ExpandMoreIcon />}
                aria-controls={`panel${index}a-content`}
                id={`panel${index}a-header`}
              >
                <Box
                  display="flex"
                  justifyContent="space-between"
                  sx={{ width: "100%" }}
                >
                  <Box>
                    <Typography>
                      {processClass.ProcessClass_Name} -{" "}
                      {processClass.Equipment_Name}
                    </Typography>
                  </Box>
                  <Box justifySelf={"flex-end"}>
                    {processClass.IsMainBatchUnit && (
                      <Chip color="info" label="Main Batch Unit" />
                    )}
                  </Box>
                </Box>
              </AccordionSummary>
              <AccordionDetails>
                {equipment
                  .filter(
                    (item) => +item.ProcessClass_ID === +processClass.PClass_ID
                  )
                  .map((filteredEquipment) => (
                    <Typography key={filteredEquipment.Name}>
                      {filteredEquipment.Name}
                    </Typography>
                  ))}
              </AccordionDetails>
            </Accordion>
          ))}
        </Box>
      </Box>
    </Modal>
  );
};

export default React.memo(ProcessClassModal);
