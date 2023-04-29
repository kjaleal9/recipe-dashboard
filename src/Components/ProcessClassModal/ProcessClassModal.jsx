import React, { useState, useEffect } from "react";
import {
  Modal,
  Box,
  Typography,
  Accordion,
  AccordionSummary,
  AccordionDetails,
  requirePropFactory,
} from "@mui/material";

import ExpandMoreIcon from "@mui/icons-material/ExpandMore";

const ProcessClassModal = (props) => {
  const { recipe, version, open, setOpenProcessClassModal, selected } = props;

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
          <Typography
            component="h1"
            variant="h5"
            color="inherit"
            noWrap
            sx={{ flexGrow: 1, alignSelf: "center" }}
          >
            Required Process Classes
          </Typography>
          {RPC.map((processClass, index) => (
            <Accordion
              expanded={expanded === `panel${index}`}
              onChange={handleExpand(`panel${index}`)}
            >
              <AccordionSummary
                expandIcon={<ExpandMoreIcon />}
                aria-controls={`panel${index}a-content`}
                id={`panel${index}a-header`}
              >
                <Typography>
                  {processClass.ProcessClass_Name} -{" "}
                  {processClass.Equipment_Name}
                </Typography>
              </AccordionSummary>
              <AccordionDetails>
                {equipment
                  .filter(
                    (item) => +item.ProcessClass_ID === +processClass.PClass_ID
                  )
                  .map((filteredEquipment) => (
                    <Typography>{filteredEquipment.Name}</Typography>
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
