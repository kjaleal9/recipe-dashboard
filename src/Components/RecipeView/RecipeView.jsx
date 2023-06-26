import React, { useState, useEffect } from "react";
import { Box, Typography, Divider, Chip, Button } from "@mui/material";

import ProcessClassModal from "../Modals/ProcessClassModal/ProcessClassModal";
import ProcedureModal from "../Modals/ProcedureModal/ProcedureModal";

const RecipeView = (props) => {
  const { selected, setMode, setOpen } = props;
  const [loading, setLoading] = useState(false);
  const [openProcessClassModal, setOpenProcessClassModal] = useState(false);
  const [openProcedureModal, setOpenProcedureModal] = useState(false);
  const [procedure, setProcedure] = useState([]);
  const [equipment, setEquipment] = useState([]);
  const [RPC, setRPC] = useState([]);

  const handleButtonProcessClasses = () => {
    setOpenProcessClassModal(true);
  };

  const handleButtonProcedure = () => {
    setOpenProcedureModal(true);
  };

  const getProcedure = () =>
    fetch(
      `/recipes/procedure/${selected.RID}/${selected.Version}/condense`
    ).then((response) => response.json());

  const getRPC = () =>
    fetch(`/process-classes/required/${selected.RID}/${selected.Version}`).then(
      (response) => response.json()
    );

  const getEquipment = () =>
    fetch("/equipment").then((response) => response.json());

  function getAllData() {
    return Promise.all([getRPC(), getEquipment(), getProcedure()]);
  }

  useEffect(() => {
    getAllData()
      .then(
        ([requiredProcessClasses, requiredEquipment, selectedProcedure]) => {
          setRPC(requiredProcessClasses);
          setEquipment(requiredEquipment);
          setProcedure(selectedProcedure);
        }
      )
      .catch((err) => console.log(err));
  }, []);

  useEffect(() => {
    setLoading(true);
    Promise.all([getRPC(), getProcedure()]).then(
      ([requiredProcessClasses, selectedProcedure]) => {
        setRPC(requiredProcessClasses);
        setProcedure(selectedProcedure);
        setLoading(false);
      }
    );
  }, [selected]);

  useEffect(() => {
    getRPC().then((data) => setRPC(data));
  }, [selected]);

  return (
    <Box
      sx={{ display: "flex", flexDirection: "column", p: 3, maxHeight: "88vh" }}
    >
      {loading ? (
        <Typography>Loading</Typography>
      ) : (
        <Box>
          <ProcessClassModal
            RPC={RPC}
            equipment={equipment}
            open={openProcessClassModal}
            setOpen={setOpen}
            setMode={setMode}
            setOpenProcessClassModal={setOpenProcessClassModal}
          />
          <ProcedureModal
            procedure={procedure}
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
                <Typography
                  component="p"
                  variant="button"
                  color="inherit"
                  noWrap
                >
                  {selected.BatchSizeMax} gal
                </Typography>
              </Box>
              <Divider orientation="vertical" variant="middle" light />
              <Box>
                <Typography
                  variant="overline"
                  style={{ borderBottom: "1px solid #bdbdbd" }}
                >
                  Min
                </Typography>
                <Typography
                  component="p"
                  variant="button"
                  color="inherit"
                  noWrap
                >
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
                <Typography
                  component="p"
                  variant="button"
                  color="inherit"
                  noWrap
                >
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
            <Button
              variant="contained"
              sx={{ m: 1, width: "75%" }}
              onClick={handleButtonProcedure}
            >
              Trains
            </Button>
          </Box>
          <Divider sx={{ m: 2 }} light />
          <Box display="flex" flexDirection="column" height={400}>
            <Typography
              component="h1"
              variant="h6"
              sx={{ mb: 1, alignSelf: "center" }}
            >
              USER COMMENT
            </Typography>
            <Typography
              component="p"
              variant="p"
              color="inherit"
              sx={{
                alignSelf: "center",
                pb: 0.25,
                overflowY: "scroll",
                height:'auto'
              }}
            >
              Lorem ipsum, dolor sit amet consectetur adipisicing elit.
              Accusamus dolor itaque accusantium cupiditate nam eaque, ut
              pariatur incidunt impedit unde quae totam tempore officia alias
              quidem fugiat dolorum? Nulla laudantium aliquam magnam repellendus
              ullam sequi nesciunt natus nisi adipisci odio, fugiat porro
              perferendis voluptatum ex beatae, modi minus quam repudiandae
              sapiente tenetur amet iusto eum accusamus? Laudantium hic sit
              alias aut necessitatibus inventore esse. Velit ea quos blanditiis
              labore, atque nam obcaecati quaerat error distinctio, porro
              accusantium reprehenderit eum nostrum a repellat, ad ipsum placeat
              fugiat animi esse nisi sit odio vero sapiente. Assumenda,
              voluptate nesciunt incidunt tempore ad adipisci pariatur placeat
              consequuntur maiores voluptas necessitatibus dignissimos deserunt
              aut suscipit quae omnis at minus officiis dolorum quibusdam harum
              minima. Iste, eius soluta quia modi est harum magnam, similique,
              natus adipisci nulla sint nam aperiam sed veritatis debitis
              reiciendis minus eligendi ut dolorum temporibus ad. Accusamus quos
              libero quam eveniet sed id rem nisi, necessitatibus ipsum cumque
              amet temporibus atque cupiditate possimus debitis, provident
              soluta perferendis voluptate. Voluptatibus quam nobis omnis! Autem
              dolorum veniam perferendis aperiam corrupti beatae similique
              velit, explicabo asperiores commodi tempora atque, optio voluptas
              error voluptates excepturi. Aspernatur eum laborum nihil
              doloremque cumque facere pariatur impedit quasi ducimus!
            </Typography>
          </Box>
        </Box>
      )}
    </Box>
  );
};

export default RecipeView;
