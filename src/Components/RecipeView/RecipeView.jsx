import React, { Fragment, useState, useEffect } from "react";
import { Box, Typography, Divider, Chip, Button } from "@mui/material";
import EditIcon from "@mui/icons-material/Edit";

const RecipeView = (props) => {
  const { selected } = props;

  return (
    <Box sx={{ display: "flex", flexDirection: "column", p: 3 }}>
      <Typography
        component="h1"
        variant="h5"
        color="inherit"
        noWrap
        sx={{ flexGrow: 1, alignSelf: "center" }}
      >
        {selected.RID}
      </Typography>
      <Typography
        component="p"
        variant="p"
        color="inherit"
        sx={{ alignSelf: "center", pb: 0.25 }}
      >
        {selected.Description}
      </Typography>
      <Divider sx={{ mb: 2, mt: 2 }} />
      <Box sx={{ display: "flex", justifyContent: "space-around" }}>
        <Box>
          <Typography
            component="p"
            variant="h6"
            color="inherit"
            noWrap
            sx={{ alignSelf: "center", pb: 0.25 }}
          >
            Recipe Info:
          </Typography>
          <Typography
            component="p"
            variant="p"
            color="inherit"
            noWrap
            sx={{ alignSelf: "center", pb: 0.25 }}
          >
            Version: {selected.Version}
          </Typography>
          <Typography
            component="p"
            variant="p"
            color="inherit"
            noWrap
            sx={{ alignSelf: "center" }}
          >
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
                  ? "info"
                  : selected.Status === "Obsolete"
                  ? "error"
                  : "info"
              }
            />
          </Typography>

          <Typography
            component="p"
            variant="p"
            color="inherit"
            noWrap
            sx={{ alignSelf: "center" }}
          >
            Date: {new Date(selected.VersionDate).toDateString()}
          </Typography>
        </Box>
        <Divider orientation="vertical" flexItem />
        <Box>
          <Typography
            component="p"
            variant="h6"
            color="inherit"
            noWrap
            sx={{ alignSelf: "center", pb: 0.25 }}
          >
            Batch Info:
          </Typography>
          <Typography
            component="p"
            variant="p"
            color="inherit"
            noWrap
            sx={{ alignSelf: "center", pb: 0.25 }}
          >
            Max: {selected.BatchSizeMax} gal
          </Typography>
          <Typography
            component="p"
            variant="p"
            color="inherit"
            noWrap
            sx={{ alignSelf: "center", pb: 0.25 }}
          >
            Min: {selected.BatchSizeMin} gal
          </Typography>
          <Typography
            component="p"
            variant="p"
            color="inherit"
            noWrap
            sx={{ alignSelf: "center", pb: 0.25 }}
          >
            Nominal: {selected.BatchSizeMin} gal
          </Typography>
        </Box>
        <Divider sx={{ mb: 2, mt: 2 }} />
      </Box>

      <Divider sx={{ mb: 2, mt: 2 }} />

      <Typography
        component="h5"
        variant="h5"
        color="inherit"
        sx={{ alignSelf: "center", pb: 0.25 }}
      >
        Product Description
      </Typography>
      <Typography
        component="p"
        variant="p"
        color="inherit"
        sx={{ alignSelf: "center", pb: 0.25 }}
      >
        Product: {selected.Name}
      </Typography>
      <Typography
        component="p"
        variant="p"
        color="inherit"
        sx={{ alignSelf: "center", pb: 0.25 }}
      >
        ID: {selected.ProductID}
      </Typography>

      <Button variant="contained" sx={{ width: "50%", alignSelf: "center" }}>
        Bill Of Materials
      </Button>

      <Divider sx={{ mb: 2, mt: 2 }} />

      <Typography
        component="h5"
        variant="h5"
        color="inherit"
        sx={{ alignSelf: "center", pb: 0.25 }}
      >
        Comment
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
        voluptatibus. Deserunt minima laboriosam aperiam, ipsum est perferendis
        suscipit! Vel, iusto modi! Consequuntur ratione ex voluptates sunt
        dolore veniam nulla modi qui ea eligendi libero fuga, minima nam?
        Officiis cum nihil, laborum facilis dignissimos culpa natus consequuntur
        sapiente nobis, ratione quisquam minima minus delectus similique
        praesentium aperiam obcaecati eligendi accusantium officia harum
        veritatis dolorum at odit neque?
      </Typography>
    </Box>
  );
};

export default RecipeView;
