import React from "react";

import {
  Button,
  ButtonGroup,
  IconButton,
  InputAdornment,
  Box,
  Input,
  FormControl,
  InputLabel,
  Switch,
  FormControlLabel,
  Toolbar,
  Tooltip,
  ToggleButtonGroup,
} from "@mui/material";

import AddBoxIcon from "@mui/icons-material/AddBox";
import ContentCopyIcon from "@mui/icons-material/ContentCopy";
import DeleteIcon from "@mui/icons-material/Delete";
import VerifiedIcon from "@mui/icons-material/Verified";
import InventoryIcon from "@mui/icons-material/Inventory";
import GppBadIcon from "@mui/icons-material/GppBad";
import SearchIcon from "@mui/icons-material/Search";

import TooltipToggleButton from "../TooltipToggleButton/TooltipToggleButton";

const EnhancedTableToolbar = ({
  anyRowSelected,
  setMode,
  setOpen,
  handleDelete,
  setPage,
  filter,
  setFilter,
}) => {
  const [status, setStatus] = React.useState("");

  const handleStatusChange = (event, value) => {
    value
      ? setFilter({ ...filter, status: value })
      : setFilter({ ...filter, status: "" });
    setStatus(value);
  };

  const handleShowAllChange = (event) => {
    setFilter({ ...filter, showAll: event.target.checked });
    setPage(0);
  };

  const handleOpen = (mode) => {
    setMode(mode);
    setOpen(true);
  };

  const handleSearch = (event) => {
    setFilter({ ...filter, search: event.target.value.toLowerCase() });
  };

  return (
    <Toolbar
      sx={{
        pl: { sm: 2 },
        pr: { xs: 1, sm: 1 },
        display: "flex",
        justifyContent: "space-between",
        // width: '100%',
      }}
    >
      <FormControl sx={{ m: 1, width: "15ch" }} variant="standard">
        <InputLabel htmlFor="recipe-search">Recipe Search</InputLabel>
        <Input
          id="recipe-search"
          type="text"
          onChange={(event) => handleSearch(event)}
          endAdornment={
            <InputAdornment position="end">
              <IconButton
                aria-label="search-icon"
                onClick={() => console.log("click")}
                onMouseDown={() => console.log("mousedown")}
              >
                <SearchIcon />
              </IconButton>
            </InputAdornment>
          }
        />
      </FormControl>
      <Box sx={{ display: "flex" }}>
        <Tooltip title="Show All Versions">
          <FormControlLabel
            control={
              <Switch
                onChange={handleShowAllChange}
                inputProps={{ "aria-label": "controlled" }}
              />
            }
            label="Show All"
          />
        </Tooltip>
        <ToggleButtonGroup
          color="primary"
          value={status}
          exclusive
          onChange={handleStatusChange}
          aria-label="Recipe Status Filter"
        >
          <TooltipToggleButton
            children={<VerifiedIcon />}
            title={"Show Approved"}
            value="Approved"
            variant="contained"
          />
          <TooltipToggleButton
            children={<InventoryIcon />}
            title={"Show Registered"}
            value="Registered"
            variant="contained"
          />
          <TooltipToggleButton
            children={<GppBadIcon />}
            title={"Show Obsolete"}
            value="Obsolete"
            variant="contained"
          />
        </ToggleButtonGroup>

        <ButtonGroup color="primary" sx={{ mx: 2 }}>
          <Tooltip title="New">
            <Box>
              <Button
                variant="contained"
                onClick={() => handleOpen("New")}
                sx={{ height: "100%" }}
              >
                <AddBoxIcon />
              </Button>
            </Box>
          </Tooltip>
          <Tooltip title="Copy">
            <Box>
              <Button
                variant="contained"
                onClick={() => handleOpen("Copy")}
                disabled={!anyRowSelected}
                sx={{ height: "100%" }}
              >
                <ContentCopyIcon />
              </Button>
            </Box>
          </Tooltip>
          <Tooltip title="Delete">
            <Box>
              <Button
                variant="contained"
                disabled={!anyRowSelected}
                onClick={handleDelete}
                sx={{ height: "100%" }}
              >
                <DeleteIcon />
              </Button>
            </Box>
          </Tooltip>
        </ButtonGroup>
        <ButtonGroup color="secondary" sx={{ mr: 2 }}>
          <Button disabled={!anyRowSelected}>Import</Button>
          <Button disabled={!anyRowSelected}>Export</Button>
        </ButtonGroup>
      </Box>
    </Toolbar>
  );
};

export default EnhancedTableToolbar;
