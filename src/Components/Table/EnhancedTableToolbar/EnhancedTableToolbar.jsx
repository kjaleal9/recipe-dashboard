import React from "react";
import PropTypes from "prop-types";

import Button from "@mui/material/Button";
import ButtonGroup from "@mui/material/ButtonGroup";
import IconButton from "@mui/material/IconButton";
import InputAdornment from "@mui/material/InputAdornment";
import SearchIcon from "@mui/icons-material/Search";
import Toolbar from "@mui/material/Toolbar";
import ToggleButton from "@mui/material/ToggleButton";
import ToggleButtonGroup from "@mui/material/ToggleButtonGroup";
import {
  Box,
  Input,
  FormControl,
  InputLabel,
  Switch,
  FormGroup,
  FormControlLabel,
  Tooltip,
} from "@mui/material";

import AddBoxIcon from "@mui/icons-material/AddBox";
import ContentCopyIcon from "@mui/icons-material/ContentCopy";
import DeleteIcon from "@mui/icons-material/Delete";
import VerifiedIcon from "@mui/icons-material/Verified";
import InventoryIcon from "@mui/icons-material/Inventory";
import GppBadIcon from "@mui/icons-material/GppBad";

const EnhancedTableToolbar = (props) => {
  const [status, setStatus] = React.useState("web");

  const handleStatusChange = (event, statusFilter) => {
    const statusFilters = {
      approved: () =>
        setFilter({
          ...filter,
          approved: true,
          registered: false,
          obsolete: false,
        }),
      registered: () =>
        setFilter({
          ...filter,
          registered: true,
          approved: false,
          obsolete: false,
        }),
      obsolete: () =>
        setFilter({
          ...filter,
          obsolete: true,
          approved: false,
          registered: false,
        }),
    };
    if (statusFilters[statusFilter]) {
      statusFilters[statusFilter]();
    } else {
      setFilter({
        ...filter,
        approved: false,
        registered: false,
        obsolete: false,
      });
    }
    setStatus(statusFilter);
  };
  const {
    anyRowSelected,
    setMode,
    setOpen,
    handleDelete,
    handleSearch,
    setPage,
    filter,
    setFilter,
  } = props;

  const handleShowAllChange = (event) => {
    setFilter({ ...filter, showAll: event.target.checked });
    setPage(0);
  };

  const handleOpen = (mode) => {
    setMode(mode);
    setOpen(true);
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
          aria-label="Platform"
        >
          <ToggleButton value="approved">
            <Tooltip title="Show Approved">
              <VerifiedIcon />
            </Tooltip>
          </ToggleButton>
          <ToggleButton value="registered">
            <Tooltip title="Show Registered">
              <InventoryIcon />
            </Tooltip>
          </ToggleButton>
          <ToggleButton value="obsolete">
            <Tooltip title="Show Obsolete">
              <GppBadIcon />
            </Tooltip>
          </ToggleButton>
        </ToggleButtonGroup>

        <ButtonGroup color="primary" sx={{ mx: 2 }}>
          <Tooltip title="New">
            <Button variant="contained" onClick={() => handleOpen("New")}>
              <AddBoxIcon />
            </Button>
          </Tooltip>
          <Tooltip title="Copy">
            <Button
              variant="contained"
              onClick={() => handleOpen("Copy")}
              disabled={!anyRowSelected}
            >
              <ContentCopyIcon />
            </Button>
          </Tooltip>
          <Tooltip title="Delete">
            <Button
              variant="contained"
              disabled={!anyRowSelected}
              onClick={handleDelete}
            >
              <DeleteIcon />
            </Button>
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

EnhancedTableToolbar.propTypes = {
  anyRowSelected: PropTypes.bool.isRequired,
};

export default EnhancedTableToolbar;
