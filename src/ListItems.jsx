import * as React from "react";
import { Link } from "react-router-dom";
import ListItemButton from "@mui/material/ListItemButton";
import ListItemIcon from "@mui/material/ListItemIcon";
import ListItemText from "@mui/material/ListItemText";
import ListSubheader from "@mui/material/ListSubheader";
import ListAlt from "@mui/icons-material/ListAlt";
import LunchDining from "@mui/icons-material/LunchDining";
import BarChartIcon from "@mui/icons-material/BarChart";
import AssignmentIcon from "@mui/icons-material/Assignment";
import { DragDropContext, Droppable, Draggable } from "react-beautiful-dnd";
import styled from "@emotion/styled";

const StyledLink = styled(Link)`
  color: Black;
  text-decoration: none;
`;

export const mainListItems = (
  <React.Fragment>
    <StyledLink to={"recipes"}>
      <ListItemButton>
        <ListItemIcon>
          <ListAlt />
        </ListItemIcon>
        <ListItemText primary="Recipes" />
      </ListItemButton>
    </StyledLink>
    <StyledLink to={"materials"}>
      <ListItemButton>
        <ListItemIcon>
          <LunchDining />
        </ListItemIcon>
        <ListItemText primary="Materials" />
      </ListItemButton>
    </StyledLink>
    <ListItemButton>
      <ListItemIcon>
        <BarChartIcon />
      </ListItemIcon>
      <ListItemText primary="Reports" />
    </ListItemButton>
  </React.Fragment>
);

export const secondaryListItems = (
  <React.Fragment>
    <ListSubheader component="div" inset>
      Saved reports
    </ListSubheader>
    <ListItemButton>
      <ListItemIcon>
        <AssignmentIcon />
      </ListItemIcon>
      <ListItemText primary="Current month" />
    </ListItemButton>
    <ListItemButton>
      <ListItemIcon>
        <AssignmentIcon />
      </ListItemIcon>
      <ListItemText primary="Last quarter" />
    </ListItemButton>
    <ListItemButton>
      <ListItemIcon>
        <AssignmentIcon />
      </ListItemIcon>
      <ListItemText primary="Year-end sale" />
    </ListItemButton>
  </React.Fragment>
);
