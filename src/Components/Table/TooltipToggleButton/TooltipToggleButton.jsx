import React from "react";
import { Tooltip, ToggleButton } from "@mui/material";

const TooltipToggleButton = ({ children, title, ...props }) => (
  <Tooltip title={title}>
    <ToggleButton {...props}>{children}</ToggleButton>
  </Tooltip>
);

export default TooltipToggleButton;
