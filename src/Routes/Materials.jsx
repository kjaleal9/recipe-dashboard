import React, { useState } from "react";
import { TextField } from "@mui/material";

const Materials = ({ label, value, onChange }) => {

  const [inputValue, setInputValue] = useState(value);

  const handleInputChange = (event) => {
    const newValue = event.target.value;
    setInputValue(newValue);
    onChange(newValue);
  };

  return (
    <TextField
      label={label}
      variant="outlined"
      multiline
      rows={4}
      value={inputValue}
      onChange={handleInputChange}
      sx={{width:'100%'}}
    />
  );
};

export default Materials;
