import React from "react";
import PropTypes from "prop-types";

import Box from "@mui/material/Box";
import TableCell from "@mui/material/TableCell";
import TableHead from "@mui/material/TableHead";
import TableRow from "@mui/material/TableRow";
import TableSortLabel from "@mui/material/TableSortLabel";
import { visuallyHidden } from "@mui/utils";

const headCells = [
  {
    id: "RID",
    numeric: false,
    disablePadding: false,
    label: "RID",
  },
  {
    id: "Version",
    numeric: true,
    disablePadding: false,
    label: "Ver",
  },
  {
    id: "VersionDate",
    type: Date,
    disablePadding: false,
    label: "Date",
  },
  // {
  //   id: 'RecipeType',
  //   numeric: false,
  //   disablePadding: false,
  //   label: 'Type',
  // },
  {
    id: "Description",
    numeric: false,
    disablePadding: false,
    label: "Recipe Description",
  },
  {
    id: "Status",
    numeric: false,
    disablePadding: false,
    label: "Status",
  },
  {
    id: "ProductID",
    numeric: true,
    disablePadding: false,
    label: "Product ID",
  },
  {
    id: "Name",
    numeric: false,
    disablePadding: false,
    label: "Product Description",
    type: "text",
  },
];

function EnhancedTableHead(props) {
  const { order, orderBy, onRequestSort } = props;
  const createSortHandler = (property) => (event) => {
    onRequestSort(event, property);
  };

  return (
    <TableHead>
      <TableRow>
        {headCells.map((headCell) => (
          <TableCell
            key={headCell.id}
            align={headCell.numeric ? "right" : "left"}
            padding={headCell.disablePadding ? "none" : "normal"}
            sortDirection={orderBy === headCell.id ? order : false}
            sx={{
              maxWidth: headCell.width ? headCell.width : "50px",
            }}
          >
            <TableSortLabel
              active={orderBy === headCell.id}
              direction={orderBy === headCell.id ? order : "asc"}
              onClick={createSortHandler(headCell.id)}
            >
              {headCell.label}
              {orderBy === headCell.id ? (
                <Box component="span" sx={visuallyHidden}>
                  {order === "desc" ? "sorted descending" : "sorted ascending"}
                </Box>
              ) : null}
            </TableSortLabel>
          </TableCell>
        ))}
      </TableRow>
    </TableHead>
  );
}

EnhancedTableHead.propTypes = {
  onRequestSort: PropTypes.func.isRequired,
  order: PropTypes.oneOf(["asc", "desc"]).isRequired,
  orderBy: PropTypes.string.isRequired,
};

export default EnhancedTableHead;
