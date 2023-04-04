import React, { useState, Fragment } from "react";

import {
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TablePagination,
  TableRow,
} from "@mui/material";

import EnhancedTableHead from "../EnhancedTableHead/EnhancedTableHead";

import { getComparator, stableSort } from "../../../utilities";

const CustomTableBody = ({
  selected,
  setSelected,
  setMode,
  setOpen,
  rows,
  page,
  setPage,
}) => {
  const [rowsPerPage, setRowsPerPage] = useState(20);
  const [order, setOrder] = useState("asc");
  const [orderBy, setOrderBy] = useState("RID");

  const handleRequestSort = (event, property) => {
    const isAsc = orderBy === property && order === "asc";
    setOrder(isAsc ? "desc" : "asc");
    setOrderBy(property);
  };

  const handleClick = (event, name, version, row) => {
    const selectedRow = selected.RID === name && selected.Version === version;
    let newSelected = "";

    if (selectedRow === false) {
      newSelected = row;
    }

    setSelected(newSelected);
  };

  const handleDoubleClick = (event, row) => {
    setSelected(row);
    setMode("View");
    setOpen(true);
  };

  const handleChangePage = (event, newPage) => {
    setPage(newPage);
  };

  const handleChangeRowsPerPage = (event) => {
    setRowsPerPage(parseInt(event.target.value, 10));
    setPage(0);
  };

  // const isSelected = row => selected.indexOf(row) !== -1;
  function isSelected(RID, version) {
    return selected.RID === RID && selected.Version === version;
  }

  // Avoid a layout jump when reaching the last page with empty rows.
  const emptyRows =
    page > 0 ? Math.max(0, (1 + page) * rowsPerPage - rows.length) : 0;

  return (
    <Fragment>
      <TableContainer>
        <Table sx={{ minWidth: 750 }} aria-labelledby="tableTitle" size="small">
          <EnhancedTableHead
            order={order}
            orderBy={orderBy}
            onRequestSort={handleRequestSort}
          />
          <TableBody>
            {stableSort(rows, getComparator(order, orderBy))
              .slice(page * rowsPerPage, page * rowsPerPage + rowsPerPage)
              .map((row, index) => {
                const isItemSelected = isSelected(row.RID, row.Version);
                const labelId = `enhanced-table-checkbox-${index}`;

                return (
                  <TableRow
                    hover
                    onClick={(event) =>
                      handleClick(event, row.RID, row.Version, row)
                    }
                    onDoubleClick={(event) => handleDoubleClick(event, row)}
                    role="checkbox"
                    aria-checked={isItemSelected}
                    tabIndex={-1}
                    key={row.RID + row.Version}
                    selected={isItemSelected}
                    sx={{ cursor: "pointer" }}
                  >
                    <TableCell
                      component="th"
                      id={labelId}
                      scope="row"
                      sx={{ width: "10ch" }}
                    >
                      {row.RID}
                    </TableCell>
                    <TableCell align="right" sx={{ width: "11ch" }}>
                      {row.Version}
                    </TableCell>
                    <TableCell align="left" sx={{ width: "11ch" }}>
                      {new Date(row.VersionDate).toLocaleDateString()}
                    </TableCell>
                    {/* <TableCell align="left" sx={{ width: "100px" }}>
                              {row.RecipeType}
                            </TableCell> */}
                    <TableCell align="left" sx={{ width: "180px" }}>
                      {row.Description}
                    </TableCell>
                    <TableCell align="left" sx={{ width: "11ch" }}>
                      {row.Status}
                    </TableCell>
                    <TableCell align="right" sx={{ width: "13ch" }}>
                      {row.ProductID}
                    </TableCell>
                    <TableCell align="left" sx={{ width: "180px" }}>
                      {row.Name}
                    </TableCell>
                  </TableRow>
                );
              })}
            {emptyRows > 0 && (
              <TableRow
                style={{
                  height: 33 * emptyRows,
                }}
              >
                <TableCell colSpan={6} />
              </TableRow>
            )}
          </TableBody>
        </Table>
      </TableContainer>
      <TablePagination
        rowsPerPageOptions={[5, 10, 15, 20]}
        component="div"
        count={rows.length}
        rowsPerPage={rowsPerPage}
        page={page}
        onPageChange={handleChangePage}
        onRowsPerPageChange={handleChangeRowsPerPage}
      />
    </Fragment>
  );
};

export default CustomTableBody;
