import React, { useState, Fragment } from "react";
import { DragDropContext, Droppable, Draggable } from "react-beautiful-dnd";

import {
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Paper,
  Box,
} from "@mui/material";

import MenuIcon from "@mui/icons-material/Menu";

const RecipeProcedureTable = ({
  rows,
  onDragEnd,
  handleClick,
  isSelected,
  isProcedureEditable,
}) => {
  return (
    <Box>
      <DragDropContext onDragEnd={onDragEnd}>
        <Droppable droppableId="table" isDropDisabled={!isProcedureEditable}>
          {(provided) => (
            <TableContainer
              sx={{ overflowY: "scroll", height: "80vh" }}
              {...provided.droppableProps}
              ref={provided.innerRef}
            >
              <Table aria-labelledby="tableTitle" size="small" stickyHeader>
                <TableHead>
                  <TableRow>
                    <TableCell></TableCell>
                    <TableCell>Step</TableCell>
                    <TableCell>Message</TableCell>
                  </TableRow>
                </TableHead>
                <TableBody>
                  {rows.map((row, index) => {
                    const isItemSelected = isSelected(row.Step);
                    const labelId = `enhanced-table-checkbox-${index}`;

                    return (
                      <Draggable
                        key={row.ID}
                        draggableId={`${row.ID}-${row.Message}`}
                        index={index}
                        isDragDisabled={!isProcedureEditable}
                      >
                        {(provided) => (
                          <TableRow
                            {...provided.draggableProps}
                            ref={provided.innerRef}
                            hover
                            onClick={(event) => {
                              handleClick(
                                event,
                                provided.draggableProps[
                                  "data-rbd-draggable-id"
                                ],
                                row
                              );
                            }}
                            role="checkbox"
                            aria-checked={isItemSelected}
                            tabIndex={-1}
                            key={row.ID}
                            selected={isItemSelected}
                            sx={{
                              cursor: "pointer",
                              border: "solid 1px grey",
                              backgroundColor: "#fefefe",
                              boxShadow: "1px 1px 2px rgba(0, 0, 0, 0.25)",
                            }}
                          >
                            <TableCell
                              {...provided.dragHandleProps}
                              alignItems={"center"}
                            >
                              <MenuIcon />
                            </TableCell>
                            <TableCell component="th" id={labelId} scope="row">
                              {index + 1}
                            </TableCell>
                            <TableCell>{row.Message}</TableCell>
                          </TableRow>
                        )}
                      </Draggable>
                    );
                  })}
                  {provided.placeholder}
                </TableBody>
              </Table>
            </TableContainer>
          )}
        </Droppable>
      </DragDropContext>
    </Box>
  );
};

export default RecipeProcedureTable;
