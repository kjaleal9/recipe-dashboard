import React from 'react';
import {
  Box,
  List,
  ListItem,
  ListItemButton,
  ListItemText,
  Stack,
} from '@mui/material';
import { DragDropContext, Droppable, Draggable } from 'react-beautiful-dnd';

const RecipeProcedure = () => {
  return (
    <DragDropContext>
      <Droppable droppableId='characters'>
        {(provided, snapshot) => (
          <Stack
            direction='column'
            spacing={1}
            className='characters'
            {...provided.droppableProps}
            ref={provided.innerRef}
          >
            {[1, 2, 3, 4, 5, 6, 7, 8].map((x, index) => (
              <Draggable key={x} draggableId={`${x}`} index={index}>
                {(provided, snapshot) => (
                  <div
                    ref={provided.innerRef}
                    {...provided.draggableProps}
                    {...provided.dragHandleProps}
                  >
                    <ListItemButton
                      role='listitem'
                      button
                      sx={{ cursor: 'pointer' }}
                    >
                      <ListItemText id={x} primary={`${x}`} />
                    </ListItemButton>
                  </div>
                )}
              </Draggable>
            ))}
          </Stack>
        )}
      </Droppable>
    </DragDropContext>
  );
};

export default RecipeProcedure;
