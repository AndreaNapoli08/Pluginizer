// licenza d'uso riservata ad Andrea Napoli e all'università si Bologna
import * as React from 'react';
import { useState } from 'react';
import Box from '@mui/material/Box';
import Button from '@mui/material/Button';
import Typography from '@mui/material/Typography';
import Modal from '@mui/material/Modal';

const boxStyle = {
    position: 'absolute' as 'absolute',
    top: '50%',
    left: '50%',
    transform: 'translate(-50%, -50%)',
    width: 300,
    bgcolor: 'background.paper',
    border: '2px solid #000',
    boxShadow: 24,
    p: 4,

    '@media only screen and (min-width: 75px) and (max-width: 375px)': {
        width: 200
    },
};


export const ModalInfo = () => {
    const [open, setOpen] = useState(false);
    const handleOpen = () => setOpen(true);
    const handleClose = () => setOpen(false);
  
    return (
      <div>
        <Button 
          color="inherit" 
          style = {{
            marginRight: "10px",
            border: "1px solid black",
            borderRadius: "10px",
            width: "75px",
            height: "40px"
          }}
          onClick={handleOpen}>
            About
        </Button>
        <Modal
          open={open}
          onClose={handleClose}
          aria-labelledby="modal-modal-title"
          aria-describedby="modal-modal-description"
        >
          <Box sx={boxStyle}>
            <Typography id="modal-modal-title" variant="h6" component="h2">
              Document Optimizer
            </Typography>
            <Typography id="modal-modal-description" sx={{ mt: 2 }}>
              <b>Autohor: </b> Andrea Napoli <br />
              <b>Goal: </b> create a plug-in that helps the user to manage and organizate the document<br />
              <b>Released: </b> 6 may 2023<br />
              <b>Version: </b> 0.1<br />
            </Typography>
          </Box>
        </Modal>
      </div>
    );
  }