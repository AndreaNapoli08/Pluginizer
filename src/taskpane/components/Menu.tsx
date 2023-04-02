import * as React from 'react';
import Accordion from '@mui/material/Accordion';
import AccordionSummary from '@mui/material/AccordionSummary';
import AccordionDetails from '@mui/material/AccordionDetails';
import Typography from '@mui/material/Typography';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';
import {TipografiaButton} from './TipografiaButton'
import {Blocks} from './Blocks';

export const Menu = () => {
  return (
    <div>
      <Accordion>
        <AccordionSummary
          expandIcon={<ExpandMoreIcon />}
          aria-controls="panel1a-content"
          id="panel1a-header"
          style={{backgroundColor:"transparent"}}
        >
          <Typography variant="h6"><b>Documents</b></Typography>
        </AccordionSummary>
        <AccordionDetails>
          <Typography>
            Da implementare nella tesi
          </Typography>
        </AccordionDetails>
      </Accordion>
      <Accordion>
        <AccordionSummary
          expandIcon={<ExpandMoreIcon />}
          aria-controls="panel2a-content"
          id="panel2a-header"
        >
          <Typography variant="h6"><b>Structures</b></Typography>
        </AccordionSummary>
        <AccordionDetails>
          <Typography>
            Da implementare nella tesi
          </Typography>
        </AccordionDetails>
      </Accordion>
      <Accordion>
        <AccordionSummary
          expandIcon={<ExpandMoreIcon />}
          aria-controls="panel3a-content"
          id="panel3a-header"
        >
          <Typography variant="h6"><b>Blocks</b></Typography>
        </AccordionSummary>
        <AccordionDetails>
          <Typography>
            <Blocks />
          </Typography>
        </AccordionDetails>
      </Accordion>
      <Accordion>
        <AccordionSummary
          expandIcon={<ExpandMoreIcon />}
          aria-controls="panel4a-content"
          id="panel4a-header"
        >
          <Typography variant="h6"><b>Inlines</b></Typography>
        </AccordionSummary>
        <AccordionDetails>
          <TipografiaButton />
        </AccordionDetails>
      </Accordion>
      <Accordion>
        <AccordionSummary
          expandIcon={<ExpandMoreIcon />}
          aria-controls="panel5a-content"
          id="panel5a-header"
        >
          <Typography variant="h6"><b>Globals Sustainability Goals</b></Typography>
        </AccordionSummary>
        <AccordionDetails>
          <Typography>
            Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse
            malesuada lacus ex, sit amet blandit leo lobortis eget.
          </Typography>
        </AccordionDetails>
      </Accordion>
      <Accordion>
        <AccordionSummary
          expandIcon={<ExpandMoreIcon />}
          aria-controls="panel6a-content"
          id="panel6a-header"
        >
          <Typography variant="h6"><b>Metadata</b></Typography>
        </AccordionSummary>
        <AccordionDetails>
          <Typography>
           Da implementare nella tesi
          </Typography>
        </AccordionDetails>
      </Accordion>
    </div>
  )
}

