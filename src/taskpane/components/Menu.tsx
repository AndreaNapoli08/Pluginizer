import * as React from 'react';
import { useState, useMemo, useEffect } from "react";
import Accordion from '@mui/material/Accordion';
import AccordionSummary from '@mui/material/AccordionSummary';
import AccordionDetails from '@mui/material/AccordionDetails';
import Typography from '@mui/material/Typography';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';
import {Blocks} from './Blocks';
import {Inlines} from './Inlines';

export const Menu = () => {
  const [expanded, setExpanded] = useState([]);
  const [color, setColor] = useState('transparent');

  const handleChange = (panel) => (isExpanded) => {
    setExpanded(prevExpanded => {
      if (isExpanded) {
        setColor('blue');
        return [...prevExpanded, panel];
      } else {
        setColor('transparent');
        return prevExpanded.filter(p => p !== panel);
      }
    });
  };

  // Derive an array of active panels based on the current state of `expanded`
  const activePanels = useMemo(() => {
    return expanded.reduce((result, panel) => {
      result[panel] = true;
      return result;
    }, {});
  }, [expanded]);

  // Force the component to re-render every time `activePanels` changes
  useEffect(() => {}, [activePanels]);

  return (
    <div>
     <Accordion 
        expanded={activePanels["panel1a"]}
        onChange={handleChange("panel1a")}
      >
        <AccordionSummary
          expandIcon={<ExpandMoreIcon />}
          aria-controls="panel1a-content"
          id="panel1a-header"
          style={{backgroundColor: color}}
        >
          <Typography variant="h6"><b>Documents</b></Typography>
        </AccordionSummary>
        <AccordionDetails>
          <Typography>
            Da implementare nella tesi
          </Typography>
        </AccordionDetails>
      </Accordion>
      <Accordion 
        expanded={activePanels["panel2a"]}
        onChange={handleChange("panel2a")}
      >
        <AccordionSummary
          expandIcon={<ExpandMoreIcon />}
          aria-controls="panel2a-content"
          id="panel2a-header"
          style={{backgroundColor: activePanels["panel2a"] ? "blue" : "transparent"}}
        >
          <Typography variant="h6"><b>Structures</b></Typography>
        </AccordionSummary>
        <AccordionDetails>
          <Typography>
            Da implementare nella tesi
          </Typography>
        </AccordionDetails>
      </Accordion>
      <Accordion 
        expanded={activePanels["panel3a"]}
        onChange={handleChange("panel3a")}
      >
        <AccordionSummary
          expandIcon={<ExpandMoreIcon />}
          aria-controls="panel3a-content"
          id="panel3a-header"
          style={{backgroundColor: activePanels["panel3a"] ? "blue" : "transparent"}}
        >
          <Typography variant="h6"><b>Blocks</b></Typography>
        </AccordionSummary>
        <AccordionDetails>
          <Typography>
            <Blocks />
          </Typography>
        </AccordionDetails>
      </Accordion>
      <Accordion 
        expanded={activePanels["panel4a"]}
        onChange={handleChange("panel4a")}
      >
        <AccordionSummary
          expandIcon={<ExpandMoreIcon />}
          aria-controls="panel4a-content"
          id="panel4a-header"
          style={{backgroundColor: activePanels["panel4a"] ? "blue" : "transparent"}}
        >
          <Typography variant="h6"><b>Inlines</b></Typography>
        </AccordionSummary>
        <AccordionDetails>
          <Inlines />
        </AccordionDetails>
      </Accordion>
      <Accordion 
        expanded={activePanels["panel5a"]}
        onChange={handleChange("panel5a")}
      >
        <AccordionSummary
          expandIcon={<ExpandMoreIcon />}
          aria-controls="panel5a-content"
          id="panel5a-header"
          style={{backgroundColor: activePanels["panel5a"] ? "blue" : "transparent"}}
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
      <Accordion 
        expanded={activePanels["panel6a"]}
        onChange={handleChange("panel6a")}
      >
        <AccordionSummary
          expandIcon={<ExpandMoreIcon />}
          aria-controls="panel6a-content"
          id="panel6a-header"
          style={{backgroundColor: activePanels["panel6a"] ? "blue" : "transparent"}}
        >
          <Typography variant="h6"><b>Metadataaa</b></Typography>
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

