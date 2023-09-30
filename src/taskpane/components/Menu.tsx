// licenza d'uso riservata ad Andrea Napoli e all'università si Bologna
import * as React from 'react';
import { useState, useMemo, useEffect } from "react";
import Accordion from '@mui/material/Accordion';
import AccordionSummary from '@mui/material/AccordionSummary';
import AccordionDetails from '@mui/material/AccordionDetails';
import Typography from '@mui/material/Typography';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';
import {Documents} from './Documents';
import {Blocks} from './Blocks';
import {Inlines} from './Inlines';
import {GSG} from './GSG';

export const Menu = () => {
  const [expanded, setExpanded] = useState([]);
  const [expandedText, setExpandedText] = useState("");

  // funzione che gestisce l'apertura e la chiusura dei pannelli
  const handleChange = (panel) => (isExpanded) => {
    setExpanded(prevExpanded => {
      if (isExpanded) {
        return [...prevExpanded, panel]; // Se il pannello è aperto, lo aggiunge all'array di pannelli aperti
      } else {
        return prevExpanded.filter(p => p !== panel); // se il pannello è chiuso lo rimuove dall'array
      }
    });
  };

  // Deriva un array di pannelli attivi in base allo stato corrente di `expanded`
  const activePanels = useMemo(() => {
    return expanded.reduce((result, panel) => {
      result[panel] = !result[panel]; // cambia lo stato del pannello da attivo a chuso e viceversa
      return result;
    }, {});
  }, [expanded]);

  
  useEffect(() => {}, [activePanels]);

  const handleExpandedText = (text) => {
    setExpandedText(text);
  }

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
          style={{backgroundColor: activePanels["panel1a"] ? "lightblue" : "transparent"}}
        >
          <Typography variant="h6"><b>Documents</b></Typography>
        </AccordionSummary>
        <AccordionDetails>
          <Typography>
            <Documents />
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
          style={{backgroundColor: activePanels["panel3a"] ? "lightblue" : "transparent"}}
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
          style={{backgroundColor: activePanels["panel4a"] ? "lightblue" : "transparent"}}
        >
          <Typography variant="h6"><b>Inlines</b></Typography>
        </AccordionSummary>
        <AccordionDetails>
          <Inlines onHandleExpandedText={handleExpandedText}/>
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
          style={{backgroundColor: activePanels["panel5a"] ? "lightblue" : "transparent"}}
        >
          <Typography variant="h6"><b>Globals Sustainability Goals</b></Typography>
        </AccordionSummary>
        <AccordionDetails>
          <Typography>
            <GSG expandedText={expandedText}/>
          </Typography>
        </AccordionDetails>
      </Accordion>
    </div>
  )
}

