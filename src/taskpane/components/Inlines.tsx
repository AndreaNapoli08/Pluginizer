import * as React from 'react';
import { useState } from 'react';
import Button from '@mui/material/Button';
import {TipografiaButton} from './Inlines/TipografiaButton'
import {FirstStyles} from './Inlines/FirstStyles'
import {ImportantEntities} from './Inlines/ImportantEntities'
import {OtherEntities} from './Inlines/OtherEntities'
import {Informative} from './Inlines/Informative'
import { ExpandWords } from './Inlines/ExpandWords';
export const Inlines = ({ onExpandedTextChangeMenu }) => {
  const [expandedText, setExpandedText] = useState(""); // stato per la variabile expandedText

  // Funzione di callback per aggiornare il valore di expandedText
  const handleExpandedTextChange = (text) => {
    setExpandedText(text);
  }

  onExpandedTextChangeMenu(expandedText);

  return (
    <div>
      <FirstStyles />

      <hr />
      <ImportantEntities />
      <OtherEntities />
      <Informative />

      <hr />
      <div style={{display: 'flex', justifyContent: 'center', marginBottom: '5px', fontSize: '20px'}}>
        Tipography    
      </div>
      <TipografiaButton expandedText={expandedText}/>

      <ExpandWords onExpandedTextChange={handleExpandedTextChange} />
      
      <div style={{ display: "flex", justifyContent: "center", alignItems: "center", marginTop: '10px' }}>
        <Button variant="outlined" color="inherit">
          Apply to all instances
        </Button>
      </div>
    </div>
  )
}