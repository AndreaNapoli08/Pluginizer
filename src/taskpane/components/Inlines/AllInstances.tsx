import * as React from 'react';
import { useState, useEffect } from 'react';
import FormControlLabel from '@mui/material/FormControlLabel';
import Checkbox from '@mui/material/Checkbox';

export const AllInstances  = ({selectedText, buttonStyle, firstOccurence}) => {
    const [allInstances, setAllInstances] = useState(false);

    const fontStyle = buttonStyle;

    const handleChangeCheckboxIstances = (event: React.ChangeEvent<HTMLInputElement>) => {
        setAllInstances(event.target.checked);
    };
    
    if(allInstances){
        const applyFormatting = async (context) => {
            const range = context.document.body.getRange();
            const searchResults = range.search(selectedText, { matchCase: false, matchWholeWord: false });
            searchResults.load("items");
            await context.sync();
        
            const occurrences = searchResults.items;

            occurrences.forEach((occurrence) => {
                if(fontStyle == "bold"){
                  if(firstOccurence){
                    occurrence.font.bold = false;
                  }else{
                    occurrence.font.bold = true;
                  }
                }else if(fontStyle == "italic"){
                  if(firstOccurence){
                    occurrence.font.italic = false;
                  }else{
                    occurrence.font.italic = true;
                  }
                }else if(fontStyle == "underline"){
                  if(firstOccurence == "Single"){
                    occurrence.font.underline = "None";
                  }else{
                    occurrence.font.underline = "Single";
                  }
                }
            });
        
            await context.sync();
        };

        Word.run(async (context) => {
            await applyFormatting(context);
        }).catch((error) => {
        console.error(error);
        });
    }

    return (
        <div>
            <div style={{ display: "flex", justifyContent: "center", alignItems: "center", marginTop: '10px' }}>
                <FormControlLabel 
                    control={<Checkbox checked={allInstances} onChange={handleChangeCheckboxIstances}/>} 
                    label="APPLY TO ALL INSTANCES" 
                    style={{display: 'flex', justifyContent: 'center', alignItems: 'center', marginTop: '10px'}}
                />
            </div>
        </div>
    )
}