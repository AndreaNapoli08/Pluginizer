import * as React from 'react';
import { useState, useEffect } from 'react';
import FormControlLabel from '@mui/material/FormControlLabel';
import Checkbox from '@mui/material/Checkbox';

export const AllInstances  = ({selectedText, fontStyle, buttonStyle, firstOccurence, first}) => {
    const [allInstances, setAllInstances] = useState(false);

    const handleChangeCheckboxIstances = (event: React.ChangeEvent<HTMLInputElement>) => {
        setAllInstances(event.target.checked);
    };

    if (allInstances) {
      const applyFormatting = async (context) => {
        const range = context.document.body.getRange();
        const searchResults = range.search(selectedText, { matchCase: false, matchWholeWord: false });
        searchResults.load("items");
        await context.sync();
    
        const occurrences = searchResults.items;
    
        occurrences.forEach((occurrence) => {
          // formattazione testo
          switch (buttonStyle) {
            case "bold":
              occurrence.font.bold = !firstOccurence;
              break;
            case "italic":
              occurrence.font.italic = !firstOccurence;
              break;
            case "underline":
              occurrence.font.underline = firstOccurence === "Single" ? "None" : "Single";
              break;
            default:
              break;
          }

          // stili di testo predenfiniti
          switch (fontStyle) {
            case "Hyperlink":
              occurrence.styleBuiltIn = first == "Hyperlink" ? "Normal" : "Hyperlink";
              break;
            case "Heading6":
              occurrence.styleBuiltIn = first == "Heading6" ? "Normal" : "Heading6";
              break;
            case "IntenseEmphasis":
              occurrence.styleBuiltIn = first == "IntenseEmphasis" ? "Normal" : "IntenseEmphasis";
              break;
            default:
              break;
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