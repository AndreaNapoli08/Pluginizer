import * as React from 'react';
import { useState, useEffect } from 'react';
import FormControlLabel from '@mui/material/FormControlLabel';
import Checkbox from '@mui/material/Checkbox';

export const AllInstances  = ({fontStyle, buttonStyle, firstOccurence, first, expandedText}) => {
    const [allInstances, setAllInstances] = useState(false);

    const handleChangeCheckboxIstances = (event: React.ChangeEvent<HTMLInputElement>) => {
        setAllInstances(event.target.checked);
    };
    
    const isLetterOrNumber = (char) => {
      if (typeof char === "undefined") {
        return false;
      }else{
        return /^[a-zA-Z0-9]+$/.test(char);
      }
    }

    if (allInstances) {
      const applyFormatting = async (context) => {
        const range = context.document.body.getRange();
        let selection = context.document.getSelection();
        selection.load();
        await context.sync();

        if(expandedText != selection.text){  // da problemi all'ultima parola della riga
          const startIndex = expandedText.indexOf(selection.text);
          const charBefore = expandedText[startIndex - 1];
          let text = selection.text;
          let spaceCount = text.split(" ").length;
         
          //selezione in avanti fino ad uno di quei caratteri
          let rngNextSent = selection.getNextTextRangeOrNullObject([" ", "." , "," , ";", "!", "?", ":"], false);
          selection = selection.expandToOrNullObject(rngNextSent.getRange("Start"));
          await context.sync();
          
          // selezione all'indietro   
          if(isLetterOrNumber(charBefore)){
            let paragraph = selection.paragraphs.getFirst();
            paragraph.load("text");
            await context.sync();
  
            let rangeToSelect = paragraph.getRange("Start").expandTo(selection);
            let textBeforeSelection = rangeToSelect.getTextRanges([" ", ".", ",", ";"]);
            textBeforeSelection.load("items");
            await context.sync();
            let lastItem = textBeforeSelection.items[textBeforeSelection.items.length - spaceCount];
            let rangeToExpand = lastItem.getRange("Start");
            selection = selection.expandToOrNullObject(rangeToExpand);
            await context.sync();
          } 
          selection.load();
          await context.sync();
      }
        const searchResults = range.search(selection.text, { matchCase: false, matchWholeWord: false });
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