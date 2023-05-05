import * as React from 'react';
import { useState, useEffect } from 'react';
import FormControlLabel from '@mui/material/FormControlLabel';
import Checkbox from '@mui/material/Checkbox';

export const AllInstances  = ({styleGSG, fontStyle, buttonStyle, firstOccurence, first, expandedText, entitiesStyle, styleOtherEntities, styleInformative}) => {
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

        if(expandedText != selection.text && selection.text != ""){  
          const startIndex = expandedText.indexOf(selection.text);
          const charBefore = expandedText[startIndex - 1];
          let text = selection.text;
          let spaceCount = text.split(" ").length;
          
          //selezione in avanti fino ad uno di quei caratteri
          const nextCharRanges = selection.getTextRanges([" ", ".", ",", ";", "!", "?", ":"], true);
          nextCharRanges.load("items");
          await context.sync();
          if (nextCharRanges.items.length > 0) {
            for(let i = 0; i < spaceCount; i++){
                console.log(nextCharRanges.items[i].text)
                selection = selection.expandTo(nextCharRanges.items[i]);
            }
          }
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
        }

        await context.sync();

        const searchResults = range.search(selection.text, { matchCase: false, matchWholeWord: false });
        searchResults.load("items");
        await context.sync();

        const occurrences = searchResults.items;
        occurrences.forEach(async(occurrence) => {
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
            case "IntenseReference":
              occurrence.styleBuiltIn = first == "IntenseReference" ? "Normal" : "IntenseReference";
              break;
            case "Heading6":
              occurrence.styleBuiltIn = first == "Heading6" ? "Normal" : "Heading6";
              break;
            case "IntenseEmphasis":
              occurrence.styleBuiltIn = first == "IntenseEmphasis" ? "Normal" : "IntenseEmphasis";
              break;
            case "Normal":
              occurrence.styleBuiltIn = "Normal"
              break;
            default:
              break;
          }

          // stili entità principali
            switch (entitiesStyle) {
                case "Date":
                    occurrence.style = "Data1";
                    break;
                case "Organization":
                    occurrence.style = "Organization";
                    break;
                case "Person":
                    occurrence.style = "Person";
                    break;
                case "Location":
                    occurrence.style = "Location";
                    break;
                case "Time":
                    occurrence.style = "Time";
                    break;
                default:
                    break;
            }

            // stili altre entità
            switch(styleOtherEntities) {
                case "object":
                    occurrence.style = "Object"
                    break;
                case "event":
                    occurrence.style = "Event";
                    break;
                case "process":
                    occurrence.style = "Process";
                    break;
                case "role":
                    occurrence.style = "Role";
                    break;
                case "term":
                    occurrence.style = "Term";
                    break;
                case "quantity":
                    occurrence.style = "Quantity";
                    break;
                default:
                    break;
            }
            
            // stili informative entities
            switch(styleInformative) {
                case "docTitle":
                    occurrence.style = "docTitle"
                    break;
                case "docNumber":
                    occurrence.style = "docNumber"
                    break;
                case "docProponent":
                    occurrence.style = "docProponent"
                    break;
                case "docDate":
                    occurrence.style = "docDate"
                    break;
                case "session":
                    occurrence.style = "session"
                    break;
                case "shortTitle":
                    occurrence.style = "shortTitle"
                    break;
                case "docAuthority":
                    occurrence.style = "docAuthority"
                    break;
                case "docPurpose":
                    occurrence.style = "docPurpose"
                    break;
                case "docCommittee":
                    occurrence.style = "docCommittee"
                    break;
                case "docIntroducer":
                    occurrence.style = "docIntroducer"
                    break;
                case "docStage":
                    occurrence.style = "docStage"
                    break;
                case "docStatus":
                    occurrence.style = "docStatus"
                    break;
                case "docJurisdiction":
                    occurrence.style = "docJurisdiction"
                    break;
                case "docketNumber":
                    occurrence.style = "docketNumber"
                    break;
                default:
                    break;
            } 
            

          // stili GSG
          switch(styleGSG){
            case 1:
                occurrence.style = "GSG"
                break;
            case 2:
                occurrence.style = "GSG2";
                break;
            case 3:
                occurrence.style = "GSG3"
                break;
            case 4:
                occurrence.style = "GSG4"
                break;
            case 5:
                occurrence.style = "GSG5"
                break;
            case 6:
                occurrence.style = "GSG6"
                break;
            case 7:
                occurrence.style = "GSG7"
                break;
            case 8:
                occurrence.style = "GSG8"
                break;
            case 9:
                occurrence.style = "GSG9"
                break;
            case 10:
                occurrence.style = "GSG10"
                break;
            case 11:
                occurrence.style = "GSG11"
                break;
            case 12:
                occurrence.style = "GSG12"
                break;
            case 13:
                occurrence.style = "GSG13"
                break;
            case 14:
                occurrence.style = "GSG14"
                break;
            case 15:
                occurrence.style = "GSG15"
                break;
            case 16:
                occurrence.style = "GSG15"
                break;
            case 17:
                occurrence.style = "GSG16"
                break;
            default:
                break;
        }
        await context.sync();
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