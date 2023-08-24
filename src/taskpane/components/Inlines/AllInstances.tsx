import * as React from 'react';
import { useState, useEffect } from 'react';
import FormControlLabel from '@mui/material/FormControlLabel';
import Checkbox from '@mui/material/Checkbox';

export const AllInstances  = ({styleGSG, buttonStyle, firstOccurence, expandedText}) => {
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

    // per prima cosa si controlla se la checkbox "All Instances" è selezionata
    if (allInstances) {
        const applyFormatting = async (context) => {
            const range = context.document.body.getRange();
            let selection = context.document.getSelection();
            selection.load();
            await context.sync();

            // se il testo espanso è diverso da quello selezionato, vuol dire che la checkbox è stata selezionata e quindi bisogna aggiorrnare la selezione
            if(expandedText != selection.text && selection.text != ""){  
                const startIndex = expandedText.indexOf(selection.text);
                const charBefore = expandedText[startIndex - 1];
                let text = selection.text;
                let spaceCount = text.split(" ").length;
                
                //selezione in avanti fino a che non trova uno di questi caratteri
                const nextCharRanges = selection.getTextRanges([" ", ".", ",", ";", "!", "?", ":"], true);
                nextCharRanges.load("items");
                await context.sync();
                if (nextCharRanges.items.length > 0) {
                    for(let i = 0; i < spaceCount; i++){
                        // espande la selezione per ogni carattere che ha trovato fino ad uno di quelli dichiarati
                        selection = selection.expandTo(nextCharRanges.items[i]);
                    }
                }
                await context.sync();
                
                // per prima cosa si controlla se il carattere subito prima della selezione fa parte del testo che va selezionato   
                if(isLetterOrNumber(charBefore)){
                    // prendiamo il paragrafo che include il testo selezionato
                    let paragraph = selection.paragraphs.getFirst();
                    paragraph.load("text");
                    await context.sync();
                    
                    // a questo punto dividiamo il paragrafo trovato in diverse parti ed ogni parte viene separato da uno di quei caratteri
                    let rangeToSelect = paragraph.getRange("Start").expandTo(selection);
                    let textBeforeSelection = rangeToSelect.getTextRanges([" ", ".", ",", ";"], true);
                    textBeforeSelection.load("items");
                    await context.sync();

                    // dopo aver diviso il paragrafo, prendiamo l'ultima parte, che equivale al testo subito precedente al testo selezionato
                    let lastItem = textBeforeSelection.items[textBeforeSelection.items.length - spaceCount];
                    
                    // espandiamo la selezione a partire dall'inizio dell'ultima parte del paragrafo
                    let rangeToExpand = lastItem.getRange("Start");
                    selection = selection.expandToOrNullObject(rangeToExpand);
                    await context.sync();
                } 
                // carichiamo la nuova selezione
                selection.load();
            }

            await context.sync();

            // serve per trovare tutte le occorrenze del testo selezionato, dopo aver aggiornato eventualmente la selezione
            const searchResults = range.search(selection.text, { matchCase: false, matchWholeWord: false });
            searchResults.load("items");
            await context.sync();
            const occurrences = searchResults.items;

            occurrences.forEach(async(occurrence) => {
                // formattazione testo di tutte le occorrenze
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

                // stili GSG per tutte le occorrenze
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