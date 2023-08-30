// licenza d'uso riservata ad Andrea Napoli e all'università si Bologna
import * as React from 'react';
import { useState, useEffect } from 'react';
import Grid from '@mui/material/Grid';
import InputLabel from '@mui/material/InputLabel';
import MenuItem from '@mui/material/MenuItem';
import FormControl from '@mui/material/FormControl';
import Select, { SelectChangeEvent } from '@mui/material/Select';

export const Informative = ({ setDis, expandedText }) => {
    const [docType, setDocType] = useState("");

    const isLetterOrNumber = (char) => {
        if (typeof char === "undefined") {
            return false;
        }else{
            return /^[a-zA-Z0-9]+$/.test(char);
        }
    }

    const deleteInformation = async (context, NAMESPACE_URI, selectedText) => {
        // Elimina informazione attuale
        Office.context.document.customXmlParts.getByNamespaceAsync(NAMESPACE_URI, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const xmlParts = result.value;
                for (const xmlPart of xmlParts) {
                    xmlPart.getXmlAsync(asyncResult => {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            const xmlData = asyncResult.value;
                            if (xmlData.includes(`text="${selectedText.toLowerCase()}"`)) {
                                xmlPart.deleteAsync();
                            }
                        } else {
                            console.error("Errore nel recupero dei contenuti personalizzati");
                        }

                    });
                }
            } else {
                console.error("Errore nel recupero dei contenuti personalizzati");
            }
        });

        await context.sync();
    }

    const handleChangeDocType = async (event: SelectChangeEvent) => {
        setDocType(event.target.value)
        let removed = false;
        let char_remove;
        await Word.run(async (context) => {
            let selection = context.document.getSelection();
            selection.load("paragraphs, text, styleBuiltIn, font");
            await context.sync();
            let paragraphCount = selection.paragraphs.items.length; 
            let emptyParagraph = 0;

            for(let i = 0; i < selection.paragraphs.items.length; i++) { // se nella selezione includo anche i paragrafi, non funziona perfettamente
                if(selection.paragraphs.items[i].text == ""){
                emptyParagraph ++;
                }
            }

            // stessa funzione di espansione
            if(expandedText != selection.text && selection.text != ""){  // ho aggiunto la seconda condizine in quanto se non avevo del testo selezionato, appena premevo i bottini di stili mi evidenzia l'ultima parola
                const startIndex = expandedText.indexOf(selection.text);
                const charBefore = expandedText[startIndex - 1];
                
                let text = selection.text;
                let spaceCount = text.split(" ").length;
                //selezione in avanti fino ad uno di quei caratteri
                const nextCharRanges = selection.getTextRanges([" ", ".", ",", ";", "!", "?", ":", "\n", "\r"], true);
                nextCharRanges.load("items");
                
                await context.sync();
                
                if (nextCharRanges.items.length > 0) {
                    if(paragraphCount>1){ // se più paragraphi sono compresi, andare a capo lo prende come una parola e quindi spaceCount va incrementato con il numero di paragrafi -1, però bisogna togliere i paragrafi vuoti
                        spaceCount = spaceCount + paragraphCount - 1 - emptyParagraph;
                    }
                    for(let i = 0; i < spaceCount; i++){
                        selection = selection.expandTo(nextCharRanges.items[i]);
                    }
                }
                selection.load("text");
                await context.sync();
                const punctuationMarks = [" ", ".", ",", ";", "!", "?", ":", "\n", "\r"];
                if(punctuationMarks.includes(selection.text[selection.text.length - 1])){
                    removed = true;
                    char_remove = selection.text[selection.text.length - 1];
                    let newText = selection.text.substring(0, selection.text.length-1);
                    selection.insertText(newText, "Replace");
                }
                await context.sync();
                
                // selezione all'indietro   
                if(isLetterOrNumber(charBefore)){
                  let paragraph = selection.paragraphs.getFirst();
                  paragraph.load("text");
                  await context.sync();
        
                  let rangeToSelect = paragraph.getRange("Start").expandTo(selection);
                  let textBeforeSelection = rangeToSelect.getTextRanges([" ", ".", ",", ";"], false);
                  textBeforeSelection.load("items");
                  await context.sync();
                  let lastItem = textBeforeSelection.items[textBeforeSelection.items.length - spaceCount];
                  let rangeToExpand = lastItem.getRange("Start");
                  selection = selection.expandToOrNullObject(rangeToExpand);
                  await context.sync();
                }
                selection.select();
            }

            selection.load("styleBuiltIn, text, style, font");
            await context.sync();

            const platform = Office.context.platform !== Office.PlatformType.OfficeOnline;
            
            switch(event.target.value) {
                case "docTitle":
                    platform ? selection.style = "docTitle" : selection.font.color = "#9ACD32";
                    break;
                case "docNumber":
                    platform ? selection.style = "docNumber" : selection.font.color = "#008B8B";
                    break;
                case "docProponent":
                    platform ? selection.style = "docProponent" : selection.font.color = "#00FFFF", selection.font.bold = true;
                    break;
                case "docDate":
                    platform ? selection.style = "docDate" : selection.font.color = "#AFEEEE", selection.font.bold = true;
                    break;
                case "session":
                    platform ? selection.style = "session" : selection.font.color = "#4682B4";
                    break;
                case "shortTitle":
                    platform ? selection.style = "shortTitle" : selection.font.color = "#00BFFF";
                    break;
                case "docAuthority":
                    platform ? selection.style = "docAuthority" : selection.font.color = "#0000FF";
                    break;
                case "docPurpose":
                    platform ? selection.style = "docPurpose" : selection.font.color = "#FFDEAD";
                    break;
                case "docCommittee":
                    platform ? selection.style = "docCommittee" : selection.font.color = "#F4A460";
                    break;
                case "docIntroducer":
                    platform ? selection.style = "docIntroducer" : selection.font.color = "#DAA520";
                    break;
                case "docStage":
                    platform ? selection.style = "docStage" : selection.font.color = "#696969";
                    break;
                case "docStatus":
                    platform ? selection.style = "docStatus" : selection.font.color = "#2F4F4F";
                    break;
                case "docJurisdiction":
                    platform ? selection.style = "docJurisdiction" : selection.font.color = "#00FA9A", selection.font.bold = true;
                    break;
                case "docketNumber":
                    platform ? selection.style = "docketNumber" : selection.font.color = "#7B68EE";
                    break;
                default:
                    selection.styleBuiltIn = "Normal"
                    break;
            }
            if(removed == true){
                selection.insertText(char_remove, "End");
                removed = false;
            }
            selection.select(Word.SelectionMode.end);
            const range = context.document.body.getRange();
            await context.sync();
            const searchResults = range.search(selection.text, { matchCase: false, matchWholeWord: false });
            searchResults.load("items");
            await context.sync();
            const occurrences = searchResults.items;
    
            occurrences.forEach(async (occurrence) => {
                switch(event.target.value) {
                    case "docTitle":
                        platform ? occurrence.style = "docTitle" : occurrence.font.color = "#9ACD32";
                        break;
                    case "docNumber":
                        platform ? occurrence.style = "docNumber" : occurrence.font.color = "#008B8B";
                        break;
                    case "docProponent":
                        platform ? occurrence.style = "docProponent" : occurrence.font.color = "#00FFFF", occurrence.font.bold = true;
                        break;
                    case "docDate":
                        platform ? occurrence.style = "docDate" : occurrence.font.color = "#AFEEEE", occurrence.font.bold = true;
                        break;
                    case "session":
                        platform ? occurrence.style = "session" : occurrence.font.color = "#4682B4";
                        break;
                    case "shortTitle":
                        platform ? occurrence.style = "shortTitle" : occurrence.font.color = "#00BFFF";
                        break;
                    case "docAuthority":
                        platform ? occurrence.style = "docAuthority" : occurrence.font.color = "#0000FF";
                        break;
                    case "docPurpose":
                        platform ? occurrence.style = "docPurpose" : occurrence.font.color = "#FFDEAD";
                        break;
                    case "docCommittee":
                        platform ? occurrence.style = "docCommittee" : occurrence.font.color = "#F4A460";
                        break;
                    case "docIntroducer":
                        platform ? occurrence.style = "docIntroducer" : occurrence.font.color = "#DAA520";
                        break;
                    case "docStage":
                        platform ? occurrence.style = "docStage" : occurrence.font.color = "#696969";
                        break;
                    case "docStatus":
                        platform ? occurrence.style = "docStatus" : occurrence.font.color = "#2F4F4F";
                        break;
                    case "docJurisdiction":
                        platform ? occurrence.style = "docJurisdiction" : occurrence.font.color = "#00FA9A", occurrence.font.bold = true;
                        break;
                    case "docketNumber":
                        platform ? occurrence.style = "docketNumber" : occurrence.font.color = "#7B68EE";
                        break;
                    default:
                        occurrence.styleBuiltIn = "Normal"
                        break;
                }
            });
        
            const NAMESPACE_URI = "prova";
            deleteInformation(context, NAMESPACE_URI, selection.text);
        });
    }
    

    return (
        <div>
            <Grid
              container
              direction="row"
              justifyContent="center"
              alignItems="flex-start"
              spacing={1}
              style={{marginTop: "5px"}}
            >
                <Grid item xs={6}>
                    <p style={{marginLeft: '18px', marginTop: '16px', fontSize: '17px'}}>Informative</p>
                </Grid>
                <Grid item xs={6}>
                    <FormControl 
                        disabled = {setDis}
                        sx={{ m: 1, minWidth: 120 }} 
                        size="small"
                        style={{position: 'relative', right: '18px'}}
                    >
                        <InputLabel id="demo-select-small">docType</InputLabel>
                        <Select
                            labelId="demo-select-small"
                            id="demo-select-small"
                            value={docType}
                            label="docType"
                            onChange={handleChangeDocType}
                        >
                            <MenuItem value="">
                                <em>None</em>
                            </MenuItem>
                            <MenuItem value="docTitle">docTitle</MenuItem>
                            <MenuItem value="docNumber">docNumber</MenuItem>
                            <MenuItem value="docProponent">docProponent</MenuItem>
                            <MenuItem value="docDate">docDate</MenuItem>
                            <MenuItem value="session">session</MenuItem>
                            <MenuItem value="shortTitle">shortTitle</MenuItem>
                            <MenuItem value="docAuthority">docAuthority</MenuItem>
                            <MenuItem value="docPurpose">docPurpose</MenuItem>
                            <MenuItem value="docCommittee">docCommittee</MenuItem>
                            <MenuItem value="docIntroducer">docIntroducer</MenuItem>
                            <MenuItem value="docStage">docStage</MenuItem>
                            <MenuItem value="docStatus">docStatus</MenuItem>
                            <MenuItem value="docJurisdiction">docJurisdiction</MenuItem>
                            <MenuItem value="docketNumber">docketNumber</MenuItem>
                        </Select>
                    </FormControl>
                </Grid>
            </Grid>
        </div>
    )
}