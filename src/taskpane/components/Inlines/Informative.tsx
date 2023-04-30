import * as React from 'react';
import { useState, useEffect } from 'react';
import Grid from '@mui/material/Grid';
import InputLabel from '@mui/material/InputLabel';
import MenuItem from '@mui/material/MenuItem';
import FormControl from '@mui/material/FormControl';
import Select, { SelectChangeEvent } from '@mui/material/Select';

export const Informative = ({expandedText, onInformativeStyle}) => {
    const [docType, setDocType] = useState("");

    const isLetterOrNumber = (char) => {
        if (typeof char === "undefined") {
            return false;
        }else{
            return /^[a-zA-Z0-9]+$/.test(char);
        }
    }

    const handleChangeDocType = async (event: SelectChangeEvent) => {
        setDocType(event.target.value)
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
                selection.load("styleBuiltIn");
                selection.font.load("color")
                await context.sync();
            }

            switch(event.target.value) {
                case "docTitle":
                    if (Office.context.platform === Office.PlatformType.OfficeOnline){
                        selection.font.color = "red";
                        selection.font.bold = true;
                    }else{
                        selection.font.underline = "DashLineLong"
                    }
                    break;
                case "docNumber":
                    if (Office.context.platform === Office.PlatformType.OfficeOnline){
                        selection.font.color = "green";
                        selection.font.bold = true;
                    }else{
                        selection.font.underline = "DotDashLine"
                    }
                    break;
                case "docProponent":
                    if (Office.context.platform === Office.PlatformType.OfficeOnline){
                        selection.font.color = "blue";
                        selection.font.bold = true;
                    }else{
                        selection.font.underline = "Double"
                    }
                    break;
                case "docDate":
                    if (Office.context.platform === Office.PlatformType.OfficeOnline){
                        selection.font.color = "purple";
                        selection.font.bold = true;
                    }else{
                        selection.font.underline = "Thick"
                    }
                    break;
                case "session":
                    if (Office.context.platform === Office.PlatformType.OfficeOnline){
                        selection.font.color = "yellow";
                        selection.font.bold = true;
                    }else{
                        selection.font.underline = "TwoDotDashLine"
                    }
                    break;
                case "shortTitle":
                    if (Office.context.platform === Office.PlatformType.OfficeOnline){
                        selection.font.color = "orange";
                        selection.font.bold = true;
                    }else{
                        selection.font.underline = "Wave"
                    }
                    break;
                case "docAuthority":
                    if (Office.context.platform === Office.PlatformType.OfficeOnline){
                        selection.font.color = "brown";
                        selection.font.bold = true;
                    }else{
                        selection.font.underline = "Word"
                    }
                    break;
                case "docPurpose":
                    if (Office.context.platform === Office.PlatformType.OfficeOnline){
                        selection.font.color = "pink";
                        selection.font.bold = true;
                    }else{
                        selection.font.underline = "TwoDotDashLineHeavy"
                    }
                    break;
                case "docCommittee":
                    if (Office.context.platform === Office.PlatformType.OfficeOnline){
                        selection.font.color = "lightblue";
                        selection.font.bold = true;
                    }else{
                        selection.font.underline = "DottedHeavy"
                    }
                    break;
                case "docIntroducer":
                    if (Office.context.platform === Office.PlatformType.OfficeOnline){
                        selection.font.color = "cyan";
                        selection.font.bold = true;
                    }else{
                        selection.font.underline = "WaveDouble"
                    }
                    break;
                case "docStage":
                    if (Office.context.platform === Office.PlatformType.OfficeOnline){
                        selection.font.color = "#c2bd34";
                        selection.font.bold = true;
                    }else{
                        selection.font.underline = "DashLineLongHeavy"
                    }
                    break;
                case "docStatus":
                    if (Office.context.platform === Office.PlatformType.OfficeOnline){
                        selection.font.color = "#b0f5c5";
                        selection.font.bold = true;
                    }else{
                        selection.font.underline = "WaveHeavy"
                    }
                    break;
                case "docJurisdiction":
                    if (Office.context.platform === Office.PlatformType.OfficeOnline){
                        selection.font.color = "#26ad89";
                        selection.font.bold = true;
                    }else{
                        selection.font.underline = "Dotted"
                    }
                    break;
                case "docketNumber":
                    if (Office.context.platform === Office.PlatformType.OfficeOnline){
                        selection.font.color = "#d6fa89";
                        selection.font.bold = true;
                    }else{
                        selection.font.underline = "Hidden"
                    }
                    break;
                default:
                    if (Office.context.platform === Office.PlatformType.OfficeOnline){
                        selection.font.color = "black";
                        selection.font.bold = false;
                    }else{
                        selection.font.underline = "None"
                    }
                    break;
            }

            onInformativeStyle(event.target.value)
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