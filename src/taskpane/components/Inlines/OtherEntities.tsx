import * as React from 'react';
import { useState } from 'react';
import Grid from '@mui/material/Grid';
import InputLabel from '@mui/material/InputLabel';
import MenuItem from '@mui/material/MenuItem';
import FormControl from '@mui/material/FormControl';
import Select, { SelectChangeEvent } from '@mui/material/Select';

export const OtherEntities = ({expandedText, onOtherEntitiesStyle}) => {
    const [concept, setConcept] = useState('');
    
    const isLetterOrNumber = (char) => {
        if (typeof char === "undefined") {
            return false;
        }else{
            return /^[a-zA-Z0-9]+$/.test(char);
        }
    }

    const handleChangeConcept = async (event: SelectChangeEvent) => {
        setConcept(event.target.value);
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
            if(expandedText != selection.text  && selection.text != ""){
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
                selection.font.load("name")
                await context.sync();
            }

            switch(event.target.value) {
                case "object":
                        selection.font.name = "Consolas";
                    break;
                case "event":
                        selection.font.name = "DilleniaUPC";
                    break;
                case "process":
                        selection.font.name = "Franklin Gothic";
                    break;
                case "role":
                        selection.font.name = "Garamond";
                    break;
                case "term":
                        selection.font.name = "Gulim";
                    break;
                case "quantity":
                        selection.font.name = "KaiTi";
                    break;
                default:
                        selection.font.name = "Calibri";
                    break;
            }

            onOtherEntitiesStyle(selection.font.name)
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
                style={{marginTop: "10px"}}
            >
                <Grid item xs={6}>
                <p style={{marginLeft: '3px', marginTop: '16px', fontSize: '17px'}}>Other Entities</p>
                </Grid>
                <Grid item xs={6}>
                    <FormControl 
                        sx={{ m: 1, minWidth: 120 }} 
                        size="small"
                        style={{position: 'relative', right: '18px'}}
                    >
                        <InputLabel id="demo-select-small">concept</InputLabel>
                        <Select
                            labelId="demo-select-small"
                            id="demo-select-small"
                            value={concept}
                            label="concept"
                            onChange={handleChangeConcept}
                        >   
                            <MenuItem value="">
                                <em>None</em>
                            </MenuItem>
                            <MenuItem value="object">Object</MenuItem>
                            <MenuItem value="event">Event</MenuItem>
                            <MenuItem value="process">Process</MenuItem>
                            <MenuItem value="role">Role</MenuItem>
                            <MenuItem value="term">Term</MenuItem>
                            <MenuItem value="quantity">Quantity</MenuItem>
                        </Select>
                    </FormControl>
                </Grid>
            </Grid>
        </div>
    )
}