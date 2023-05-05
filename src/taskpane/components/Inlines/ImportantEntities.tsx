import * as React from 'react';
import { useState, useEffect } from 'react';
import Grid from '@mui/material/Grid';
import IconButton from '@mui/material/IconButton';
import CalendarMonthIcon from '@mui/icons-material/CalendarMonth';
import FolderOpenIcon from '@mui/icons-material/FolderOpen';
import PersonIcon from '@mui/icons-material/Person';
import LocationOnIcon from '@mui/icons-material/LocationOn';
import AccessTimeIcon from '@mui/icons-material/AccessTime';

export const ImportantEntities = ({expandedText, onEntitiesStyle}) => {
    const isLetterOrNumber = (char) => {
        if (typeof char === "undefined") {
            return false;
        }else{
            return /^[a-zA-Z0-9]+$/.test(char);
        }
    }  

    const updateStyle = async (entities) => { 
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
            if(expandedText != selection.text && selection.text != ""){
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
                selection.load("styleBuiltIn, style");
                selection.font.load("color")
                await context.sync();
            }
            
            switch(entities) {
                case "Date" :
                    selection.style = "Data1"
                    break;
                case "Organization" :
                    selection.style = "Organization"
                    break
                case "Person":
                    selection.style = "Person"
                    break;
                case "Location":
                    selection.style = "Location"
                    break;
                case "Time":
                    selection.style = "Time"
                    break;
                default:
                    break;
            }

            await context.sync();
            onEntitiesStyle(entities)
            onEntitiesStyle("")
        });
    }

  return (
    <div>
        <div style={{display: 'flex', justifyContent: 'center', marginBottom: '5px', fontSize: '20px'}}>
                Entities    
        </div>
        <Grid
            container
            direction="row"
            justifyContent="center"
            alignItems="flex-start"
            spacing={2}
        >
            <Grid item xs={2.4}>
            <IconButton color="error" onClick={() => updateStyle('Date')}>
                <CalendarMonthIcon fontSize="large" />
            </IconButton>
            <div style={{fontSize: '10px', position: 'relative', left: '12px', color: 'red'}}>Date</div>
            </Grid>
            <Grid item xs={2.4}>
            <IconButton color="success" onClick={() => updateStyle('Organization')}>
                <FolderOpenIcon fontSize="large" />
            </IconButton>
            <div style={{fontSize: '10px', position: 'relative', right: '6px', color: 'green'}}>Organization</div>
            </Grid>
            <Grid item xs={2.4}>
            <IconButton color="info" onClick={() => updateStyle('Person')}>
                <PersonIcon fontSize="large" />
            </IconButton>
            <div style={{fontSize: '10px', position: 'relative', left: '10px', color: 'blue'}}>Person</div>
            </Grid>
            <Grid item xs={2.4}>
            <IconButton onClick={() => updateStyle('Location')} style={{color: 'orange'}}>
                <LocationOnIcon fontSize="large" />
            </IconButton>
            <div style={{fontSize: '10px', position: 'relative', left: '7px', color: 'orange'}}>Location</div>
            </Grid>
            <Grid item xs={2.4}>
            <IconButton onClick={() => updateStyle('Time')} style={{color: 'purple'}}>
                <AccessTimeIcon fontSize="large" />
            </IconButton>
            <div style={{fontSize: '10px', position: 'relative', left: '12px', color: 'purple'}}>Time</div>
            </Grid>
        </Grid>
    </div>
  )
}