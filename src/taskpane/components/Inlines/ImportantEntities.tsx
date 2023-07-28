import * as React from 'react';
import { useState, useEffect } from 'react';
import Grid from '@mui/material/Grid';
import IconButton from '@mui/material/IconButton';
import CalendarMonthIcon from '@mui/icons-material/CalendarMonth';
import FolderOpenIcon from '@mui/icons-material/FolderOpen';
import PersonIcon from '@mui/icons-material/Person';
import LocationOnIcon from '@mui/icons-material/LocationOn';

export const ImportantEntities = ({ info, setDis, expandedText, onEntitiesStyle }) => {
    let message, entity, dialog;
    const isLetterOrNumber = (char) => {
        if (typeof char === "undefined") {
            return false;
        } else {
            return /^[a-zA-Z0-9]+$/.test(char);
        }
    }

    const processMessage = (arg) => {
        const messageFromDialog = JSON.parse(arg.message);
        let val;
        switch (entity) {
            case "Date":
                val = messageFromDialog.day + ' ' + messageFromDialog.month + ' ' + messageFromDialog.year + ', ' + messageFromDialog.time;
                break;
            case "Organization":
                val = messageFromDialog.organization;
                break
            case "Person":
                val = messageFromDialog.person;
                break;
            case "Location":
                val = messageFromDialog.location;
                break;
            default:
                break;
        }
        info(message + " " + val);
        dialog.close();
    }

    const updateStyle = async (entities) => {
        await Word.run(async (context) => {
            entity = entities;
            let selection = context.document.getSelection();
            selection.load("paragraphs, text, styleBuiltIn, font");
            await context.sync();
            let paragraphCount = selection.paragraphs.items.length;
            let emptyParagraph = 0;
            for (let i = 0; i < selection.paragraphs.items.length; i++) { // se nella selezione includo anche i paragrafi vuoti, non funziona perfettamente
                if (selection.paragraphs.items[i].text == "") {
                    emptyParagraph++;
                }
            }

            // stessa funzione di espansione
            if (expandedText != selection.text && selection.text != "") {
                const startIndex = expandedText.indexOf(selection.text);
                const charBefore = expandedText[startIndex - 1];

                let text = selection.text;
                let spaceCount = text.split(" ").length;
                //selezione in avanti fino ad uno di quei caratteri
                const nextCharRanges = selection.getTextRanges([" ", ".", ",", ";", "!", "?", ":", "\n", "\r"], true);
                nextCharRanges.load("items");

                await context.sync();

                if (nextCharRanges.items.length > 0) {
                    if (paragraphCount > 1) { // se più paragraphi sono compresi, andare a capo lo prende come una parola e quindi spaceCount va incrementato con il numero di paragrafi -1, però bisogna togliere i paragrafi vuoti
                        spaceCount = spaceCount + paragraphCount - 1 - emptyParagraph;
                    }
                    for (let i = 0; i < spaceCount; i++) {
                        selection = selection.expandTo(nextCharRanges.items[i]);
                    }
                }
                await context.sync();

                // selezione all'indietro   
                if (isLetterOrNumber(charBefore)) {
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
            
            let dialogUrl = 'https://localhost:3000/assets/';
            switch (entities) {
                case "Date":
                    selection.style = "Data1"
                    dialogUrl += 'date.html';
                    message = "value of type Date with this characteristics: ";
                    break;
                case "Organization":
                    selection.style = "Organization";
                    dialogUrl += 'organization.html';
                    message = "value of type Organization with this characteristics: ";
                    break
                case "Person":
                    selection.style = "Person"
                    dialogUrl += 'person.html';
                    message = "value of type Person with this characteristics: ";
                    break;
                case "Location":
                    selection.style = "Location"
                    dialogUrl += 'location.html';
                    message = "value of type Location with this characteristics: ";
                    break;
                default:
                    break;
            }
            if(entities === "Date"){
                Office.context.ui.displayDialogAsync(dialogUrl, {
                    height: 70,
                    width: 45,
                    displayInIframe: true,
                }, 
                function (asyncResult) {
                    dialog = asyncResult.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                });
            }else{
                Office.context.ui.displayDialogAsync(dialogUrl, {
                    height: 50,
                    width: 20,
                    displayInIframe: true,
                },
                function (asyncResult) {
                    dialog = asyncResult.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                });
            }
            await context.sync();
            // passiamo al componente padre l'entità che l'utente ha scelto
            onEntitiesStyle(entities)
            onEntitiesStyle("")
        });
    }

    return (
        <div>
            <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '5px', fontSize: '20px' }}>
                Entities
            </div>
            <Grid
                container
                direction="row"
                justifyContent="center"
                alignItems="flex-start"
                spacing={2}
            >
                <Grid item xs={3}>
                    <IconButton disabled={setDis} color="error" onClick={() => updateStyle('Date')}>
                        <CalendarMonthIcon fontSize="large" />
                    </IconButton>
                    <div style={{ fontSize: '10px', position: 'relative', left: '12px', color: setDis ? 'grey' : 'red'}}>Date</div>
                </Grid>
                <Grid item xs={3}>
                    <IconButton disabled={setDis} color="success" onClick={() => updateStyle('Organization')}>
                        <FolderOpenIcon fontSize="large" />
                    </IconButton>
                    <div style={{ fontSize: '10px', position: 'relative', right: '6px', color: setDis ? 'grey' : 'green'}}>Organization</div>
                </Grid>
                <Grid item xs={3}>
                    <IconButton disabled={setDis} color="info" onClick={() => updateStyle('Person')}>
                        <PersonIcon fontSize="large" />
                    </IconButton>
                    <div style={{ fontSize: '10px', position: 'relative', left: '10px', color: setDis ? 'grey' : 'blue'}}>Person</div>
                </Grid>
                <Grid item xs={3}>
                    <IconButton disabled={setDis} color="warning" onClick={() => updateStyle('Location')}>
                        <LocationOnIcon fontSize="large" />
                    </IconButton>
                    <div style={{ fontSize: '10px', position: 'relative', left: '7px', color: setDis ? 'grey' : 'orange'}}>Location</div>
                </Grid>
            </Grid>
        </div>
    )
}