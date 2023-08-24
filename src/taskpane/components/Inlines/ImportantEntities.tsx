import * as React from 'react';
import { useState, useEffect } from 'react';
import Grid from '@mui/material/Grid';
import IconButton from '@mui/material/IconButton';
import CalendarMonthIcon from '@mui/icons-material/CalendarMonth';
import FolderOpenIcon from '@mui/icons-material/FolderOpen';
import PersonIcon from '@mui/icons-material/Person';
import LocationOnIcon from '@mui/icons-material/LocationOn';

export const ImportantEntities = ({ setDis, expandedText }) => {
    let entity, dialog;
    const isLetterOrNumber = (char) => {
        if (typeof char === "undefined") {
            return false;
        } else {
            return /^[a-zA-Z0-9]+$/.test(char);
        }
    }

    const updateStyleBuiltIn = async (context, messageFromDialog) => {
        const selection = context.document.getSelection();
        selection.load("style, text");
        await context.sync();
        const selectedText = selection.text;
        const NAMESPACE_URI = "prova";
        const uniqueId = Date.now();
        const xmlData = `<root xmlns="${NAMESPACE_URI}"><data id="${uniqueId}" style="${entity}" text="${selectedText.toLowerCase()}">${JSON.stringify(messageFromDialog)}</data></root>`;

        switch (entity) {
            case "Date":
                selection.style = "Data1";
                break;
            case "Organization":
                selection.style = "Organization";
                break
            case "Person":
                selection.style = "Person";
                break;
            case "Location":
                selection.style = "Location";
                break;
            default:
                break;
        }
        selection.select("end");
        const range = context.document.body.getRange();
        await context.sync();
        const searchResults = range.search(selection.text, { matchCase: false, matchWholeWord: false });
        searchResults.load("items");
        await context.sync();
        const occurrences = searchResults.items;
        occurrences.forEach(async(occurrence) => {
            switch (entity) {
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
                default:
                    break;
            }
        });
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

        // Inserisci la nuova informazione aggiunta 
        Office.context.document.customXmlParts.addAsync(xmlData, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Dati personalizzati aggiunti con successo");
            } else {
                console.error("Errore durante l'aggiunta dei dati personalizzati");
            }
        });
        await context.sync();
    };

    const processMessage = async (arg) => {
        const messageFromDialog = JSON.parse(arg.message);
        dialog.close();

        await Word.run(async (context) => {
            await updateStyleBuiltIn(context, messageFromDialog);
        });
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
                    dialogUrl += 'date.html';
                    break;
                case "Organization":
                    dialogUrl += 'organization.html';
                    break
                case "Person":
                    dialogUrl += 'person.html';
                    break;
                case "Location":
                    dialogUrl += 'location.html';
                    break;
                default:
                    break;
            }
            if (entities === "Date") {
                Office.context.ui.displayDialogAsync(dialogUrl, {
                    height: 70,
                    width: 45,
                    displayInIframe: true,
                },
                    function (asyncResult) {
                        dialog = asyncResult.value;
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                    });
            } else {
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
                    <div style={{ fontSize: '10px', position: 'relative', left: '12px', color: setDis ? 'grey' : 'red' }}>Date</div>
                </Grid>
                <Grid item xs={3}>
                    <IconButton disabled={setDis} color="success" onClick={() => updateStyle('Organization')}>
                        <FolderOpenIcon fontSize="large" />
                    </IconButton>
                    <div style={{ fontSize: '10px', position: 'relative', right: '6px', color: setDis ? 'grey' : 'green' }}>Organization</div>
                </Grid>
                <Grid item xs={3}>
                    <IconButton disabled={setDis} color="info" onClick={() => updateStyle('Person')}>
                        <PersonIcon fontSize="large" />
                    </IconButton>
                    <div style={{ fontSize: '10px', position: 'relative', left: '10px', color: setDis ? 'grey' : 'blue' }}>Person</div>
                </Grid>
                <Grid item xs={3}>
                    <IconButton disabled={setDis} color="warning" onClick={() => updateStyle('Location')}>
                        <LocationOnIcon fontSize="large" />
                    </IconButton>
                    <div style={{ fontSize: '10px', position: 'relative', left: '7px', color: setDis ? 'grey' : 'orange' }}>Location</div>
                </Grid>
            </Grid>
        </div>
    )
}