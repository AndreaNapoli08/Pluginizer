// licenza d'uso riservata ad Andrea Napoli e all'università si Bologna
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

    const getInformation = async (NAMESPACE_URI, selectedText) => {
        return new Promise(async (resolve) => {
            Office.context.document.customXmlParts.getByNamespaceAsync(NAMESPACE_URI, async (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const xmlParts = result.value;
                    for (const xmlPart of xmlParts) {
                        await xmlPart.getXmlAsync(asyncResult => {    // questa istruzione non aspetta il completamento di ciascuna chiamata
                            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                                const xmlData = asyncResult.value;
                                if (xmlData.includes(`text="${selectedText.toLowerCase()}"`)) {
                                    const parser = new DOMParser();
                                    const xmlDoc = parser.parseFromString(xmlData, "text/xml");
                                    const dataElement = xmlDoc.querySelector(`data[text="${selectedText.toLowerCase()}"]`);
                                    if (dataElement) {
                                        let jsonData = JSON.parse(dataElement.textContent);
                                        switch (jsonData.entity) {
                                            case "date":
                                                resolve({
                                                    day: jsonData.day,
                                                    month: jsonData.month,
                                                    year: jsonData.year,
                                                    time: jsonData.time
                                                });
                                                break;
                                            case "organization":
                                                resolve({
                                                    organization: jsonData.organization
                                                });
                                                break;
                                            case "person":
                                                resolve({
                                                    person: jsonData.person
                                                });
                                                break;
                                            case "location":
                                                resolve({
                                                    location: jsonData.location
                                                });
                                                break;
                                            default:
                                                resolve({});
                                                break;
                                        }
                                    } else {
                                        resolve({});
                                    }
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
        });
    }

    const insertInformation = async (context, xmlData) => {
        // inserimento nuova informazione
        Office.context.document.customXmlParts.addAsync(xmlData, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Dati personalizzati aggiunti con successo");
            } else {
                console.error("Errore durante l'aggiunta dei dati personalizzati");
            }
        });
        await context.sync();
    }

    const updateStyleBuiltIn = async (context, messageFromDialog) => {
        const selection = context.document.getSelection();
        selection.load("style, text, styleBuiltIn, font");
        await context.sync();
        const selectedText = selection.text;
        const NAMESPACE_URI = "prova";
        const uniqueId = Date.now();
        const xmlData = `<root xmlns="${NAMESPACE_URI}"><data id="${uniqueId}" style="${entity}" text="${selectedText.toLowerCase()}">${JSON.stringify(messageFromDialog)}</data></root>`;
        const platform = Office.context.platform !== Office.PlatformType.OfficeOnline;

        switch (entity) {
            case "Date":
                platform ? selection.style = "Data1" : selection.font.color = "red", selection.font.bold = true;   
                break;
            case "Organization":
                platform ? selection.style = "Organization" : selection.font.color = "green", selection.font.bold = true;   
                break
            case "Person":
                platform ? selection.style = "Person" : selection.font.color = "blue", selection.font.bold = true; 
                break;
            case "Location":
                platform ? selection.style = "Location" : selection.font.color = "orange", selection.font.bold = true; 
                break;
            default:
                break;
        }
        selection.select(Word.SelectionMode.end);
        const range = context.document.body.getRange();
        await context.sync();
        const searchResults = range.search(selection.text, { matchCase: false, matchWholeWord: false });
        searchResults.load("items");
        await context.sync();
        const occurrences = searchResults.items;
        occurrences.forEach(async (occurrence) => {
            switch (entity) {
                case "Date":
                    platform ? occurrence.style = "Data1" : occurrence.font.color = "red", occurrence.font.bold = true; 
                    break;
                case "Organization":
                    platform ? occurrence.style = "Organization" : occurrence.font.color = "green", occurrence.font.bold = true;   
                    break;
                case "Person":
                    platform ? occurrence.style = "Person" : occurrence.font.color = "blue", occurrence.font.bold = true;
                    break;
                case "Location":
                    platform ? occurrence.style = "Location" : occurrence.font.color = "orange", occurrence.font.bold = true;
                    break;
                default:
                    break;
            }
        });
        // Elimina informazione attuale
        deleteInformation(context, NAMESPACE_URI, selectedText);

        // Inserisci la nuova informazione aggiunta 
        insertInformation(context, xmlData);
    };

    const processMessage = async (arg) => {
        const messageFromDialog = JSON.parse(arg.message);
        dialog.close();

        await Word.run(async (context) => {
            await updateStyleBuiltIn(context, messageFromDialog);
        });
    }

    const updateStyle = async (entities) => {
        let removed = false;
        let char_remove;
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
                if (isLetterOrNumber(charBefore)) {
                    let paragraph = selection.paragraphs.getFirst();
                    paragraph.load("text");
                    await context.sync();

                    let rangeToSelect = paragraph.getRange("Start").expandTo(selection);
                    let textBeforeSelection = rangeToSelect.getTextRanges([" ", ".", ",", ";"], true);
                    textBeforeSelection.load("items");
                    await context.sync();
                    let lastItem = textBeforeSelection.items[textBeforeSelection.items.length - spaceCount];
                    let rangeToExpand = lastItem.getRange("Start");
                    selection = selection.expandToOrNullObject(rangeToExpand);
                    await context.sync();
                }
                selection.select();
                if(removed == true){
                    selection.insertText(char_remove, "End");
                    removed = false;
                }
            }

            selection.load("styleBuiltIn, style, text, font");
            await context.sync();
            let dialogUrl = 'https://localhost:3000/assets/';
            switch (entities) {
                case "Date":
                    dialogUrl += 'date.html';
                    if (selection.style === "Data1" || selection.font.color === "#FF0000") {
                        const information = await getInformation("prova", selection.text);
                        const informationString = JSON.stringify(information);
                        dialogUrl += `?information=${encodeURIComponent(informationString)}`;
                    }
                    break;
                case "Organization":
                    dialogUrl += 'organization.html';
                    console.log(selection.font.color)
                    if(selection.style === "Organization" || selection.font.color == "#008000"){
                        const information = await getInformation("prova", selection.text);
                        const informationString = JSON.stringify(information);
                        dialogUrl += `?information=${encodeURIComponent(informationString)}`;
                    }
                    break
                case "Person":
                    dialogUrl += 'person.html';
                    console.log(selection.font.color)
                    if(selection.style === "Person" || selection.font.color == "#0000FF"){
                        const information = await getInformation("prova", selection.text);
                        const informationString = JSON.stringify(information);
                        dialogUrl += `?information=${encodeURIComponent(informationString)}`;
                    }
                    break;
                case "Location":
                    dialogUrl += 'location.html';
                    console.log(selection.font.color)
                    if(selection.style === "Location" || selection.font.color == "#FFA500"){
                        const information = await getInformation("prova", selection.text);
                        const informationString = JSON.stringify(information);
                        dialogUrl += `?information=${encodeURIComponent(informationString)}`;
                    }
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