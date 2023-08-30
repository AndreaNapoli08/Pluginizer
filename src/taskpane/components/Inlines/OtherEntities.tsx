// licenza d'uso riservata ad Andrea Napoli e all'università si Bologna
import * as React from 'react';
import { useState } from 'react';
import Grid from '@mui/material/Grid';
import InputLabel from '@mui/material/InputLabel';
import MenuItem from '@mui/material/MenuItem';
import FormControl from '@mui/material/FormControl';
import Select, { SelectChangeEvent } from '@mui/material/Select';

export const OtherEntities = ({ setDis, expandedText }) => {
    let dialog, concept;

    const isLetterOrNumber = (char) => {
        if (typeof char === "undefined") {
            return false;
        } else {
            return /^[a-zA-Z0-9]+$/.test(char);
        }
    }

    const processEntities = async (arg) => {
        const messageFromDialog = JSON.parse(arg.message);
        if (messageFromDialog.URL != "") {
            dialog.close();
            await Word.run(async (context) => {
                await updateStyle(context, messageFromDialog);
            });
        }
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
                                        resolve({
                                            URL: jsonData.URL,
                                        });
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

    const updateStyle = async (context, messageFromDialog) => {
        let previousText;
        let selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        previousText = selection.text;
        selection.insertText(messageFromDialog.showAs, "Replace");
        selection.load("style, paragraphs, text, styleBuiltIn, font");
        await context.sync();
        let text = selection.text;
        let spaceCount = text.split(" ").length;
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
        selection.select();
        selection.load("text");
        await context.sync();
        let selectedText = selection.text;
        const platform = Office.context.platform !== Office.PlatformType.OfficeOnline;

        switch (concept) {
            case "object":
                platform ? selection.style = "Object" : selection.font.color = "#FF1493", selection.font.bold = true;
                break;
            case "event":
                platform ? selection.style = "Event" : selection.font.color = "#9932CC";
                break;
            case "process":
                platform ? selection.style = "Process" : selection.font.color = "#4B0082";
                break;
            case "role":
                platform ? selection.style = "Role" : selection.font.color = "#FFA07A";
                break;
            case "term":
                platform ? selection.style = "Term" : selection.font.color = "#FF6347";
                break;
            case "quantity":
                platform ? selection.style = "Quantity" : selection.font.color = "#ADFF2F", selection.font.bold = true;
                break;
            default:
                break;
        }
        selection.hyperlink = messageFromDialog.URL;
        selection.select(Word.SelectionMode.end)
        const range = context.document.body.getRange();
        await context.sync();
        const searchResults = range.search(previousText, { matchCase: false, matchWholeWord: false });
        searchResults.load("items");
        await context.sync();
        const occurrences = searchResults.items;

        occurrences.forEach(async (occurrence) => {
            occurrence.insertText(messageFromDialog.showAs, "Replace");
            switch (concept) {
                case "object":
                    platform ? occurrence.style = "Object" : occurrence.font.color = "#FF1493", occurrence.font.bold = true;
                    break;
                case "event":
                    platform ? occurrence.style = "Event" : occurrence.font.color = "#9932CC";
                    break;
                case "process":
                    platform ? occurrence.style = "Process" : occurrence.font.color = "#4B0082";
                    break;
                case "role":
                    platform ? occurrence.style = "Role" : occurrence.font.color = "#FFA07A";
                    break;
                case "term":
                    platform ? occurrence.style = "Term" : occurrence.font.color = "#FF6347";
                    break;
                case "quantity":
                    platform ? occurrence.style = "Quantity" : occurrence.font.color = "#ADFF2F", occurrence.font.bold = true;
                    break;
                default:
                    break;
            }
            occurrence.hyperlink = messageFromDialog.URL;
        });

        const NAMESPACE_URI = "prova";
        const uniqueId = Date.now();
        const xmlData = `<root xmlns="${NAMESPACE_URI}"><data id="${uniqueId}" text="${selectedText.toLowerCase()}">${JSON.stringify(messageFromDialog)}</data></root>`;

        deleteInformation(context, NAMESPACE_URI, selectedText);
        insertInformation(context, xmlData);
        await context.sync();
    }

    const handleChangeConcept = async (event: SelectChangeEvent) => {
        let removed = false;
        let char_remove;
        concept = event.target.value;
        await Word.run(async (context) => {
            let selection = context.document.getSelection();
            selection.load("paragraphs, text, styleBuiltIn, font");
            await context.sync();
            let paragraphCount = selection.paragraphs.items.length;
            let emptyParagraph = 0;
            for (let i = 0; i < selection.paragraphs.items.length; i++) { // se nella selezione includo anche i paragrafi, non funziona perfettamente
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
            selection.load("styleBuiltIn, text, style, hyperlink");
            await context.sync();

            // se la parola è selezionata a metà non si apre la finestra 
            let dialogUrl = 'https://localhost:3000/assets/';
            switch (event.target.value) {
                case "object":
                    dialogUrl += "object.html"
                    break;
                case "event":
                    dialogUrl += "event.html"
                    break;
                case "process":
                    dialogUrl += "process.html"
                    break;
                case "role":
                    dialogUrl += "role.html"
                    break;
                case "term":
                    dialogUrl += "term.html"
                    break;
                case "quantity":
                    dialogUrl += "quantity.html"
                    break;
                default:
                    dialogUrl += "";
                    break;
            }
            if (selection.hyperlink) {
                const information = await getInformation("prova", selection.text);
                const informationString = JSON.stringify(information);
                dialogUrl += `?information=${encodeURIComponent(informationString)}&`;
            }
            if (dialogUrl != "https://localhost:3000/assets/") {
                if (dialogUrl.includes("&")) {
                    dialogUrl += `selectedText=${encodeURIComponent(selection.text)}`;
                } else {
                    dialogUrl += `?selectedText=${encodeURIComponent(selection.text)}`;
                }
                Office.context.ui.displayDialogAsync(dialogUrl, {
                    height: 50,
                    width: 20,
                    displayInIframe: true,
                },
                    function (asyncResult) {
                        dialog = asyncResult.value;
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processEntities);
                    });
            }
            if(removed == true){
                selection.insertText(char_remove, "End");
                removed = false;
            }
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
                style={{ marginTop: "10px" }}
            >
                <Grid item xs={6}>
                    <p style={{ marginLeft: '3px', marginTop: '16px', fontSize: '17px' }}>Other Entities</p>
                </Grid>
                <Grid item xs={6}>
                    <FormControl
                        disabled={setDis}
                        sx={{ m: 1, minWidth: 120 }}
                        size="small"
                        style={{ position: 'relative', right: '18px' }}
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