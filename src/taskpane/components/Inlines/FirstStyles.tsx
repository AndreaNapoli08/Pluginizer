// licenza d'uso riservata ad Andrea Napoli e all'università si Bologna
import * as React from 'react';
import { useState } from 'react';
import IconButton from '@mui/material/IconButton';
import Grid from '@mui/material/Grid';
import LinkIcon from '@mui/icons-material/Link';
import LiveHelpIcon from '@mui/icons-material/LiveHelp';
import NoteAltIcon from '@mui/icons-material/NoteAlt';
import { json } from 'express';

export const FirstStyles = ({ setDis, expandedText }) => {
    let dialog;

    const isLetterOrNumber = (char) => {
        if (typeof char === "undefined") {
            return false;
        } else {
            return /^[a-zA-Z0-9]+$/.test(char);
        }
    }

    const processReference = async (arg) => {
        const messageFromDialog = JSON.parse(arg.message);

        await Word.run(async (context) => {
            await updateStyleBuiltIn(context, messageFromDialog);
        });
    }

    const processFootnote = async (arg) => {
        const messageFromDialog = JSON.parse(arg.message);
        if (messageFromDialog.definition != "") {
            dialog.close();

            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                selection.load("styleBuiltIn, text");
                await context.sync();
                let selectedText = selection.text;
                selection.styleBuiltIn = "IntenseEmphasis";
                selection.insertText(" ", "End");
                selection.insertFootnote(selectedText + ": " + messageFromDialog.definition);
                selection.select(Word.SelectionMode.end);
                const range = context.document.body.getRange();
                await context.sync();
                const searchResults = range.search(selectedText, { matchCase: false, matchWholeWord: false });
                searchResults.load("items");
                await context.sync();
                const occurrences = searchResults.items;
                occurrences.forEach(async (occurrence) => {
                    occurrence.styleBuiltIn = "IntenseEmphasis";
                });

                const NAMESPACE_URI = "prova";
                const uniqueId = Date.now();
                const xmlData = `<root xmlns="${NAMESPACE_URI}"><data id="${uniqueId}" text="${selectedText.toLowerCase()}">${JSON.stringify(messageFromDialog)}</data></root>`;

                // eliminiamo l'informazione attuale
                deleteInformation(context, NAMESPACE_URI, selectedText);
                // Inserisci la nuova informazione aggiunta 
                insertInformation(context, xmlData);
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
                                        switch (jsonData.entity) {
                                            case "reference":
                                                switch (jsonData.type) {
                                                    case "ref":
                                                        resolve({
                                                            type: "ref",
                                                            number: jsonData.numeroArticolo,
                                                            documento: jsonData.documento
                                                        });
                                                        break;
                                                    case "mref":
                                                        resolve({
                                                            type: "mref",
                                                            number: jsonData.numeriArticoli,
                                                            documento: jsonData.documento
                                                        });
                                                        break;
                                                    case "rref":
                                                        resolve({
                                                            type: "rref",
                                                            dal: jsonData.dal,
                                                            al: jsonData.al,
                                                            documento: jsonData.documento
                                                        });
                                                        break;
                                                    default:
                                                        resolve({});
                                                        break;
                                                }
                                                break;
                                            case "footnote":
                                                resolve({
                                                    definition: jsonData.definition,
                                                })
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
        console.log("inserimentoooo")
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
        selection.load("styleBuiltIn, text");
        await context.sync();
        const selectedText = selection.text;
        const NAMESPACE_URI = "prova";
        const uniqueId = Date.now();
        const xmlData = `<root xmlns="${NAMESPACE_URI}"><data id="${uniqueId}" text="${selectedText.toLowerCase()}">${JSON.stringify(messageFromDialog)}</data></root>`;

        switch (messageFromDialog.type) {
            case "ref":
                if (messageFromDialog.numeroArticolo != "" && messageFromDialog.documento != "") {
                    dialog.close();
                    selection.styleBuiltIn = "IntenseReference";
                }
                break;
            case "mref":
                if (messageFromDialog.numeriArticoli != "" && messageFromDialog.documento != "") {
                    dialog.close();
                    selection.styleBuiltIn = "IntenseReference";
                }
                break
            case "rref":
                if (messageFromDialog.dal != "" && messageFromDialog.al != "" && messageFromDialog.documento != "") {
                    dialog.close();
                    selection.styleBuiltIn = "IntenseReference";
                }
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

        occurrences.forEach(async (occurrence) => {
            console.log(occurrence)
            switch (messageFromDialog.type) {
                case "ref":
                    occurrence.styleBuiltIn = "IntenseReference";
                    break;
                case "mref":
                    occurrence.styleBuiltIn = "IntenseReference";
                    break
                case "rref":
                    occurrence.styleBuiltIn = "IntenseReference";
                    break;
                default:
                    break;
            }
        });

        // eliminiamo l'informazione attuale
        deleteInformation(context, NAMESPACE_URI, selectedText);
        // Inserisci la nuova informazione aggiunta 
        insertInformation(context, xmlData);

    }

    // funzione che aggiorna lo stile del testo
    const updateStyle = async (style) => {
        await Word.run(async (context) => {
            let selection = context.document.getSelection();
            selection.load("paragraphs, text, styleBuiltIn");
            await context.sync();
            let paragraphCount = selection.paragraphs.items.length;
            let emptyParagraph = 0;

            for (let i = 0; i < selection.paragraphs.items.length; i++) { // se nella selezione includo anche i paragrafi vuoti, la selezione espanda non funziona correttamente 
                if (selection.paragraphs.items[i].text == "") {
                    emptyParagraph++;
                }
            }

            // stessa funzione di espansione del testo
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
                    if (paragraphCount > 1) { // se più paragraphi sono compresi, andare a capo lo prende come una parola e quindi spaceCount va incrementato con il numero di paragrafi -1. Inoltre bisogna togliere anche i paragrafi vuoti
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
            }
            selection.load("styleBuiltIn, text, style");
            await context.sync();

            const NAMESPACE_URI = "prova";
            let dialogUrl = 'https://localhost:3000/assets/';
            // solo nel caso in cui si tratti di un reference si deve aprire la finestra di dialogo
            switch (style) {
                case "IntenseReference":
                    dialogUrl += 'reference.html';
                    if (selection.styleBuiltIn == "IntenseReference") {
                        const information = await getInformation("prova", selection.text);
                        const informationString = JSON.stringify(information);
                        dialogUrl += `?information=${encodeURIComponent(informationString)}`;
                    }
                    Office.context.ui.displayDialogAsync(dialogUrl, {
                        height: 70,
                        width: 45,
                        displayInIframe: true,
                    },
                        function (asyncResult) {
                            dialog = asyncResult.value;
                            dialog.addEventHandler(Office.EventType.DialogMessageReceived, processReference);
                        });
                    break;
                case "Heading8":
                    selection.styleBuiltIn = "Heading8";
                    break;
                case "IntenseEmphasis":
                    dialogUrl += 'footnote.html';
                    if (selection.styleBuiltIn == "IntenseEmphasis") {
                        const information = await getInformation("prova", selection.text);
                        const informationString = JSON.stringify(information);
                        dialogUrl += `?information=${encodeURIComponent(informationString)}`;
                    }
                    Office.context.ui.displayDialogAsync(dialogUrl, {
                        height: 70,
                        width: 45,
                        displayInIframe: true,
                    },
                        function (asyncResult) {
                            dialog = asyncResult.value;
                            dialog.addEventHandler(Office.EventType.DialogMessageReceived, processFootnote);
                        });
                    break;
                case "Normal":
                    selection.styleBuiltIn = "Normal";
                    deleteInformation(context, NAMESPACE_URI, selection.text);
                    selection.select(Word.SelectionMode.end);
                    selection.load("hyperlink");
                    await context.sync();
                    selection.hyperlink = null;
                    /*selection.load("footnotes");
                    await context.sync();
                    selection.footnotes.load("items");
                    await context.sync();
                    selection.footnotes.items[0].delete();*/
                    break;
                default:
                    break;
            }

            const range = context.document.body.getRange();
            await context.sync();
            const searchResults = range.search(selection.text, { matchCase: false, matchWholeWord: false });
            searchResults.load("items");
            await context.sync();
            const occurrences = searchResults.items;

            occurrences.forEach(async (occurrence) => {
                switch (style) {
                    case "Heading8":
                        occurrence.styleBuiltIn = "Heading8";
                        break;
                    case "Normal":
                        occurrence.styleBuiltIn = "Normal";
                        occurrence.hyperlink = null;
                        /*occurrence.footnotes.load("items");
                        await context.sync();
                        occurrence.footnotes.items[0].delete();*/
                        // qui non cancelliamo le informazioni perché gia prima le cancella per tutte le occorrenze
                        break;
                    default:
                        break;
                }
            });
        });
    }

    return (
        <div>
            <div style={{ marginBottom: "15px" }}>
                <IconButton disabled={setDis} color="inherit" style={{ borderRadius: '10px' }} onClick={() => updateStyle('IntenseReference')}>
                    <span style={{ fontSize: "18px" }}>Reference</span>
                    <LinkIcon style={{ marginLeft: "10px" }} />
                </IconButton>
            </div>
            <Grid
                container
                direction="row"
                justifyContent="left"
                alignItems="flex-start"
                spacing={2}
            >
                <Grid item xs={6}>
                    <div>
                        <IconButton disabled={setDis} color="inherit" style={{ borderRadius: '10px' }} onClick={() => updateStyle('Heading8')}>
                            <span style={{ fontSize: "18px" }}>Definition</span>
                            <LiveHelpIcon style={{ marginLeft: "10px" }} />
                        </IconButton>
                    </div>
                </Grid>
                <Grid item xs={6}>
                    <div style={{ marginBottom: "15px", marginLeft: "40px" }}>
                        <IconButton disabled={setDis} color="inherit" style={{ borderRadius: '10px' }} onClick={() => updateStyle('Normal')}>
                            <span style={{ fontSize: "18px" }}>Normal</span>
                        </IconButton>
                    </div>
                </Grid>
            </Grid>

            <div>
                <IconButton disabled={setDis} color="inherit" style={{ borderRadius: '10px' }} onClick={() => updateStyle('IntenseEmphasis')}>
                    <span style={{ fontSize: "18px" }}>Footnote</span>
                    <NoteAltIcon style={{ marginLeft: "10px" }} />
                </IconButton>
            </div>
        </div>
    )
}