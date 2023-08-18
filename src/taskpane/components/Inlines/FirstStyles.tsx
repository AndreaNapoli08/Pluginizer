import * as React from 'react';
import { useState } from 'react';
import IconButton from '@mui/material/IconButton';
import Grid from '@mui/material/Grid';
import LinkIcon from '@mui/icons-material/Link';
import LiveHelpIcon from '@mui/icons-material/LiveHelp';
import NoteAltIcon from '@mui/icons-material/NoteAlt';

export const FirstStyles = ({ info, setDis, onFontStyle, onFirst, expandedText }) => {
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
        if(messageFromDialog.definition != ""){
            dialog.close();

            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                selection.load("styleBuiltIn");
                await context.sync();
                selection.styleBuiltIn = "IntenseEmphasis";
                selection.insertFootnote(messageFromDialog.definition);
            });
        }
        
    }

    const updateStyleBuiltIn = async (context, messageFromDialog) => {
        let message;
        const selection = context.document.getSelection();
        selection.load("styleBuiltIn");
        await context.sync();
        switch (messageFromDialog.type) {
            case "ref":
                if(messageFromDialog.numeroArticolo != "" && messageFromDialog.documento != ""){
                    message = "value of type Ref with article " + messageFromDialog.numeroArticolo + " in a document " + messageFromDialog.documento;
                    dialog.close();
                }
                break;
            case "mref":
                if(messageFromDialog.numeriArticoli != "" && messageFromDialog.documento != ""){
                    message = "value of type MRef with this articles: " + messageFromDialog.numeriArticoli + " in a document " + messageFromDialog.documento;
                    dialog.close();
                }
                break
            case "rref":
                if(messageFromDialog.dal != "" && messageFromDialog.al != "" && messageFromDialog.documento != ""){
                    message = "value of type RRef with articles from " + messageFromDialog.dal + " to " + messageFromDialog.al + " in a document " + messageFromDialog.documento;
                    dialog.close();
                }
                break;
            default:
                break;
        }
        selection.styleBuiltIn = "IntenseReference";
        info(message);
        await context.sync();
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
                selection.load("styleBuiltIn");
                await context.sync();
            }

            // passiamo al componente padre lo stile del testo prima di essere modificato
            onFirst(selection.styleBuiltIn);

            let dialogUrl = 'https://localhost:3000/assets/';
            // solo nel caso in cui si tratti di un reference si deve aprire la finestra di dialogo
            switch (style) {
                case "IntenseReference":
                    dialogUrl += 'reference.html';
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
                case "Heading6":
                    selection.styleBuiltIn = "Heading6"
                    break
                case "IntenseEmphasis":
                    dialogUrl += 'footnote.html';
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
                    selection.load("hyperlink, footnotes");
                    await context.sync();
                    selection.hyperlink = null;
                    selection.footnotes.load("items");
                    await context.sync();
                    selection.footnotes.items[0].delete();
                default:
                    break;
            }

            onFontStyle(style); // passiamo al componente padre lo stile che l'utente ha scelto
            onFontStyle(""); // lo setto a "" così quando ci sarà una nuova selezione non rimane salvato l'ultimo stile
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
                        <IconButton disabled={setDis} color="inherit" style={{ borderRadius: '10px' }} onClick={() => updateStyle('Heading6')}>
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