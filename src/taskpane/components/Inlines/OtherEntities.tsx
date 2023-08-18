import * as React from 'react';
import { useState } from 'react';
import Grid from '@mui/material/Grid';
import InputLabel from '@mui/material/InputLabel';
import MenuItem from '@mui/material/MenuItem';
import FormControl from '@mui/material/FormControl';
import Select, { SelectChangeEvent } from '@mui/material/Select';

export const OtherEntities = ({ info, setDis, expandedText, onOtherEntitiesStyle }) => {
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

    const updateStyle = async (context, messageFromDialog) => {
        let message;
        let selection = context.document.getSelection();
        await context.sync();
        selection.clear();
        selection.insertText(messageFromDialog.showAs);
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
        await context.sync();

        switch (concept) {
            case "object":
                selection.style = "Object"
                break;
            case "event":
                selection.style = "Event";
                break;
            case "process":
                selection.style = "Process";
                break;
            case "role":
                selection.style = "Role";
                break;
            case "term":
                selection.style = "Term";
                break;
            case "quantity":
                selection.style = "Quantity";
                break;
            default:
                selection.styleBuiltIn = "Normal";
                break;
        }
        selection.hyperlink = messageFromDialog.URL;
        message = "value of type Reference " + concept + " with URL: " + messageFromDialog.URL;
        info(message);
        await context.sync();
    }

        const handleChangeConcept = async (event: SelectChangeEvent) => {
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
                selection.load("styleBuiltIn, text");
                await context.sync();
            }

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
                    dialogUrl += ""
                    break;
            }

            if (dialogUrl != "https://localhost:3000/assets/") {    
                
                dialogUrl += `?selectedText=${encodeURIComponent(selection.text)}`;
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
            // passo al componente padre l'entità che l'utente ha scelto
            onOtherEntitiesStyle(event.target.value)
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