// licenza d'uso riservata ad Andrea Napoli e all'università si Bologna
import * as React from 'react';
import IconButton from '@mui/material/IconButton';

export const GSG = ({ expandedText }) => {

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

    const updateStyleGSG = async (styleGSG) => {
        let removed=false;
        let char_remove;
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

            // solita funzione per espansione del testo
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
                if(removed == true){
                    selection.insertText(char_remove, "End");
                    removed = false;
                }
            }

            selection.load("styleBuiltIn, text");
            selection.font.load("color, highlightColor")
            await context.sync();

            const platform = Office.context.platform !== Office.PlatformType.OfficeOnline;

            // impostare lo stile selezionato
            switch (styleGSG) {
                case 1:
                    platform ? selection.style = "GSG" : selection.font.color = "red", selection.font.bold = true;    
                    break;
                case 2:
                    platform ? selection.style = "GSG2" : selection.font.color = "orange", selection.font.bold = true;   
                    break;
                case 3:
                    platform ? selection.style = "GSG3" : selection.font.color = "green", selection.font.bold = true;   
                    break;
                case 4:
                    platform ? selection.style = "GSG4" : selection.font.color = "#B22222", selection.font.bold = true;   
                    break;
                case 5:
                    platform ? selection.style = "GSG5" : selection.font.color = "#FF0000", selection.font.bold = true;   
                    break;
                case 6:
                    platform ? selection.style = "GSG6" : selection.font.color = "#00FFFF", selection.font.bold = true;   
                    break;
                case 7:
                    platform ? selection.style = "GSG7" : selection.font.color = "yellow", selection.font.bold = true;   
                    break;
                case 8:
                    platform ? selection.style = "GSG8" : selection.font.color = "#A52A2A", selection.font.bold = true;   
                    break;
                case 9:
                    platform ? selection.style = "GSG9" : selection.font.color = "orange", selection.font.bold = true;   
                    break;
                case 10:
                    platform ? selection.style = "GSG10" : selection.font.color = "#FF00FF", selection.font.bold = true;   
                    break;
                case 11:
                    platform ? selection.style = "GSG11" : selection.font.color = "#FF7F50", selection.font.bold = true;   
                    break;
                case 12:
                    platform ? selection.style = "GSG12" : selection.font.color = "#F0E68C", selection.font.bold = true;   
                    break;
                case 13:
                    platform ? selection.style = "GSG13" : selection.font.color = "#32CD32", selection.font.bold = true;   
                    break;
                case 14:
                    platform ? selection.style = "GSG14" : selection.font.color = "#20B2AA", selection.font.bold = true;   
                    break;
                case 15:
                    platform ? selection.style = "GSG15" : selection.font.color = "#9ACD32", selection.font.bold = true;   
                    break;
                case 16:
                    platform ? selection.style = "GSG15" : selection.font.color = "#4682B4", selection.font.bold = true;   
                    break;
                case 17:
                    platform ? selection.style = "GSG16" : selection.font.color = "#00008B", selection.font.bold = true;   
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
                switch (styleGSG) {
                    case 1:
                        platform ? occurrence.style = "GSG" : occurrence.font.color = "red", occurrence.font.bold = true;    
                        break;
                    case 2:
                        platform ? occurrence.style = "GSG2" : occurrence.font.color = "orange", occurrence.font.bold = true;   
                        break;
                    case 3:
                        platform ? occurrence.style = "GSG3" : occurrence.font.color = "green", occurrence.font.bold = true;   
                        break;
                    case 4:
                        platform ? occurrence.style = "GSG4" : occurrence.font.color = "#B22222", occurrence.font.bold = true;   
                        break;
                    case 5:
                        platform ? occurrence.style = "GSG5" : occurrence.font.color = "#FF0000", occurrence.font.bold = true;   
                        break;
                    case 6:
                        platform ? occurrence.style = "GSG6" : occurrence.font.color = "#00FFFF", occurrence.font.bold = true;   
                        break;
                    case 7:
                        platform ? occurrence.style = "GSG7" : occurrence.font.color = "yellow", occurrence.font.bold = true;   
                        break;
                    case 8:
                        platform ? occurrence.style = "GSG8" : occurrence.font.color = "#A52A2A", occurrence.font.bold = true;   
                        break;
                    case 9:
                        platform ? occurrence.style = "GSG9" : occurrence.font.color = "orange", occurrence.font.bold = true;   
                        break;
                    case 10:
                        platform ? occurrence.style = "GSG10" : occurrence.font.color = "#FF00FF", occurrence.font.bold = true;   
                        break;
                    case 11:
                        platform ? occurrence.style = "GSG11" : occurrence.font.color = "#FF7F50", occurrence.font.bold = true;   
                        break;
                    case 12:
                        platform ? occurrence.style = "GSG12" : occurrence.font.color = "#F0E68C", occurrence.font.bold = true;   
                        break;
                    case 13:
                        platform ? occurrence.style = "GSG13" : occurrence.font.color = "#32CD32", occurrence.font.bold = true;   
                        break;
                    case 14:
                        platform ? occurrence.style = "GSG14" : occurrence.font.color = "#20B2AA", occurrence.font.bold = true;   
                        break;
                    case 15:
                        platform ? occurrence.style = "GSG15" : occurrence.font.color = "#9ACD32", occurrence.font.bold = true;   
                        break;
                    case 16:
                        platform ? occurrence.style = "GSG15" : occurrence.font.color = "#4682B4", occurrence.font.bold = true;   
                        break;
                    case 17:
                        platform ? occurrence.style = "GSG16" : occurrence.font.color = "#00008B", occurrence.font.bold = true;   
                        break;
                    default:
                        break;
                }
            });

            const NAMESPACE_URI = "prova";
            deleteInformation(context, NAMESPACE_URI, selection.text);
        });
    }

    return (
        <div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(1)}>
                    <img src="assets/GSG1.png" width={40} title="GSG1" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>1 No poverty</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(2)}>
                    <img src="assets/GSG2.png" width={40} title="GSG2" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>2 Zero Hunger</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(3)}>
                    <img src="assets/GSG3.png" width={40} title="GSG3" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>3 Good health and weel-being</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(4)}>
                    <img src="assets/GSG4.png" width={40} title="GSG4" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>4 Quality education</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(5)}>
                    <img src="assets/GSG5.png" width={40} title="GSG5" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>5 Gender equality</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(6)}>
                    <img src="assets/GSG6.png" width={40} title="GSG6" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>6 Clean water and sanitation</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(7)}>
                    <img src="assets/GSG7.png" width={40} title="GSG7" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>7 Affordable and clean energy</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(8)}>
                    <img src="assets/GSG8.png" width={40} title="GSG8" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>8 Decent Work And economic growth</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(9)}>
                    <img src="assets/GSG9.png" width={40} title="GSG9" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>9 Industry, innovation and infrastructure</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(10)}>
                    <img src="assets/GSG10.png" width={40} title="GSG10" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>10 Reduced inequalities</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(11)}>
                    <img src="assets/GSG11.png" width={40} title="GSG11" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>11 Sustainable cities and communities</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(12)}>
                    <img src="assets/GSG12.png" width={40} title="GSG12" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>12 Responsible consumption and production</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(13)}>
                    <img src="assets/GSG13.png" width={40} title="GSG13" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>13 Climate action</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(14)}>
                    <img src="assets/GSG14.png" width={40} title="GSG14" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>14 Life below water</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(15)}>
                    <img src="assets/GSG15.png" width={40} title="GSG15" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>15 Life on land</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(16)}>
                    <img src="assets/GSG16.png" width={40} title="GSG16" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>16 Peace, Justice and strong institutions</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{ borderRadius: '10px', textAlign: "left" }} onClick={() => updateStyleGSG(17)}>
                    <img src="assets/GSG17.png" width={40} title="GSG17" />
                    <span style={{ fontSize: "16px", marginLeft: "15px", fontFamily: "cursive" }}>17 Partnerships for the goals</span>
                </IconButton>
            </div>
        </div>
    )
}