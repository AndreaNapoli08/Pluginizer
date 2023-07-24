import * as React from 'react';
import IconButton from '@mui/material/IconButton';

export const GSG = ({expandedText, onUpdateStyleGSG}) => {

    const isLetterOrNumber = (char) => {
        if (typeof char === "undefined") {
            return false;
        }else{
            return /^[a-zA-Z0-9]+$/.test(char);
        }
    }  

    const updateStyleGSG = async (styleGSG) => { 
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

            // solita funzione per espansione del testo
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
                selection.load("styleBuiltIn");
                selection.font.load("color, highlightColor")
                await context.sync();
            }
            
            // impostare lo stile selezionato
            switch(styleGSG){
                case 1:
                    selection.style = "GSG"
                    break;
                case 2:
                    selection.style = "GSG2";
                    break;
                case 3:
                    selection.style = "GSG3"
                    break;
                case 4:
                    selection.style = "GSG4"
                    break;
                case 5:
                    selection.style = "GSG5"
                    break;
                case 6:
                    selection.style = "GSG6"
                    break;
                case 7:
                    selection.style = "GSG7"
                    break;
                case 8:
                    selection.style = "GSG8"
                    break;
                case 9:
                    selection.style = "GSG9"
                    break;
                case 10:
                    selection.style = "GSG10"
                    break;
                case 11:
                    selection.style = "GSG11"
                    break;
                case 12:
                    selection.style = "GSG12"
                    break;
                case 13:
                    selection.style = "GSG13"
                    break;
                case 14:
                    selection.style = "GSG14"
                    break;
                case 15:
                    selection.style = "GSG15"
                    break;
                case 16:
                    selection.style = "GSG15"
                    break;
                case 17:
                    selection.style = "GSG16"
                    break;
                default:
                    break;
            }
            
            onUpdateStyleGSG(styleGSG) // passo al componente padre lo stile scelto
            onUpdateStyleGSG("") // setto "" così non rimane salvato l'ultimo stile usato
        });
    }

    return (
        <div>
            <div>
            <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(1)}>
                <img src="assets/GSG1.png" width={40} title="GSG1"/>
                <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>1 No poverty</span>
            </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(2)}>
                    <img src="assets/GSG2.png" width={40} title="GSG2"/>
                    <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>2 Zero Hunger</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(3)}>
                    <img src="assets/GSG3.png" width={40} title="GSG3"/>
                    <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>3 Good health and weel-being</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(4)}>
                    <img src="assets/GSG4.png" width={40} title="GSG4"/>
                    <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>4 Quality education</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(5)}>
                    <img src="assets/GSG5.png" width={40} title="GSG5"/>
                    <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>5 Gender equality</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(6)}>
                    <img src="assets/GSG6.png" width={40} title="GSG6"/>
                    <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>6 Clean water and sanitation</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(7)}>
                    <img src="assets/GSG7.png" width={40} title="GSG7"/>
                    <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>7 Affordable and clean energy</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(8)}>
                    <img src="assets/GSG8.png" width={40} title="GSG8"/>
                    <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>8 Decent Work And economic growth</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(9)}>
                    <img src="assets/GSG9.png" width={40} title="GSG9"/>
                    <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>9 Industry, innovation and infrastructure</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(10)}>
                    <img src="assets/GSG10.png" width={40} title="GSG10"/>
                    <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>10 Reduced inequalities</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(11)}>
                    <img src="assets/GSG11.png" width={40} title="GSG11"/>
                    <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>11 Sustainable cities and communities</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(12)}>
                    <img src="assets/GSG12.png" width={40} title="GSG12"/>
                    <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>12 Responsible consumption and production</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(13)}>
                    <img src="assets/GSG13.png" width={40} title="GSG13"/>
                    <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>13 Climate action</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(14)}>
                    <img src="assets/GSG14.png" width={40} title="GSG14"/>
                    <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>14 Life below water</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(15)}>
                    <img src="assets/GSG15.png" width={40} title="GSG15"/>
                    <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>15 Life on land</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(16)}>
                    <img src="assets/GSG16.png" width={40} title="GSG16"/>
                    <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>16 Peace, Justice and strong institutions</span>
                </IconButton>
            </div>
            <div>
                <IconButton color="inherit" style={{borderRadius: '10px', textAlign: "left"}} onClick={() => updateStyleGSG(17)}>
                    <img src="assets/GSG17.png" width={40} title="GSG17"/>
                    <span style={{fontSize: "16px", marginLeft: "15px", fontFamily: "cursive"}}>17 Partnerships for the goals</span>
                </IconButton>
            </div>
        </div>
    )
}