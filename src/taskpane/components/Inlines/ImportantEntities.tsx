import * as React from 'react';
import { useState, useEffect } from 'react';
import Grid from '@mui/material/Grid';
import IconButton from '@mui/material/IconButton';
import CalendarMonthIcon from '@mui/icons-material/CalendarMonth';
import FolderOpenIcon from '@mui/icons-material/FolderOpen';
import PersonIcon from '@mui/icons-material/Person';
import LocationOnIcon from '@mui/icons-material/LocationOn';
import AccessTimeIcon from '@mui/icons-material/AccessTime';

export const ImportantEntities = ({expandedText, onFirstStyleEntities, onEntitiesStyle}) => {
    /*const ooXMLtext = "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">" +
                                           "<pkg:part pkg:name =\"/_rels/.rels\" pkg:padding=\"512\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\">" +
                                            "<pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                            "<Relationship Target=\"word/document.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Id=\"rId1\" /></Relationships>" +
                                            "</pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/_rels/document.xml.rels\" pkg:padding=\"256\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\">" +
                                            "<pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                            "<Relationship Target=\"styles.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Id=\"rId1\" /></Relationships>" +
                                            "</pkg:xmlData></pkg:part>" +
                                            "<pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"><pkg:xmlData>" +
                                            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:pPr><w:pStyle w:val=\"AeP_ChapterHead\" /></w:pPr><w:t>Some text</w:t></w:r></w:p></w:body></w:document>" +
                                            "</pkg:xmlData></pkg:part>" +
                                            "<pkg:part pkg:name =\"/word/styles.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\">" +
                                            "<pkg:xmlData><w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                            "<w:style w:type=\"character\" w:styleId=\"DefaultParagraphFont\" w:default=\"1\"><w:name w:val=\"Default Paragraph Font\" /><w:uiPriority w:val=\"1\" /><w:semiHidden /><w:unhideWhenUsed /></w:style>" +
                                            "<w:style w:type=\"table\" w:styleId=\"TableNormal\" w:default=\"1\"><w:name w:val=\"Normal Table\" /><w:uiPriority w:val=\"99\" /><w:semiHidden /><w:unhideWhenUsed /><w:tblPr><w:tblInd w:type=\"dxa\" w:w=\"0\" /><w:tblCellMar><w:top w:type=\"dxa\" w:w=\"0\" /><w:left w:type=\"dxa\" w:w=\"108\" /><w:bottom w:type=\"dxa\" w:w=\"0\" /><w:right w:type=\"dxa\" w:w=\"108\" /></w:tblCellMar></w:tblPr></w:style>" +
                                            "<w:style w:type=\"numbering\" w:styleId=\"NoList\" w:default=\"1\"><w:name w:val=\"No List\" /><w:uiPriority w:val=\"99\" /><w:semiHidden /><w:unhideWhenUsed /></w:style>" +
                                            "<w:style w:type=\"paragraph\" w:styleId=\"AeP_ChapterHead\" w:customStyle=\"1\"><w:name w:val=\"Chapter Head\" /><w:uiPriority w:val=\"1\" /><w:qFormat /><w:pPr /><w:rPr><w:color w:val=\"365F91\" /></w:rPr></w:style></w:styles></pkg:xmlData>" +
                                            "</pkg:part></pkg:package>";*/
                                            
    /*const [styleMap, setStyleMap] = useState(new Map());
    useEffect(() => {
        const handleSelectionChange = async () => {
            try {
                await Word.run(async (context) => {
                    const body = context.document.body;
                    const paragraphs = body.paragraphs;
                    paragraphs.load("items");
                    await context.sync();
                    for (let i = 0; i < paragraphs.items.length; i++) {
                        const text = paragraphs.items[i].text;
                        const font = paragraphs.items[i].font;
                        font.load("name, size, color, bold, italic, underline");
                        await context.sync();
                        const fontInfo = {
                          name: font.name,
                          size: font.size,
                          color: font.color,
                          italic: font.italic,
                          bold: font.bold,
                          underline: font.underline
                        }
                        setStyleMap(styleMap.set(text, fontInfo));
                    }
                });
            } catch (error) {
            console.error(error);
            }
        };
        
        handleSelectionChange();
        }, []);*/

    
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
            if(expandedText != selection.text){
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
                selection.font.load("color")
                await context.sync();
            }

            onFirstStyleEntities(selection.font.color); 

            switch(entities) {
                case "Date" :
                    if(selection.font.color == "#FF0000"){
                        selection.styleBuiltIn = "Normal"
                        /*
                        let paragraph = selection.paragraphs.getFirst();
                        paragraph.load("text");
                        await context.sync();
                        const fontInfo = styleMap.get(paragraph.text)
                        selection.font.name = fontInfo.name;
                        selection.font.color = fontInfo.color;
                        selection.font.size = fontInfo.size;
                        selection.font.bold = fontInfo.bold;
                        selection.font.italic = fontInfo.italic;
                        selection.font.underline = fontInfo.underline;*/
                    }else{
                        selection.font.italic = true;
                        selection.font.bold = true;
                        selection.font.underline = "None";
                        selection.font.color = "red";
                        selection.font.name = "Abadi";
                        selection.font.size = 16;
                    }
                    break;
                case "Organization" :
                    if(selection.font.color == "#008000"){
                        selection.styleBuiltIn = "Normal";
                    }else{
                        selection.font.italic = false;
                        selection.font.bold = true;
                        selection.font.underline = "None"
                        selection.font.color = "green";
                        selection.font.name = "Times New Roman"
                        selection.font.size = 16;
                    }
                    break
                case "Person":
                    if(selection.font.color == "#0000FF"){
                        selection.styleBuiltIn = "Normal";
                    }else{
                        selection.font.italic = false;
                        selection.font.bold = true;
                        selection.font.underline = "DashLine";
                        selection.font.color = "blue";
                        selection.font.name = "Arial";
                        selection.font.size = 16;
                    }
                    break;
                case "Location":
                    if(selection.font.color == "#FFA500"){
                        selection.styleBuiltIn = "Normal";
                    }else{
                        selection.font.italic = false;
                        selection.font.bold = true;
                        selection.font.underline = "None";
                        selection.font.color = "orange";
                        selection.font.name = "Calibri"
                        selection.font.size = 16;
                    }
                    break;
                case "Time":
                    if(selection.font.color == "#800080"){
                        selection.styleBuiltIn = "Normal";
                    }else{
                        selection.font.italic = false;
                        selection.font.bold = true;
                        selection.font.underline = "None";
                        selection.font.color = "purple";
                        selection.font.name = "Century Gothic";
                        selection.font.size = 16;
                    }
                    break;
                default:
                    break;
            }

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