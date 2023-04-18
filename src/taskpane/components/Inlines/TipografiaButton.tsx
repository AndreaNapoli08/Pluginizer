import * as React from 'react';
import { useState, useEffect } from 'react';
import IconButton from '@mui/material/IconButton';

export const TipografiaButton = ({setDis, onFirstOccurence, onButtonStyle, expandedText}) => {
 
  const isLetterOrNumber = (char) => {
    if (typeof char === "undefined") {
      return false;
    }else{
      return /^[a-zA-Z0-9]+$/.test(char);
    }
    
  }

  const toggleFontStyle = async (style) => {
    await Word.run(async (context) => {
      let selection = context.document.getSelection();
      
      selection.load("text, font");
      await context.sync();
      if (selection.isNullObject) {
        return;
      }


      // https://stackoverflow.com/questions/58357313/getting-a-ranges-surrounding-text-in-office-js
      // https://stackoverflow.com/questions/51159644/word-js-apis-extending-a-range

      // Expand to end of sentence
      if(expandedText != selection.text){
        const startIndex = expandedText.indexOf(selection.text);
        const charBefore = expandedText[startIndex - 1];
        let text = selection.text;
        let spaceCount = text.split(" ").length;
       
        //selezione in avanti fino ad uno di quei caratteri
        let rngNextSent = selection.getNextTextRangeOrNullObject([".", " ", ","]);
        selection = selection.expandToOrNullObject(rngNextSent.getRange("Start"));
        
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

          /*let rngDocStart = selection.parentBody.getRange("Start");
          let rngToStart = selection.expandToOrNullObject(rngDocStart);
          let paras = rngToStart.paragraphs;
          paras.load("items");
          await context.sync();
          let nrParas = rngToStart.paragraphs.items.length-1;
          let rngPrev = rngToStart.paragraphs.items[nrParas];

          // Get the last sentence in that paragraph
          let rngParaSent = rngPrev.getTextRanges([".", " ", "/n"]);
          rngParaSent.load("items"); 
          await context.sync();
          let rngPrevSent = rngParaSent.items[0];

          // Extend the extended range to the last sentence in the previous paragraph
          selection = selection.expandToOrNullObject(rngPrevSent);*/
        }
        
        selection.select();
        selection.load("text, font");
        await context.sync();
      }
      

      
      // if (selection.text !== expandedText) {
      //   let selectionText = selection.text;
      //   const startIndex = expandedText.indexOf(selectionText);
      //   const endIndex = startIndex + selectionText.length - 1;
      
      //   if(startIndex >= 0){
      //     const charBefore = expandedText[startIndex - 1]; 
      //     const charAfter = expandedText[endIndex + 1];
      //     //controllo se dopo la parte selezionata ci sono lettere
      //     if(expandedText[endIndex] == " "){
      //     }else{
      //       if(isLetterOrNumber(charAfter)){
      //         let currentIndex = endIndex + 1;
      //         while (currentIndex < expandedText.length) {
      //           const currentChar = expandedText[currentIndex];
      //           if (isLetterOrNumber(currentChar)) {
      //             // currentChar.font.bold=true; non funziona
      //             currentIndex++;
      //           } else {
      //             break;
      //           }
      //         }
      //       }
      //     }
      //     //controllo se prima della parte selezionata ci sono lettere
      //     if(expandedText[startIndex] == " " || startIndex == 0){
      //     }else{
      //       if(isLetterOrNumber(charBefore)){
      //         let currentIndex = startIndex - 1;
      //         while (currentIndex >= 0) {
      //           const currentChar = expandedText[currentIndex];
      //           if (isLetterOrNumber(currentChar)) {
      //             // currentChar.font.bold=true; non funziona
      //             currentIndex--;
      //           } else {
      //             break;
      //           }
      //         }
      //       }   
      //     }
      //   }
      //   selection.insertText(expandedText, "Replace");
      // }

      switch (style) {
        case 'bold':
          onFirstOccurence(selection.font.bold);
          selection.font.bold = !selection.font.bold;
          onButtonStyle("bold");
          onButtonStyle(""); // lo setto a "" così quando ci sarà una nuova selezione non rimane salvato bold nella variabile
          break;
        case 'italic':
          onFirstOccurence(selection.font.italic);
          selection.font.italic = !selection.font.italic;
          onButtonStyle("italic");
          onButtonStyle("");
          break;
        case 'underline':
          if (selection.font.underline === "Mixed" || selection.font.underline === "None") {
            onFirstOccurence("None");
            selection.font.underline = "Single";
          } else {  
            onFirstOccurence("Single");
            selection.font.underline = "None";
          }
          onButtonStyle("underline");
          onButtonStyle("");
          break;
        default:
          break;
      }

      await context.sync();
    });
  } 

  return (
    <div style={{marginTop: '20px'}}>
      <div style={{display: "flex", justifyContent: "center", alignItems: "center"}}>
        <IconButton 
          disabled={setDis}
          color="inherit" 
          style = {{
            marginRight: "10px",
            border: "1px solid black",
            borderRadius: "10px",
            width: "75px",
            height: "40px"
          }}
          onClick={() => toggleFontStyle('bold')}>
            <b>G</b>
        </IconButton>
        <IconButton 
          disabled={setDis} 
          color="inherit" 
          style = {{
            marginRight: "10px",
            border: "1px solid black",
            borderRadius: "10px",
            width: "75px",
            height: "40px"
          }}
          onClick={() => toggleFontStyle('italic')}>
            <i>I</i>
        </IconButton>
        <IconButton 
          disabled={setDis} 
          color="inherit" 
          style = {{
            marginRight: "10px",
            border: "1px solid black",
            borderRadius: "10px",
            width: "75px",
            height: "40px"
          }}
          onClick={() => toggleFontStyle('underline')}>
            <u>S</u>
        </IconButton>
      </div>
    </div>
  );
}
