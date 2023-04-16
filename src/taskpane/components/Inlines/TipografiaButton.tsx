import * as React from 'react';
import { useState, useEffect } from 'react';
import IconButton from '@mui/material/IconButton';

export const TipografiaButton = ({setDis, onFirstOccurence, onButtonStyle, expandedText}) => {

  const isLetterOrNumber = (char) => {
    return /^[a-zA-Z0-9]+$/.test(char);
  }

  const toggleFontStyle = async (style) => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text, font");
      await context.sync();
      if (selection.isNullObject) {
        return;
      }

      if (selection.text !== expandedText) {
        let selectionText = selection.text;
        const startIndex = expandedText.indexOf(selectionText);
        const endIndex = startIndex + selectionText.length - 1;
      
        if(startIndex >= 0){
          const charBefore = expandedText[startIndex - 1]; 
          const charAfter = expandedText[endIndex + 1];
          //controllo se dopo la parte selezionata ci sono lettere
          if(expandedText[endIndex] == " "){
          }else{
            if(isLetterOrNumber(charAfter)){
              let currentIndex = endIndex + 1;
              while (currentIndex < expandedText.length) {
                const currentChar = expandedText[currentIndex];
                if (isLetterOrNumber(currentChar)) {
                  // currentChar.font.bold=true; non funziona
                  currentIndex++;
                } else {
                  break;
                }
              }
            }
          }
          //controllo se prima della parte selezionata ci sono lettere
          if(expandedText[startIndex] == " " || startIndex == 0){
          }else{
            if(isLetterOrNumber(charBefore)){
              let currentIndex = startIndex - 1;
              while (currentIndex >= 0) {
                const currentChar = expandedText[currentIndex];
                if (isLetterOrNumber(currentChar)) {
                  // currentChar.font.bold=true; non funziona
                  currentIndex--;
                } else {
                  break;
                }
              }
            }   
          }
        }
        selection.insertText(expandedText, "Replace");
      }

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
