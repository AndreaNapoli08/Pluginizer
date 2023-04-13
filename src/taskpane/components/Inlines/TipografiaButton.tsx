import * as React from 'react';
import { useState, useEffect } from 'react';
import IconButton from '@mui/material/IconButton';

export const TipografiaButton = ({expandedText}) => {
  const [selectedText, setSelectedText] = useState(""); 
  const [dis, setDis] = useState(true);

  useEffect(() => {
    const handleSelectionChange = async () => {
      try {
        await Word.run(async (context) => {
          const newSelection = context.document.getSelection();
          newSelection.load("text");
          await context.sync();
          const newSelectedText = newSelection.text;
          if(newSelectedText.length === 0){
            setSelectedText("Nessun testo selezionato")
            setDis(true)
          }else{
            setSelectedText(newSelectedText)
            setDis(false)
          } 
        });
      } catch (error) {
        console.error(error);
      }
    };

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      handleSelectionChange
    );

    return () => {
      Office.context.document.removeHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        { handler: handleSelectionChange }
      );
    };
  }, []);

  const isLetterOrNumber = (char) => {
    return /^[a-zA-Z0-9]+$/.test(char);
  }

  const toggleFontStyle = async (style) => {
    await Word.run(async (context) => {
      const textNode = document.body.firstChild;
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
                  // carattere da eliminare
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
                  const range = document.createRange();
                  range.setStart(textNode, currentIndex)
                  range.setEnd(textNode, startIndex);
                  //selection.expandTo(range);
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
          selection.font.bold = !selection.font.bold;
          break;
        case 'italic':
          selection.font.italic = !selection.font.italic;
          break;
        case 'underline':
          if (selection.font.underline === "Mixed" || selection.font.underline === "None") {
            selection.font.underline = "Single";
          } else {  
            selection.font.underline = "None";
          }
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
          disabled={dis}
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
          disabled={dis} 
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
          disabled={dis} 
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
