import * as React from 'react';
import { useState, useEffect } from 'react';
import FormControlLabel from '@mui/material/FormControlLabel';
import Checkbox from '@mui/material/Checkbox';

export const ExpandWords = ({ onExpandedTextChange }) => {
    const [expandWords, setExpandWords] = useState(true);
    const [selectedText, setSelectedText] = useState("Nessun testo selezionato")
    const [bodyText, setBodyText] = useState('')

    const handleChangeCheckbox = (event: React.ChangeEvent<HTMLInputElement>) => {
        setExpandWords(event.target.checked);
    };

    const isLetterOrNumber = (char) => {
      return /^[a-zA-Z0-9]+$/.test(char);
    }

    let expandedText = selectedText;
    if(expandWords){
        const startIndex = bodyText.indexOf(selectedText)
        const endIndex = startIndex + selectedText.length - 1;
        if(startIndex >= 0){
          const charBefore = bodyText[startIndex - 1]; 
          const charAfter = bodyText[endIndex + 1];
          
          //controllo se dopo la parte selezionata ci sono lettere
          if(bodyText[endIndex] == " "){
            //console.log("la parola è completa alla fine")
          }else{
            if(isLetterOrNumber(charAfter)){
              //console.log("la parola non è completa, dopo ci sono lettere");
              let currentIndex = endIndex + 1;
              while (currentIndex < bodyText.length) {
                const currentChar = bodyText[currentIndex];
                if (isLetterOrNumber(currentChar)) {
                  expandedText += currentChar;
                  currentIndex++;
                } else {
                  break;
                }
              }
              //console.log("La parola completa è: ", expandedText)
            }
          }

          //controllo se prima della parte selezionata ci sono lettere
          if(bodyText[startIndex] == " " || startIndex == 0){
            //console.log("la parola è completa all'inizio")
          }else{
            if(isLetterOrNumber(charBefore)){
              //console.log("la parola non è completa, prima ci sono lettere");
              let currentIndex = startIndex - 1;
              while (currentIndex >= 0) {
                const currentChar = bodyText[currentIndex];
                if (isLetterOrNumber(currentChar)) {
                  expandedText = currentChar + expandedText;
                  currentIndex--;
                } else {
                  break;
                }
              }
              //console.log("La parola completa è: ", expandedText)
            }   
          }
        }
    }

    onExpandedTextChange(expandedText);
    useEffect(() => {
        Word.run(async (context) => {
            // Ottenere il testo selezionato dal documento
            const selection = context.document.getSelection();
            selection.load("text");
            const body = context.document.body;
            body.load("text");
            await context.sync();

            Office.context.document.addHandlerAsync(
              Office.EventType.DocumentSelectionChanged,
              () => {
                Word.run(async (context) => {
                  const newSelection = context.document.getSelection();
                  newSelection.load("text");
                  const newBody = context.document.body;
                  newBody.load("text");
                  await context.sync();
                  const newSelectedText = newSelection.text;
                  const newBodyText = newBody.text;
                  setBodyText(newBodyText); 
                  if(newSelectedText.length === 0){
                    setSelectedText("Nessun testo selezionato")
                  }else{
                    setSelectedText(newSelectedText)
                  } 
                });
              });
            return context.sync();
        });
    });

    return (
        <div>
            <p>Testo: {selectedText}</p><br />
            <p>Testo nel documento: {bodyText}</p>
            <FormControlLabel 
                control={<Checkbox checked={expandWords} onChange={handleChangeCheckbox}/>} 
                label="Expand to whole words" 
                style={{display: 'flex', justifyContent: 'center', alignItems: 'center', marginTop: '10px'}}
            />
        </div>
    )
}