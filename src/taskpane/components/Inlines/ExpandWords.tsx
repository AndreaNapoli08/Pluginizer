import * as React from 'react';
import { useState, useEffect } from 'react';
import FormControlLabel from '@mui/material/FormControlLabel';
import Checkbox from '@mui/material/Checkbox';

export const ExpandWords = ({bodyText, selectedText, onExpandedTextChange }) => {
    const [expandWords, setExpandWords] = useState(true);

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
            }
          }

          //controllo se prima della parte selezionata ci sono lettere
          if(bodyText[startIndex] == " " || startIndex == 0){
          }else{
            if(isLetterOrNumber(charBefore)){
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
            }   
          }
        }
    }

    onExpandedTextChange(expandedText);
    
    return (
        <div>
            <FormControlLabel 
                control={<Checkbox checked={expandWords} onChange={handleChangeCheckbox}/>} 
                label="Expand to whole words" 
                style={{display: 'flex', justifyContent: 'center', alignItems: 'center', marginTop: '10px'}}
            />
        </div>
    )
}