import * as React from 'react';
import { useState, useEffect } from 'react';
import {TipografiaButton} from './Inlines/TipografiaButton'
import {FirstStyles} from './Inlines/FirstStyles'
import {ImportantEntities} from './Inlines/ImportantEntities'
import {OtherEntities} from './Inlines/OtherEntities'
import {Informative} from './Inlines/Informative'
import { ExpandWords } from './Inlines/ExpandWords';
import { AllInstances } from './Inlines/AllInstances';

export const Inlines = () => {
  const [expandedText, setExpandedText] = useState("");
  const [buttonStyle, setButtonStyle] = useState("");
  const [firstOccurence, setFirstOccurence] = useState("");
  const [selectedText, setSelectedText] = useState(""); 
  const [dis, setDis] = useState(true);
  const [bodyText, setBodyText] = useState('')
  const [fontStyle, setFontStyle] = useState("")
  const [first, setFirst] = useState("none");
  const [styleEntities, setFirstStyleEntities] = useState("none");
  const [entitiesStyle, setEntitiesStyle] = useState("");
  // Funzione di callback per aggiornare il valore di expandedText
  const handleExpandedTextChange = (text) => {
    setExpandedText(text);
  }
  
  const handleButtonStyle = (text) => {
    setButtonStyle(text);
  }

  const handleFirstOccurence = (text) => {
    setFirstOccurence(text);
  }

  const handleFontStyle = (text) => {
    setFontStyle(text);
  }

  const handleFirst = (text) => {
    setFirst(text);
  }

  const handleFirstStyleEntities = (text) => {
    setFirstStyleEntities(text);
  }

  const handleEntitiesStyle = (text) => {
    setEntitiesStyle(text);
  }

  useEffect(() => {
    const handleSelectionChange = async () => {
      try {
        await Word.run(async (context) => {
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

  return (
    <div>
      <FirstStyles onFontStyle={handleFontStyle} onFirst={handleFirst} expandedText={expandedText}/>

      <hr />
      <ImportantEntities onEntitiesStyle={handleEntitiesStyle} onFirstStyleEntities={handleFirstStyleEntities} expandedText={expandedText}/>
      <OtherEntities />
      <Informative />

      <hr />
      <div style={{display: 'flex', justifyContent: 'center', marginBottom: '5px', fontSize: '20px'}}>
        Tipography    
      </div>
      <TipografiaButton setDis={dis} onFirstOccurence={handleFirstOccurence} onButtonStyle={handleButtonStyle} expandedText={expandedText}/>

      <ExpandWords bodyText={bodyText} selectedText={selectedText} onExpandedTextChange={handleExpandedTextChange}/>
      
      <AllInstances fontStyle={fontStyle} buttonStyle={buttonStyle} firstOccurence={firstOccurence} first={first} expandedText={expandedText} firststyleEntities={styleEntities} entitiesStyle={entitiesStyle}/>
      
    </div>
  )
}