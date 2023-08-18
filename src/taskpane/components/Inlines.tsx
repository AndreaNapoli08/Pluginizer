import * as React from 'react';
import { useState, useEffect } from 'react';
import { TipografiaButton } from './Inlines/TipografiaButton'
import { FirstStyles } from './Inlines/FirstStyles'
import { ImportantEntities } from './Inlines/ImportantEntities'
import { OtherEntities } from './Inlines/OtherEntities'
import { Informative } from './Inlines/Informative'
import { ExpandWords } from './Inlines/ExpandWords';
import { AllInstances } from './Inlines/AllInstances';
import { ShowInfo } from './Inlines/ShowInfo';

export const Inlines = ({onHandleExpandedText, styleGSG}) => {

  // definizione dei vari stati utilizzati all'interno del componente
  const [expandedText, setExpandedText] = useState("");
  const [buttonStyle, setButtonStyle] = useState("");
  const [firstOccurence, setFirstOccurence] = useState("");
  const [selectedText, setSelectedText] = useState(""); 
  const [dis, setDis] = useState(true);
  const [bodyText, setBodyText] = useState("")
  const [fontStyle, setFontStyle] = useState("")
  const [first, setFirst] = useState("none");
  const [entitiesStyle, setEntitiesStyle] = useState("");
  const [styleOtherEntities, setStyleOtherEntities] = useState("");
  const [styleInformative, setStyleInformative] = useState("")
  const [info, setInfo] = useState("");

  // funzioni per aggiornare gli stati del componente
  const handleExpandedTextChange = (text) => {
    setExpandedText(text);
  }

  const handleInfo = (text) => {
    setInfo(text);
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

  const handleEntitiesStyle = (text) => {
    setEntitiesStyle(text);
  }

  const handleOtherEntitiesStyle = (text) => {
    setStyleOtherEntities(text);
  }

  const handleInformativeEntities = (text) => {
    setStyleInformative(text);
  }

  // il blocco all'interno di useEffect viene eseguito in automatico appena viene richiamato il componente
  useEffect(() => {
    const handleSelectionChange = async () => {  // funzione che viene richiamata ogni volta che il documento o la selezione del testo cambia
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
            setDis(true)  // stato che abilita o disabilita i bottoni della tipografia
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
      Office.EventType.DocumentSelectionChanged,  // gestore che viene richiamato ogni qual volta viene cambiato il documento o la selezione del testo
      handleSelectionChange
    );

    return () => {
      Office.context.document.removeHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        { handler: handleSelectionChange }
      );
    };
  }, []);

  onHandleExpandedText(expandedText)

  return (
    <div>
      {dis ? null : <ShowInfo info={info} selectedText={selectedText}/>}
      <FirstStyles info={handleInfo} setDis={dis} onFontStyle={handleFontStyle} onFirst={handleFirst} expandedText={expandedText}/>

      <hr />
      <ImportantEntities info={handleInfo} setDis={dis} onEntitiesStyle={handleEntitiesStyle} expandedText={expandedText}/>
      <OtherEntities info={handleInfo} setDis={dis} expandedText={expandedText} onOtherEntitiesStyle={handleOtherEntitiesStyle}/>
      <Informative setDis={dis} onInformativeStyle={handleInformativeEntities} expandedText={expandedText}/>

      <hr />
      <div style={{display: 'flex', justifyContent: 'center', marginBottom: '5px', fontSize: '20px'}}>
        Tipography      
      </div>
      <TipografiaButton setDis={dis} onFirstOccurence={handleFirstOccurence} onButtonStyle={handleButtonStyle} expandedText={expandedText}/>

      <ExpandWords bodyText={bodyText} selectedText={selectedText} onExpandedTextChange={handleExpandedTextChange}/>
      
      <AllInstances styleGSG={styleGSG} fontStyle={fontStyle} buttonStyle={buttonStyle} firstOccurence={firstOccurence} first={first} expandedText={expandedText} entitiesStyle={entitiesStyle} styleOtherEntities={styleOtherEntities} styleInformative={styleInformative}/>
      
    </div>
  )
}