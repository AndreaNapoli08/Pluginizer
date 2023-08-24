import * as React from 'react';
import { useState } from 'react';

export const ShowInfo = ({ expandedText }) => {
  const [sel, setSel] = useState("");
  const [info, setInfo] = useState("");
  let NAMESPACE_URI = "prova";
  const isLetterOrNumber = (char) => {
    if (typeof char === "undefined") {
      return false;
    } else {
      return /^[a-zA-Z0-9]+$/.test(char);
    }
  }

  const runWordAutomation = async () => {
    try {
      await Word.run(async (context) => {
        let selection = context.document.getSelection();
        let exit = false;
        selection.load("paragraphs, text, styleBuiltIn, font");
        await context.sync();
        let paragraphCount = selection.paragraphs.items.length;
        let emptyParagraph = 0;
        for (let i = 0; i < selection.paragraphs.items.length; i++) { // se nella selezione includo anche i paragrafi vuoti, non funziona perfettamente
          if (selection.paragraphs.items[i].text == "") {
            emptyParagraph++;
          }
        }
        // stessa funzione di espansione
        //if (expandedText != selection.text && selection.text != "") {  // se tolgo l'if in automatico espande a prescindere la parola
        const startIndex = expandedText.indexOf(selection.text);
        const charBefore = expandedText[startIndex - 1];

        let text = selection.text;
        let spaceCount = text.split(" ").length;

        //selezione in avanti fino ad uno di quei caratteri
        if (text[text.length - 1] != " ") {
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
          await context.sync();
        }
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
        //}
        let parola_trovata = false;
        selection.load("text");
        await context.sync();
        text = selection.text;
        // se l'ultimo carattere è uno spazio bianco lo tolgo perché causa problemi
        if (text[text.length - 1] == " ") {
          text = text.slice(0, -1);
          await context.sync();
        }
        setSel(text);
        await context.sync();

        Office.context.document.customXmlParts.getByNamespaceAsync(NAMESPACE_URI, async (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const xmlParts = result.value;
            for (const xmlPart of xmlParts) {   // questo for viene eseguito più volte, non so perchè
              if (exit) {
                break;
              }
              await xmlPart.getXmlAsync(asyncResult => {    // questa istruzione non aspetta il completamento di ciascuna chiamata
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  const xmlData = asyncResult.value;
                  if (xmlData.includes(`text="${sel.toLowerCase()}"`)) { // ricerca dell'informazione associata alla parola
                    parola_trovata = true;
                    const parser = new DOMParser();
                    const xmlDoc = parser.parseFromString(xmlData, "text/xml");
                    const dataElement = xmlDoc.querySelector(`data[text="${sel.toLowerCase()}"]`);
                    if (dataElement) {
                      let jsonData = JSON.parse(dataElement.textContent);
                      let message;
                      switch (jsonData.entity) {
                        case "date":
                          message = "value of type Date with this characteristics: " + jsonData.day + ' ' + jsonData.month + ' ' + jsonData.year + ', ' + jsonData.time;;
                          break;
                        case "organization":
                          message = "value of type Organization with this characteristics: " + jsonData.organization;
                          break;
                        case "person":
                          message = "value of type Person with this characteristics: " + jsonData.person;;
                          break;
                        case "location":
                          message = "value of type Location with this characteristics: " + jsonData.location;
                          break;
                        case "reference":
                          switch (jsonData.type) {
                            case "ref":
                              if (jsonData.numeroArticolo != "" && jsonData.documento != "") {
                                message = "value of type Ref with article " + jsonData.numeroArticolo + " in a document " + jsonData.documento;
                              }
                              break;
                            case "mref":
                              if (jsonData.numeriArticoli != "" && jsonData.documento != "") {
                                message = "value of type MRef with this articles: " + jsonData.numeriArticoli + " in a document " + jsonData.documento;
                              }
                              break
                            case "rref":
                              if (jsonData.dal != "" && jsonData.al != "" && jsonData.documento != "") {
                                message = "value of type RRef with articles from " + jsonData.dal + " to " + jsonData.al + " in a document " + jsonData.documento;
                              }
                              break;
                            default:
                              break;
                          }
                          break;
                        case "Other_Entities":
                          message = "value of type " + jsonData.type + " with this URL: " + jsonData.URL + " show as " + jsonData.showAs;
                          break;
                        case "footnote":
                          message = "value of type footnote with definition: " + jsonData.definition;
                      }
                      setInfo(message);
                      exit = true;
                    }
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
      });
    } catch (error) {
      console.error("Word Automation Error:", error);
    }
  };

  React.useEffect(() => {
    runWordAutomation();
  });

  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '5px', fontSize: '20px' }}>
        Information:
      </div>
      <div style={{ fontSize: '15px' }}>
        <b>{sel}: </b> {info}
      </div>
      <hr />
    </div>
  )
}