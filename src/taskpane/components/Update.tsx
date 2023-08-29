import * as React from 'react';
import { useState } from 'react';
import Button from '@mui/material/Button';

export const Update = () => {
    let contextDocument;
    const getInformation = async (NAMESPACE_URI, word) => {
        Office.context.document.customXmlParts.getByNamespaceAsync(NAMESPACE_URI, async (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const xmlParts = result.value;
                for (const xmlPart of xmlParts) {
                    await xmlPart.getXmlAsync(async asyncResult => {    // questa istruzione non aspetta il completamento di ciascuna chiamata
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            const xmlData = asyncResult.value;
                            if (xmlData.includes(`text="${word.toLowerCase()}"`)) {
                                const parser = new DOMParser();
                                const xmlDoc = parser.parseFromString(xmlData, "text/xml");
                                const dataElement = xmlDoc.querySelector(`data[text="${word.toLowerCase()}"]`);
                                if (dataElement) {
                                    let jsonData = JSON.parse(dataElement.textContent);
                                    const range = contextDocument.document.body.getRange();
                                    await contextDocument.sync();
                                    const searchResults = range.search(word, { matchCase: false, matchWholeWord: false });
                                    searchResults.load("items");
                                    await contextDocument.sync();
                                    const occurrences = searchResults.items;
                                    occurrences.forEach(async (occurrence) => {
                                        switch (jsonData.entity) {
                                            case "reference":
                                                occurrence.styleBuiltIn = "IntenseReference";
                                                break;
                                            case "footnote":
                                                occurrence.styleBuiltIn = "IntenseEmphasis";
                                                break;
                                            case "date":
                                                occurrence.style = "Data1";
                                                break;
                                            case "organization":
                                                occurrence.style = "Organization";
                                                break;
                                            case "person":
                                                occurrence.style = "Person";
                                                break;
                                            case "location":
                                                occurrence.style = "Location";
                                                break;
                                            case "Other_Entities":
                                                switch (jsonData.type) {
                                                    case "object":
                                                        occurrence.style = "Object"
                                                        break;
                                                    case "event":
                                                        occurrence.style = "Event";
                                                        break;
                                                    case "process":
                                                        occurrence.style = "Process";
                                                        break;
                                                    case "quantity":
                                                        occurrence.style = "Quantity";
                                                        break;
                                                    case "role":
                                                        occurrence.style = "Role";
                                                        break;
                                                    case "term":
                                                        occurrence.style = "Term";
                                                        break;
                                                }
                                        }
                                        await contextDocument.sync();
                                    });
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
    }

    const updateDocument = async () => {
        await Word.run(async (context) => {
            contextDocument = context;
            const body = context.document.body;
            body.load("text");
            await context.sync();
            const BodyText = body.text;
            const wordsArray = BodyText.split(" ");
            wordsArray.forEach((word) => {
                getInformation("prova", word);
            });
        });
    }

    return (
        <div>
            <Button
                color="inherit"
                style={{
                    marginTop: "20px",
                    marginRight: "10px",
                    width: "220px",
                    height: "40px",
                    textDecoration: "underline",
                }}
                onClick={updateDocument}>
                Update Document
            </Button>
        </div>
    );
}