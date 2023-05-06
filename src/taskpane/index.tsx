import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { templateDocument } from './template'

initializeIcons();

let isOfficeInitialized = false;
const title = "Contoso Task Pane Add-in";

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <ThemeProvider>
        <Component title={title} isOfficeInitialized={isOfficeInitialized} />
      </ThemeProvider>
    </AppContainer>,
    document.getElementById("container")
  );
};

// Reinderizza l'applicazione dopo aver inizializzato l'ambiente Office
Office.onReady(() => {
  isOfficeInitialized = true;
    Word.run(async (context) => {
      var customDocProps = context.document.properties.customProperties;
      context.load(customDocProps); // carichiamo le proprietà del documento attuale
      await context.sync(); 

      if(customDocProps.items.length == 0 || !customDocProps.items[0].value || customDocProps.items[0].key != "AKN Template"){ // vuol dire che non ha proprietà o che non possiede la proprietà AKN Template
        const myNewDoc = context.application.createDocument(templateDocument); // creazione del nuovo documento contente il template con gli stili
        context.load(myNewDoc);
        await context.sync();

        // prelevo il body del documento attuale
        const body = context.document.body;

        // prelevo l'XML del documento attuale
        const bodyXML = body.getOoxml();

        // Sincronizzio il documento e ritonro una problema che indica il completamento della task
        return context.sync().then(async function () {
            myNewDoc.body.insertOoxml(bodyXML.value, 'End'); // inserisco l'XML del documento attuale all'interno del nuovo documento contente il template
            console.log(bodyXML.value)
            await context.sync()
            await myNewDoc.save();  // aspetto il salvataggio del documento
            myNewDoc.open(); // apro il documento
            await context.sync()
        });
      }
    });

  render(App);
});

// permette di caricare in modo dinamico i moduli del'applicazione, senza dover aggiornare manualmente la pagina
if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
