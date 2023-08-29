// licenza d'uso riservata ad Andrea Napoli e all'università si Bologna

// Ottieni i parametri dalla query string dell'URL
const queryParams = new URLSearchParams(window.location.search);

// Leggi il valore di 'selectedText' dalla query string
const selectedText = queryParams.get('selectedText');
const information = queryParams.get('information');
const parsedInformation = JSON.parse(information);

// Utilizza il valore di 'selectedText' nella tua finestra di dialogo come desideri
const showAsInput = document.getElementById('showAsInput');
const urlInput = document.getElementById('urlInput');

if(parsedInformation != null){
  urlInput.value = parsedInformation.URL;
}

if (showAsInput) {
    showAsInput.value = selectedText;
}

function submitForm(typeEntity) {
    const formData = {
        entity: "Other_Entities",
        type: typeEntity,
        showAs: showAsInput.value,
        URL: urlInput.value
    };

    console.log(typeEntity);
    Office.onReady(function (info) {
      if (info.host === Office.HostType.Word || info.host === Office.HostType.Excel || info.host === Office.HostType.PowerPoint) {
        Office.context.ui.messageParent(JSON.stringify(formData));
      } else {
        console.log("Errore: ambiente Office non riconosciuto");
      }
    });
    
}