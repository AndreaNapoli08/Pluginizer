// Ottieni i parametri dalla query string dell'URL
const queryParams = new URLSearchParams(window.location.search);

// Leggi il valore di 'selectedText' dalla query string
const selectedText = queryParams.get('selectedText');

// Utilizza il valore di 'selectedText' nella tua finestra di dialogo come desideri
const showAsInput = document.getElementById('showAsInput');
const urlInput = document.getElementById('urlInput');
if (showAsInput) {
    showAsInput.value = selectedText;
}

function submitForm() {
    const formData = {
        showAs: showAsInput.value,
        URL: urlInput.value
    };

    Office.onReady(function (info) {
      if (info.host === Office.HostType.Word || info.host === Office.HostType.Excel || info.host === Office.HostType.PowerPoint) {
        Office.context.ui.messageParent(JSON.stringify(formData));
      } else {
        console.log("Errore: ambiente Office non riconosciuto");
      }
    });
    
}