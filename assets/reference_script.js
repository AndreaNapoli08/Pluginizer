function showRefFields() {
  document.getElementById("refFields").style.display = "block";
  document.getElementById("mrefFields").style.display = "none";
  document.getElementById("rrefFields").style.display = "none";
}

function showMRefFields() {
  document.getElementById("refFields").style.display = "none";
  document.getElementById("mrefFields").style.display = "block";
  document.getElementById("rrefFields").style.display = "none";
}

function showRRefFields() {
  document.getElementById("refFields").style.display = "none";
  document.getElementById("mrefFields").style.display = "none";
  document.getElementById("rrefFields").style.display = "block";
}

function submitForm() {
  const selectedRiferimento = document.querySelector('input[name="riferimento"]:checked').value;
  let formData;

  switch (selectedRiferimento) {
    case "ref":
      formData = {
        type: "ref",
        numeroArticolo: document.getElementsByName("numeroArticoloRef")[0].value,
        documento: document.getElementsByName("documentoRef")[0].value
      };
      break;
    case "mref":
      formData = {
        type: "mref",
        numeriArticoli: document.getElementsByName("numeriArticoliMRef")[0].value,
        documento: document.getElementsByName("documentoMRef")[0].value
      };
      break;
    case "rref":
      formData = {
        type: "rref",
        dal: document.getElementsByName("dalRRef")[0].value,
        al: document.getElementsByName("alRRef")[0].value,
        documento: document.getElementsByName("documentoRRef")[0].value
      };
      break;
  }

  Office.onReady(function (info) {
    if (info.host === Office.HostType.Word || info.host === Office.HostType.Excel || info.host === Office.HostType.PowerPoint) {
      // Invia i dati all'add-in utilizzando Office.context.ui.messageParent
      Office.context.ui.messageParent(JSON.stringify(formData));
    } else {
      // Se l'add-in non è eseguito in un ambiente Office corretto, gestisci il caso di errore
      console.log("Errore: ambiente Office non riconosciuto");
    }
  });

}
