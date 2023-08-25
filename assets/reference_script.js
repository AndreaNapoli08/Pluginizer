const queryParams = new URLSearchParams(window.location.search);
const information = queryParams.get('information');
const parsedInformation = JSON.parse(information);

if(parsedInformation != null){
  if(parsedInformation.type == "ref"){
    document.getElementById("reference_ref").checked = true;
    showRefFields();
  }else if(parsedInformation.type == "mref"){
    document.getElementById("reference_mref").checked = true;
    showMRefFields();
  }else{
    document.getElementById("reference_rref").checked = true;
    showRRefFields();
  }
}

function showRefFields() {
  document.getElementById("refFields").style.display = "block";
  document.getElementById("mrefFields").style.display = "none";
  document.getElementById("rrefFields").style.display = "none";
  if(parsedInformation != null){
    document.getElementById("numeroArticoloRef").value = parsedInformation.number;
    document.getElementById("documentoRef").value = parsedInformation.documento;
  }
}

function showMRefFields() {
  document.getElementById("refFields").style.display = "none";
  document.getElementById("mrefFields").style.display = "block";
  document.getElementById("rrefFields").style.display = "none";
  if(parsedInformation != null){
    console.log(parsedInformation.number)
    document.getElementById("numeriArticoliMRef").value = parsedInformation.number;
    document.getElementById("documentoMRef").value = parsedInformation.documento;
  }
}

function showRRefFields() {
  document.getElementById("refFields").style.display = "none";
  document.getElementById("mrefFields").style.display = "none";
  document.getElementById("rrefFields").style.display = "block";
  console.log(parsedInformation)
  if(parsedInformation != null){
    document.getElementById("dalRRef").value = parsedInformation.dal;
    document.getElementById("alRRef").value = parsedInformation.al;
    document.getElementById("documentoRRef").value = parsedInformation.documento;
  }
}

function submitForm() {
  const selectedRiferimento = document.querySelector('input[name="riferimento"]:checked').value;
  let formData;

  switch (selectedRiferimento) {
    case "ref":
      formData = {
        entity: "reference",
        type: "ref",
        numeroArticolo: document.getElementsByName("numeroArticoloRef")[0].value,
        documento: document.getElementsByName("documentoRef")[0].value
      };
      break;
    case "mref":
      formData = {
        entity: "reference",
        type: "mref",
        numeriArticoli: document.getElementsByName("numeriArticoliMRef")[0].value,
        documento: document.getElementsByName("documentoMRef")[0].value
      };
      break;
    case "rref":
      formData = {
        entity: "reference",
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
