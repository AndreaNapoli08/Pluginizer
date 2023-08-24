function submitForm() {
    const formData = {
        entity: "footnote",
        definition: document.getElementById('definition').value
    };

    Office.onReady(function (info) {
      if (info.host === Office.HostType.Word || info.host === Office.HostType.Excel || info.host === Office.HostType.PowerPoint) {
        Office.context.ui.messageParent(JSON.stringify(formData));
      } else {
        console.log("Errore: ambiente Office non riconosciuto");
      }
    });
    
}