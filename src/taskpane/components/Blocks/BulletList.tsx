// licenza d'uso riservata ad Andrea Napoli e all'università si Bologna
import * as React from 'react';
import Grid from '@mui/material/Grid';
import FormatListBulletedIcon from '@mui/icons-material/FormatListBulleted';    
import IconButton from '@mui/material/IconButton';

export const BulletList = () => {
  let textListMap = new Map(); // Mappa che associa il testo del paragrafo al tipo di bullet scelto

  // Funzione che gestisce la creazione delle liste
  const handleList = async (bulletType) => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text, paragraphs");
      await context.sync();

      // Se il paragrafo selezionato è già una lista, lo rimuove dalla lista
      if (selection.paragraphs.items[0].isListItem) {
        for(let i=0; i<selection.paragraphs.items.length; i++){
          selection.paragraphs.items[i].detachFromList();
          selection.paragraphs.items[i].leftIndent -=36;  // su word desktop una volta tolto dalla lista non si allinea perfettamente a sinistra
        }
      } else {
        const selectedParagraph = selection.paragraphs.getFirstOrNullObject();
        const previousParagraph = selectedParagraph.getPreviousOrNullObject();
        previousParagraph.load("isListItem")
        await context.sync();

        // Se la riga selezionata è la prima del documento o se nel paragrafo precedente non è già presente una lista, crea una nuova lista
        if (previousParagraph.isNullObject || !previousParagraph.isListItem) {
          const list = selection.paragraphs.items[0].startNewList();
          await context.sync();
          switch (bulletType) {
            case 'solid':
              list.setLevelBullet(0, Word.ListBullet.solid);
              break;
            case 'diamonds':
              list.setLevelBullet(0, Word.ListBullet.diamonds);
              break;
            case 'square':
              list.setLevelBullet(0, Word.ListBullet.square);
              break;
            case 'checkmark':
              list.setLevelBullet(0, Word.ListBullet.checkmark);
              break;
            default:
              break;
          }
          // Aggiunge il testo del paragrafo e il tipo di bullet utilizzato alla mappa textListMap
          textListMap.set(selection.paragraphs.items[0].text, bulletType);
          list.load();
          await context.sync();
          // Aggiunge gli eventuali paragrafi successivi alla lista
          for (let i = 1; i < selection.paragraphs.items.length; i++) {
            list.insertParagraph(selection.paragraphs.items[i].text, "End");
            switch (bulletType) {
              case 'solid':
                list.setLevelBullet(0, Word.ListBullet.solid);
                break;
              case 'diamonds':
                list.setLevelBullet(0, Word.ListBullet.diamonds);
                break;
              case 'square':
                list.setLevelBullet(0, Word.ListBullet.square);
                break;
              case 'checkmark':
                list.setLevelBullet(0, Word.ListBullet.checkmark);
                break;
              default:
                break;
            }
            textListMap.set(selection.paragraphs.items[i].text, bulletType);
            selection.paragraphs.items[i].delete();
          }
        }else{
          // se la non è la prima riga del documento e precedentemente c'è una lista, allora attacca questi nuovi paragrafi alla lista esistente
          previousParagraph.load("list")
          await context.sync();
          previousParagraph.list.load("id")
          await context.sync();
          for (let i = 0; i < selection.paragraphs.items.length; i++) {
            selection.paragraphs.items[i].attachToList(previousParagraph.list.id, 0)
            textListMap.set(selection.paragraphs.items[i].text, bulletType);
          }
        }
      await context.sync();
      }
    });
  };

  return (
    <div>
        <div style={{display: 'flex', justifyContent: 'center', marginTop: '5px', marginBottom: '5px'}}>
            Bullet Lists:
        </div>
        <Grid
          container
          direction="row"
          justifyContent="left"
          alignItems="flex-start"
          spacing={2}
        >
          <Grid item xs={3}>
            <IconButton color="inherit" title="Bulleted List" onClick={() => handleList('solid')}>
              <FormatListBulletedIcon fontSize="small" />
            </IconButton>
          </Grid>
          <Grid item xs={3}>
            <IconButton color="inherit" title="Diamond List" onClick={() => handleList('diamonds')} style={{position: 'relative', bottom: '3px'}}>
              <img title="diamond" width={30} src="assets/bulletlist2.png" />
            </IconButton>
          </Grid>
          <Grid item xs={3}>
            <IconButton color="inherit" title="Square List" onClick={() => handleList('square')} style={{position: 'relative', bottom: '4px'}}>
              <img title="square" width={30} src="assets/squared.png" />
            </IconButton>
          </Grid>
          <Grid item xs={3}>
            <IconButton color="inherit" title="Check List" onClick={() => handleList('checkmark')} style={{position: 'relative', bottom: '2px'}}>
              <img title="check" width={30} src="assets/checkmark.png" />
            </IconButton>
          </Grid>
        </Grid>
    </div>
  )
}