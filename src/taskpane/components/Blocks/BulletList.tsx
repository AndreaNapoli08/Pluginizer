import * as React from 'react';
import Grid from '@mui/material/Grid';
import FormatListBulletedIcon from '@mui/icons-material/FormatListBulleted';    
import IconButton from '@mui/material/IconButton';
import { listItemSecondaryActionClasses } from '@mui/material';

export const BulletList = () => {
  let textListMap = new Map();

  const handleList = async (bulletType) => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const lists = context.document.body.lists;
      selection.load("text, paragraphs");
      await context.sync();

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

        if (previousParagraph.isNullObject || !previousParagraph.isListItem) {
        // vuol dire che la riga selezionata è la prima del documento quindi per forza bisogna creare una lista
        // oppure che nel paragrafo precedente non è già presente una lista
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
          textListMap.set(selection.paragraphs.items[0].text, bulletType);
          list.load();
          await context.sync();
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
              <img width={30} src="assets/bulletlist2.png" />
            </IconButton>
          </Grid>
          <Grid item xs={3}>
            <IconButton color="inherit" title="Square List" onClick={() => handleList('square')} style={{position: 'relative', bottom: '4px'}}>
              <img width={30} src="assets/squared.png" />
            </IconButton>
          </Grid>
          <Grid item xs={3}>
            <IconButton color="inherit" title="Check List" onClick={() => handleList('checkmark')} style={{position: 'relative', bottom: '2px'}}>
              <img width={30} src="assets/checkmark.png" />
            </IconButton>
          </Grid>
        </Grid>
    </div>
  )
}