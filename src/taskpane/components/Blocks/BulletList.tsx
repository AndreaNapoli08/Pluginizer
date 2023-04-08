import * as React from 'react';
import Grid from '@mui/material/Grid';
import FormatListBulletedIcon from '@mui/icons-material/FormatListBulleted';    
import IconButton from '@mui/material/IconButton';

export const BulletList = () => {

  const createBulletList = async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.paragraphs.load();
      await context.sync();
      
      if(selection.paragraphs.items[0].isListItem){
        for (let i = 0; i < selection.paragraphs.items.length; i++) {
          selection.paragraphs.items[i].delete();
        }
        for (let i = 0; i < selection.paragraphs.items.length; i++) {
          selection.insertText(selection.paragraphs.items[i].text + "\n" , "Before")
        }
      }else{
        const list = selection.paragraphs.items[0].startNewList();
        list.setLevelBullet(0, Word.ListBullet.solid);
        list.load();
        await context.sync();
        for (let i = 1; i < selection.paragraphs.items.length; i++) {
          list.insertParagraph(selection.paragraphs.items[i].text, "End");
          list.setLevelBullet(0, Word.ListBullet.solid);
          selection.paragraphs.items[i].delete();
        }
      }
      await context.sync();
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
          <IconButton color="inherit" title="Bulleted List" onClick={createBulletList}>
            <FormatListBulletedIcon fontSize="small" />
          </IconButton>
          </Grid>
        </Grid>
    </div>
  )
}