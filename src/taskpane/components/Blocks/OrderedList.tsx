import * as React from 'react';
import Grid from '@mui/material/Grid'; 
import IconButton from '@mui/material/IconButton';

export const OrderedList = () => {
  const listNumbers = async () => {
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
        list.setLevelNumbering(0, Word.ListNumbering.arabic);
        list.load();
        await context.sync();
        for (let i = 1; i < selection.paragraphs.items.length; i++) {
          list.insertParagraph(selection.paragraphs.items[i].text, "End");
          list.setLevelNumbering(0, Word.ListNumbering.arabic);
          selection.paragraphs.items[i].delete();
        }
      }
      await context.sync();
    });
  }
  
  const listLettersUpper = async () => {
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
        list.setLevelNumbering(0, Word.ListNumbering.upperLetter);
        list.load();
        await context.sync();
        for (let i = 1; i < selection.paragraphs.items.length; i++) {
          list.insertParagraph(selection.paragraphs.items[i].text, "End");
          list.setLevelNumbering(0, Word.ListNumbering.upperLetter);
          selection.paragraphs.items[i].delete();
        }
      }
      await context.sync();
    });
  }

  const listLettersLower = async () => {
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
        list.setLevelNumbering(0, Word.ListNumbering.lowerLetter);
        list.load();
        await context.sync();
        for (let i = 1; i < selection.paragraphs.items.length; i++) {
          list.insertParagraph(selection.paragraphs.items[i].text, "End");
          list.setLevelNumbering(0, Word.ListNumbering.lowerLetter);
          selection.paragraphs.items[i].delete();
        }
      }
      await context.sync();
    });
  }

  const listLettersRomanUpper = async () => {
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
        list.setLevelNumbering(0, Word.ListNumbering.upperRoman);
        list.load();
        await context.sync();
        for (let i = 1; i < selection.paragraphs.items.length; i++) {
          list.insertParagraph(selection.paragraphs.items[i].text, "End");
          list.setLevelNumbering(0, Word.ListNumbering.upperRoman);
          selection.paragraphs.items[i].delete();
        }
      }
      await context.sync();
    });
  }

  const listLettersRomanLower = async () => {
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
        list.setLevelNumbering(0, Word.ListNumbering.lowerRoman);
        list.load();
        await context.sync();
        for (let i = 1; i < selection.paragraphs.items.length; i++) {
          list.insertParagraph(selection.paragraphs.items[i].text, "End");
          list.setLevelNumbering(0, Word.ListNumbering.lowerRoman);
          selection.paragraphs.items[i].delete();
        }
      }
      await context.sync();
    });
  }

  return (
    <div>
      <div style={{display: 'flex', justifyContent: 'center', marginTop: '5px', marginBottom: '5px'}}>
          Ordered Lists:
      </div>
      <Grid
        container
        direction="row"
        justifyContent="center"
        alignItems="flex-start"
        spacing={2}
      >
        <Grid item xs={2.4}>
          <IconButton color="inherit" title="List Numbers" onClick={listNumbers}>
            <img width={30} src="../../../assets/listsNumbers.png" />
          </IconButton>
        </Grid>
        <Grid item xs={2.4}>
          <IconButton color="inherit" title="List Letters" onClick={listLettersUpper}>
            <img width={30} src="../../../assets/listsLetters.png" />
          </IconButton>
        </Grid>
        <Grid item xs={2.4}>
          <IconButton color="inherit" title="List Letters Lower" onClick={listLettersLower}>
            <img width={30} src="../../../assets/listsLettersLower.png" />
          </IconButton>
        </Grid>
        <Grid item xs={2.4}>
          <IconButton color="inherit" title="List Letters Roman" onClick={listLettersRomanUpper}>
            <img width={30} src="../../../assets/listsLettersRomans.png" />
          </IconButton>
        </Grid>
        <Grid item xs={2.4}>
          <IconButton color="inherit" title="List Letters Roman Lower" onClick={listLettersRomanLower}>
            <img width={30} src="../../../assets/listsLettersRomansLower.png" />
          </IconButton>
        </Grid>
      </Grid>
    </div>
  )
}