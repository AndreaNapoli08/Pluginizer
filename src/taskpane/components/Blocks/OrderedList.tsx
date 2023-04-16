import * as React from 'react';
import Grid from '@mui/material/Grid'; 
import IconButton from '@mui/material/IconButton';

export const OrderedList = () => {

  const formatList = async (numberingType) => {
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
        switch(numberingType){
          case "numbers":
            list.setLevelNumbering(0, Word.ListNumbering.arabic);
            break;
          case "lettersUpper":
            list.setLevelNumbering(0, Word.ListNumbering.upperLetter);
            break;
          case "lettersLower":
            list.setLevelNumbering(0, Word.ListNumbering.lowerLetter);
            break;
          case "lettersRomanUpper":
            list.setLevelNumbering(0, Word.ListNumbering.upperRoman);
            break;
          case "lettersRomanLower":
            list.setLevelNumbering(0, Word.ListNumbering.lowerRoman);
            break;
          default:
            break;
        }
        list.load();
        await context.sync();
        for (let i = 1; i < selection.paragraphs.items.length; i++) {
          list.insertParagraph(selection.paragraphs.items[i].text, "End");
          switch(numberingType){
            case "numbers":
              list.setLevelNumbering(0, Word.ListNumbering.arabic);
              break;
            case "lettersUpper":
              list.setLevelNumbering(0, Word.ListNumbering.upperLetter);
              break;
            case "lettersLower":
              list.setLevelNumbering(0, Word.ListNumbering.lowerLetter);
              break;
            case "lettersRomanUpper":
              list.setLevelNumbering(0, Word.ListNumbering.upperRoman);
              break;
            case "lettersRomanLower":
              list.setLevelNumbering(0, Word.ListNumbering.lowerRoman);
              break;
            default:
              break;
          }
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
          <IconButton color="inherit" title="List Numbers" onClick={() => formatList('numbers')}>
            <img width={30} src="../../../assets/listsNumbers.png" />
          </IconButton>
        </Grid>
        <Grid item xs={2.4}>
          <IconButton color="inherit" title="List Letters" onClick={() => formatList('lettersUpper')}>
            <img width={30} src="../../../assets/listsLetters.png" />
          </IconButton>
        </Grid>
        <Grid item xs={2.4}>
          <IconButton color="inherit" title="List Letters Lower" onClick={() => formatList('lettersLower')}>
            <img width={30} src="../../../assets/listsLettersLower.png" />
          </IconButton>
        </Grid>
        <Grid item xs={2.4}>
          <IconButton color="inherit" title="List Letters Roman" onClick={() => formatList('lettersRomanUpper')}>
            <img width={30} src="../../../assets/listsLettersRomans.png" />
          </IconButton>
        </Grid>
        <Grid item xs={2.4}>
          <IconButton color="inherit" title="List Letters Roman Lower" onClick={() => formatList('lettersRomanLower')}>
            <img width={30} src="../../../assets/listsLettersRomansLower.png" />
          </IconButton>
        </Grid>
      </Grid>
    </div>
  )
}