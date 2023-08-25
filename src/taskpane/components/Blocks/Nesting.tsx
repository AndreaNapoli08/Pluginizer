import * as React from 'react';
import { useState, useEffect } from 'react'
import Grid from '@mui/material/Grid';
import FormatIndentDecreaseIcon from '@mui/icons-material/FormatIndentDecrease';
import FormatIndentIncreaseIcon from '@mui/icons-material/FormatIndentIncrease';
import IconButton from '@mui/material/IconButton';

export const Nesting = () => {
  // inizializzazione delle variabili per la gestione dei margini
  let leftMargin = 0;
  let rightMargin = 0;
  let interlinea = 5;
  const Nesting = async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.paragraphs.load("style, leftIndent");
      await context.sync();
      
      if(leftMargin < 0){ 
        // vuol dire che il paragrafo selezionato si può spostare verso sinistra
        selection.paragraphs.items.forEach((paragraph) => {
          paragraph.leftIndent -= 30;
        });
        leftMargin += 30;
        rightMargin -= 30;
      }
    });
  }

  const Unnesting = async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.paragraphs.load("style, leftIndent");
      await context.sync();
      
      if(rightMargin <= 400){
        // abbiamo visto che 400 indica la fine del margine destro, quindi se siamo sotto questa soglia possiamo spostarlo verdo destra
        selection.paragraphs.items.forEach((paragraph) => {
          paragraph.leftIndent += 30;
      });
        rightMargin += 30;
        leftMargin -= 30;
      }
    });
  }

  const lineSpacing = async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.paragraphs.load();
      await context.sync();

      // abbiamo impostato un interlinea massimo di 30 e ogni volta che si preme si aumenta di 5
      if(interlinea <= 30){
        interlinea += 5;
      }else{
        interlinea = 10;
      }

      selection.paragraphs.items.forEach((paragraph) => {
        paragraph.lineSpacing = interlinea;
      });

      await context.sync();
    });
  }

    return (
        <div>
            <div style={{display: 'flex', justifyContent: 'center', marginTop: '5px', marginBottom: '5px'}}>
                Nesting:    
            </div>
            <Grid
              container
              direction="row"
              justifyContent="center"
              alignItems="flex-start"
              spacing={1}
            >
              <Grid item xs={4}>
                <IconButton color="inherit" onClick={Nesting}>
                  <FormatIndentDecreaseIcon style={{marginLeft: "20px", position: 'relative', right: '10px'}}/>
                </IconButton>
                <div style={{marginLeft: '7px', fontSize: '10px'}}>move left</div>
              </Grid>
              <Grid item xs={4}>
                <IconButton color="inherit" onClick={Unnesting}>
                  <FormatIndentIncreaseIcon style={{marginLeft: "20px", position: 'relative', right: '10px'}}/>
                </IconButton>
                <div style={{marginLeft: '6px', fontSize: '10px'}}>move right</div>
              </Grid>
              <Grid item xs={4}>
                <IconButton color="inherit" onClick={lineSpacing}>
                  <img title="interstial" width={30} src="assets/interstitial.png" style={{marginLeft: "20px", position: 'relative', right: '10px'}} />
                </IconButton>
                <div style={{marginLeft: '10px', fontSize: '10px'}}>line spacing</div>
              </Grid>
            </Grid>
        </div>
    )
}