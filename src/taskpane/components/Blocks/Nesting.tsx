import * as React from 'react';
import { useState, useEffect } from 'react'
import Grid from '@mui/material/Grid';
import FormatIndentDecreaseIcon from '@mui/icons-material/FormatIndentDecrease';
import FormatIndentIncreaseIcon from '@mui/icons-material/FormatIndentIncrease';
import IconButton from '@mui/material/IconButton';

export const Nesting = () => {
  let leftMargin = 0;
  let rightMargin = 0;
  let interlinea = 5;
  const Nesting = async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.paragraphs.load("style, leftIndent");
      await context.sync();
      
      if(leftMargin < 0){
        selection.paragraphs.items.forEach((paragraph) => {
          paragraph.leftIndent -= 30;
        });
        leftMargin += 30;
        rightMargin -= 30;
      }
      
      console.log("leftMargin: ", leftMargin)
      console.log("rightMargin: ", rightMargin)
    });
  }

  const Unnesting = async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.paragraphs.load("style, leftIndent");
      await context.sync();
      
      if(rightMargin <= 400){
        selection.paragraphs.items.forEach((paragraph) => {
          paragraph.leftIndent += 30;
      });
        rightMargin += 30;
        leftMargin -= 30;
      }
      console.log("leftMargin: ", leftMargin)
      console.log("rightMargin: ", rightMargin)
    });
  }

  const lineSpacing = async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.paragraphs.load();
      await context.sync();

      if(interlinea <= 30){
        interlinea += 5;
      }else{
        interlinea = 10;
      }

      selection.paragraphs.items.forEach((paragraph) => {
        paragraph.lineSpacing = interlinea;
      });

      console.log(interlinea)
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
                  <img width={30} src="../../../assets/interstitial.png" style={{marginLeft: "20px", position: 'relative', right: '10px'}} />
                </IconButton>
                <div style={{marginLeft: '10px', fontSize: '10px'}}>line spacing</div>
              </Grid>
            </Grid>
        </div>
    )
}