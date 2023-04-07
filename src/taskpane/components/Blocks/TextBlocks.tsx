import * as React from 'react';
import Grid from '@mui/material/Grid';
import FormatAlignLeftIcon from '@mui/icons-material/FormatAlignLeft';
import FormatAlignCenterIcon from '@mui/icons-material/FormatAlignCenter';
import FormatAlignRightIcon from '@mui/icons-material/FormatAlignRight';
import FormatAlignJustifyIcon from '@mui/icons-material/FormatAlignJustify';
import IconButton from '@mui/material/IconButton';

export const TextBlocks = () => {
    const AlignLeft = async () => {
        await Word.run(async (context) => {
          const selection = context.document.getSelection();
          selection.paragraphs.load();
          await context.sync();
    
          selection.paragraphs.items.forEach((paragraph) => {
            paragraph.alignment = "Left";
          });
    
          await context.sync();
        });
    };

    const AlignCenter = async () => {
        await Word.run(async (context) => {
          const selection = context.document.getSelection();
          selection.paragraphs.load();
          await context.sync();
    
          selection.paragraphs.items.forEach((paragraph) => {
            paragraph.alignment = "Centered";
          });
    
          await context.sync();
        });
    };

    const AlignRight = async () => {
        await Word.run(async (context) => {
          const selection = context.document.getSelection();
          selection.paragraphs.load();
          await context.sync();
    
          selection.paragraphs.items.forEach((paragraph) => {
            paragraph.alignment = "Right";
          });
    
          await context.sync();
        });
    };

    const Justify = async () => {
        await Word.run(async (context) => {
          const selection = context.document.getSelection();
          selection.paragraphs.load();
          await context.sync();
    
          selection.paragraphs.items.forEach((paragraph) => {
            paragraph.alignment = "Justified";
          });
    
          await context.sync();
        });
    };
    return (
        <div>
            <div style={{display: 'flex', justifyContent: 'center', marginBottom: '5px'}}>
                Text blocks:
            </div>
            <Grid
              container
              direction="row"
              justifyContent="center"
              alignItems="flex-start"
              spacing={2}
            >
              <Grid item xs={3}>
              <IconButton color="inherit" title="Align Left" onClick={AlignLeft}>
                <FormatAlignLeftIcon fontSize="small" />
              </IconButton>
              </Grid>
              <Grid item xs={3}>
              <IconButton color="inherit" title="Align Center" onClick={AlignCenter}>
                <FormatAlignCenterIcon fontSize="small" />
              </IconButton>
              </Grid>
              <Grid item xs={3}>
              <IconButton color="inherit" title="Align Right" onClick={AlignRight}>
                <FormatAlignRightIcon fontSize="small" />
              </IconButton>
              </Grid>
              <Grid item xs={3}>
              <IconButton color="inherit" title="Justify" onClick={Justify}>
                <FormatAlignJustifyIcon fontSize="small" />
              </IconButton>
              </Grid>
            </Grid>
        </div>
    )
}