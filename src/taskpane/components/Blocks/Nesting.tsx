import * as React from 'react';
import Grid from '@mui/material/Grid';
import FormatIndentDecreaseIcon from '@mui/icons-material/FormatIndentDecrease';
import FormatIndentIncreaseIcon from '@mui/icons-material/FormatIndentIncrease';
import IconButton from '@mui/material/IconButton';

export const Nesting = () => {
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
                <IconButton color="inherit">
                  <FormatIndentDecreaseIcon style={{marginLeft: "20px", position: 'relative', right: '10px'}}/>
                </IconButton>
                <div style={{marginLeft: '7px', fontSize: '10px'}}>nested list</div>
              </Grid>
              <Grid item xs={4}>
                <IconButton color="inherit">
                  <FormatIndentIncreaseIcon style={{marginLeft: "20px", position: 'relative', right: '10px'}}/>
                </IconButton>
                <div style={{marginLeft: '6px', fontSize: '10px'}}>unnested list</div>
              </Grid>
              <Grid item xs={4}>
                <IconButton color="inherit">
                  <img width={30} src="../../../assets/interstitial.png" style={{marginLeft: "20px", position: 'relative', right: '10px'}} />
                </IconButton>
                <div style={{marginLeft: '10px', fontSize: '10px'}}>interstitial</div>
              </Grid>
            </Grid>
        </div>
    )
}