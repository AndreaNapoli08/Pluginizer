import * as React from 'react';
import Grid from '@mui/material/Grid';
import GridOnIcon from '@mui/icons-material/GridOn';
import ImageIcon from '@mui/icons-material/Image';
import IconButton from '@mui/material/IconButton';

export const OtherBlocks = () => {
    return (
        <div>
            <div style={{display: 'flex', justifyContent: 'center', marginTop: '5px', marginBottom: '5px'}}>
                Other Blocks:    
            </div>
            <Grid
              container
              direction="row"
              justifyContent="center"
              alignItems="flex-start"
              spacing={2}
            >
              <Grid item xs={3} style={{position: 'relative', left: '20px'}}>
                <IconButton color="inherit">
                  <GridOnIcon fontSize='large'/>
                </IconButton>
                <div style={{fontSize: '10px'}}>add table</div>
              </Grid>
              <Grid item xs={3}>
                <IconButton color="inherit">
                  <ImageIcon fontSize='large'/>
                </IconButton>
                <div style={{fontSize: '10px'}}>add image</div>
              </Grid>
            </Grid>
        </div>
    )
}