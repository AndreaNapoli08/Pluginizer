import * as React from 'react';
import Grid from '@mui/material/Grid';
import FormatListBulletedIcon from '@mui/icons-material/FormatListBulleted';    
import IconButton from '@mui/material/IconButton';

export const BulletList = () => {
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
              <IconButton color="inherit" title="Bulleted List">
                <FormatListBulletedIcon fontSize="small" />
              </IconButton>
              </Grid>
            </Grid>
        </div>
    )
}