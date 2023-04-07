import * as React from 'react';
import Grid from '@mui/material/Grid'; 
import IconButton from '@mui/material/IconButton';

export const OrderedList = () => {
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
                <IconButton color="inherit" title="List Numbers">
                  <img width={30} src="../../../assets/listsNumbers.png" />
                </IconButton>
              </Grid>
              <Grid item xs={2.4}>
                <IconButton color="inherit" title="List Letters">
                  <img width={30} src="../../../assets/listsLetters.png" />
                </IconButton>
              </Grid>
              <Grid item xs={2.4}>
                <IconButton color="inherit" title="List Letters Lower">
                  <img width={30} src="../../../assets/listsLettersLower.png" />
                </IconButton>
              </Grid>
              <Grid item xs={2.4}>
                <IconButton color="inherit" title="List Letters Roman">
                  <img width={30} src="../../../assets/listsLettersRomans.png" />
                </IconButton>
              </Grid>
              <Grid item xs={2.4}>
                <IconButton color="inherit" title="List Letters Roman Lower">
                  <img width={30} src="../../../assets/listsLettersRomansLower.png" />
                </IconButton>
              </Grid>
            </Grid>
        </div>
    )
}