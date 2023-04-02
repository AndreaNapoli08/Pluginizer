import * as React from 'react';
import FormatAlignLeftIcon from '@mui/icons-material/FormatAlignLeft';
import FormatAlignCenterIcon from '@mui/icons-material/FormatAlignCenter';
import FormatAlignRightIcon from '@mui/icons-material/FormatAlignRight';
import FormatAlignJustifyIcon from '@mui/icons-material/FormatAlignJustify';
import Grid from '@mui/material/Grid';
import FormatListBulletedIcon from '@mui/icons-material/FormatListBulleted';    
import FormatIndentDecreaseIcon from '@mui/icons-material/FormatIndentDecrease';
import FormatIndentIncreaseIcon from '@mui/icons-material/FormatIndentIncrease';
import GridOnIcon from '@mui/icons-material/GridOn';
import ImageIcon from '@mui/icons-material/Image';
export const Blocks = () => {
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
                <FormatAlignLeftIcon />
              </Grid>
              <Grid item xs={3}>
                <FormatAlignCenterIcon />
              </Grid>
              <Grid item xs={3}>
                <FormatAlignRightIcon />
              </Grid>
              <Grid item xs={3}>
                <FormatAlignJustifyIcon />
              </Grid>
            </Grid>
            <hr />
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
                <FormatListBulletedIcon />
              </Grid>
            </Grid>
            <hr />
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
              <img width={30} src="../../../assets/listsNumbers.png" />
              </Grid>
              <Grid item xs={2.4}>
                <img width={30} src="../../../assets/listsLetters.png" />
              </Grid>
              <Grid item xs={2.4}>
              <img width={30} src="../../../assets/listsLettersLower.png" />
              </Grid>
              <Grid item xs={2.4}>
              <img width={30} src="../../../assets/listsLettersRomans.png" />
              </Grid>
              <Grid item xs={2.4}>
              <img width={30} src="../../../assets/listsLettersRomansLower.png" />
              </Grid>
            </Grid>
            <hr />
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
                <FormatIndentDecreaseIcon style={{marginLeft: "20px"}}/>
                <div style={{marginLeft: '7px', fontSize: '10px'}}>nested list</div>
              </Grid>
              <Grid item xs={4}>
                <FormatIndentIncreaseIcon style={{marginLeft: "20px"}}/>
                <div style={{marginLeft: '6px', fontSize: '10px'}}>unnested list</div>
              </Grid>
              <Grid item xs={4}>
              <img width={30} src="../../../assets/interstitial.png" style={{marginLeft: "20px"}} />
              <div style={{marginLeft: '10px', fontSize: '10px'}}>interstitial</div>
              </Grid>
            </Grid>
            <hr />
            <div style={{display: 'flex', justifyContent: 'center', marginTop: '5px', marginBottom: '5px'}}>
                Other Blocks:    
            </div>
            <Grid
              container
              direction="row"
              justifyContent="center"
              alignItems="flex-start"
              spacing={1}
            >
              <Grid item xs={3}>
                <GridOnIcon fontSize='large'/>
                <div style={{fontSize: '10px'}}>add table</div>
              </Grid>
              <Grid item xs={3}>
                <ImageIcon fontSize='large'/>
                <div style={{fontSize: '10px'}}>add image</div>
              </Grid>
            </Grid>
        </div>
    )
}