import * as React from 'react';
import { useState } from 'react';
import Grid from '@mui/material/Grid';
import IconButton from '@mui/material/IconButton';
import CalendarMonthIcon from '@mui/icons-material/CalendarMonth';
import FolderOpenIcon from '@mui/icons-material/FolderOpen';
import PersonIcon from '@mui/icons-material/Person';
import LocationOnIcon from '@mui/icons-material/LocationOn';
import AccessTimeIcon from '@mui/icons-material/AccessTime';
import { Client } from '@microsoft/microsoft-graph-client';

export const ImportantEntities = () => {
    
    const DateStyle = async () => { 
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load();
            await context.sync();
            
            
            
        });
    }
  return (
    <div>
        <div style={{display: 'flex', justifyContent: 'center', marginBottom: '5px', fontSize: '20px'}}>
                Entities    
        </div>
        <Grid
            container
            direction="row"
            justifyContent="center"
            alignItems="flex-start"
            spacing={2}
        >
            <Grid item xs={2.4}>
            <IconButton color="inherit" onClick={DateStyle}>
                <CalendarMonthIcon fontSize="large" />
            </IconButton>
            <div style={{fontSize: '10px', position: 'relative', left: '12px'}}>Date</div>
            </Grid>
            <Grid item xs={2.4}>
            <IconButton color="inherit">
                <FolderOpenIcon fontSize="large" />
            </IconButton>
            <div style={{fontSize: '10px', position: 'relative', right: '6px'}}>Organization</div>
            </Grid>
            <Grid item xs={2.4}>
            <IconButton color="inherit">
                <PersonIcon fontSize="large" />
            </IconButton>
            <div style={{fontSize: '10px', position: 'relative', left: '10px'}}>Person</div>
            </Grid>
            <Grid item xs={2.4}>
            <IconButton color="inherit">
                <LocationOnIcon fontSize="large" />
            </IconButton>
            <div style={{fontSize: '10px', position: 'relative', left: '7px'}}>Location</div>
            </Grid>
            <Grid item xs={2.4}>
            <IconButton color="inherit">
                <AccessTimeIcon fontSize="large" />
            </IconButton>
            <div style={{fontSize: '10px', position: 'relative', left: '12px'}}>Time</div>
            </Grid>
        </Grid>
    </div>
  )
}