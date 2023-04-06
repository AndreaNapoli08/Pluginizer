import * as React from 'react';
import { useState } from 'react';
import Grid from '@mui/material/Grid';
import InputLabel from '@mui/material/InputLabel';
import MenuItem from '@mui/material/MenuItem';
import FormControl from '@mui/material/FormControl';
import Select, { SelectChangeEvent } from '@mui/material/Select';
import IconButton from '@mui/material/IconButton';
import LinkIcon from '@mui/icons-material/Link';
import LiveHelpIcon from '@mui/icons-material/LiveHelp';
import NoteAltIcon from '@mui/icons-material/NoteAlt';
import CalendarMonthIcon from '@mui/icons-material/CalendarMonth';
import FolderOpenIcon from '@mui/icons-material/FolderOpen';
import PersonIcon from '@mui/icons-material/Person';
import LocationOnIcon from '@mui/icons-material/LocationOn';
import AccessTimeIcon from '@mui/icons-material/AccessTime';
import FormControlLabel from '@mui/material/FormControlLabel';
import Checkbox from '@mui/material/Checkbox';
import Button from '@mui/material/Button';
import {TipografiaButton} from './TipografiaButton'

export const Inlines = () => {
  const [concept, setConcept] = useState('');
  const [docType, setDocType] = useState('');
  const [expandWords, setExpandWords] = useState(true);

  const handleChangeConcept = (event: SelectChangeEvent) => {
    setConcept(event.target.value);
  };

  const handleChangeDocType = (event: SelectChangeEvent) => {
    setDocType(event.target.value);
  };

  const handleChangeCheckbox = (event: React.ChangeEvent<HTMLInputElement>) => {
    setExpandWords(event.target.checked);
  };
    return (
        <div>
            <div style={{marginBottom:"15px"}}>
                <span style={{fontSize: "18px"}}>Reference</span>
                <LinkIcon style={{marginLeft: "10px"}} />
            </div>
            <div style={{marginBottom:"15px"}}>
                <span style={{fontSize: "18px"}}>Definition</span>
                <LiveHelpIcon style={{marginLeft: "10px"}} />
            </div>
            <div style={{marginBottom:"15px"}}>
                <span style={{fontSize: "18px"}}>Footnote</span>
                <NoteAltIcon style={{marginLeft: "10px"}} />
            </div>
            <hr />
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
                <IconButton color="inherit">
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
            <Grid
              container
              direction="row"
              justifyContent="center"
              alignItems="flex-start"
              spacing={1}
              style={{marginTop: "10px"}}
            >
              <Grid item xs={6}>
                <p style={{marginLeft: '3px', marginTop: '16px', fontSize: '17px'}}>Other Entities</p>
              </Grid>
              <Grid item xs={6}>
              <FormControl 
                sx={{ m: 1, minWidth: 120 }} 
                size="small"
                style={{position: 'relative', right: '18px'}}
              >
                <InputLabel id="demo-select-small">concept</InputLabel>
                <Select
                  labelId="demo-select-small"
                  id="demo-select-small"
                  value={concept}
                  label="concept"
                  onChange={handleChangeConcept}
                >
                  <MenuItem value="">
                    <em>None</em>
                  </MenuItem>
                  <MenuItem value="object">Object</MenuItem>
                  <MenuItem value="event">Event</MenuItem>
                  <MenuItem value="process">Process</MenuItem>
                  <MenuItem value="role">Role</MenuItem>
                  <MenuItem value="term">Term</MenuItem>
                  <MenuItem value="quantity">Quantity</MenuItem>
                </Select>
              </FormControl>
              </Grid>
            </Grid>
            
            <Grid
              container
              direction="row"
              justifyContent="center"
              alignItems="flex-start"
              spacing={1}
              style={{marginTop: "5px"}}
            >
              <Grid item xs={6}>
                <p style={{marginLeft: '18px', marginTop: '16px', fontSize: '17px'}}>Informative</p>
              </Grid>
              <Grid item xs={6}>
              <FormControl 
                sx={{ m: 1, minWidth: 120 }} 
                size="small"
                style={{position: 'relative', right: '18px'}}
              >
                <InputLabel id="demo-select-small">docType</InputLabel>
                <Select
                  labelId="demo-select-small"
                  id="demo-select-small"
                  value={docType}
                  label="docType"
                  onChange={handleChangeDocType}
                >
                  <MenuItem value="">
                    <em>None</em>
                  </MenuItem>
                  <MenuItem value="docTitle">docTitle</MenuItem>
                  <MenuItem value="docNumber">docNumber</MenuItem>
                  <MenuItem value="docProponent">docProponent</MenuItem>
                  <MenuItem value="docDate">docDate</MenuItem>
                  <MenuItem value="session">session</MenuItem>
                  <MenuItem value="shortTitle">shortTitle</MenuItem>
                  <MenuItem value="docAuthority">docAuthority</MenuItem>
                  <MenuItem value="docPurpose">docPurpose</MenuItem>
                  <MenuItem value="docCommittee">docCommittee</MenuItem>
                  <MenuItem value="docIntroducer">docIntroducer</MenuItem>
                  <MenuItem value="docStage">docStage</MenuItem>
                  <MenuItem value="docStatus">docStatus</MenuItem>
                  <MenuItem value="docJurisdiction">docJurisdiction</MenuItem>
                  <MenuItem value="docketNumber">docketNumber</MenuItem>
                </Select>
              </FormControl>
              </Grid>
            </Grid>
            <hr />
            <div style={{display: 'flex', justifyContent: 'center', marginBottom: '5px', fontSize: '20px'}}>
                Tipography    
            </div>
            <TipografiaButton />
            <FormControlLabel 
              control={<Checkbox checked={expandWords} onChange={handleChangeCheckbox}/>} 
              label="Expand to whole words" 
              style={{display: 'flex', justifyContent: 'center', alignItems: 'center', marginTop: '10px'}}
            />
            <div style={{ display: "flex", justifyContent: "center", alignItems: "center", marginTop: '10px' }}>
              <Button variant="outlined" color="inherit">
                Apply to all instances
              </Button>
            </div>
        </div>
    )
}