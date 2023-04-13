import * as React from 'react';
import { useState } from 'react';
import Grid from '@mui/material/Grid';
import InputLabel from '@mui/material/InputLabel';
import MenuItem from '@mui/material/MenuItem';
import FormControl from '@mui/material/FormControl';
import Select, { SelectChangeEvent } from '@mui/material/Select';

export const Informative = () => {
    const [docType, setDocType] = useState('');

    const handleChangeDocType = (event: SelectChangeEvent) => {
        setDocType(event.target.value);
    };
    
    return (
        <div>
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
        </div>
    )
}