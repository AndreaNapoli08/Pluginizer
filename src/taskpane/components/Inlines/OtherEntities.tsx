import * as React from 'react';
import { useState } from 'react';
import Grid from '@mui/material/Grid';
import InputLabel from '@mui/material/InputLabel';
import MenuItem from '@mui/material/MenuItem';
import FormControl from '@mui/material/FormControl';
import Select, { SelectChangeEvent } from '@mui/material/Select';

export const OtherEntities = () => {
    const [concept, setConcept] = useState('');
    
    const handleChangeConcept = (event: SelectChangeEvent) => {
        setConcept(event.target.value);
    };

    return (
        <div>
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
        </div>
    )
}