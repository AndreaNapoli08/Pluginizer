import { FormControl, Grid, InputLabel, MenuItem, Input, IconButton, Button } from '@mui/material';
import Select, { SelectChangeEvent } from '@mui/material/Select';
import * as React from 'react';
import { useState } from 'react';
import AddIcon from '@mui/icons-material/Add';
import { makeStyles } from '@mui/styles';
import { LocalizationProvider } from '@mui/x-date-pickers/LocalizationProvider';
import { AdapterDayjs } from '@mui/x-date-pickers/AdapterDayjs';
import { DatePicker } from '@mui/x-date-pickers/DatePicker';
import DeleteIcon from '@mui/icons-material/Delete';

export const Documents = () => {
    const [resolution, setResolution] = useState('');
    const [identifier, setIdentifier] = useState('');
    const [valDate, setValDate] = useState('');

    const useStyles = makeStyles({
        datePicker: {
            borderRadius: '10px',
        },
    });
    const classes = useStyles();

    const handleChangeDocumentType = async (event: SelectChangeEvent) => {
        setResolution(event.target.value);
    }

    const handleInputChange = (event) => {
        setIdentifier(event.target.value);
    };

    return (
        <div>
            <Grid
                container
                direction="row"
                justifyContent="center"
                alignItems="flex-start"
                spacing={1}
            >
                <Grid item xs={6}>
                    <p style={{ marginLeft: '3px', fontSize: '17px' }}>Document Type</p>
                </Grid>
                <Grid item xs={6}>
                    <FormControl
                        sx={{ m: 1, minWidth: 120 }}
                        size="small"
                    >
                        <InputLabel id="demo-select-small">resolutions</InputLabel>
                        <Select
                            labelId="demo-select-small"
                            id="demo-select-small"
                            value={resolution}
                            label="resolutions"
                            onChange={handleChangeDocumentType}
                        >
                            <MenuItem value="">
                                <em>None</em>
                            </MenuItem>
                            <MenuItem value="decisions">Decisions</MenuItem>
                            <MenuItem value="report">Report</MenuItem>
                            <MenuItem value="agenda">Agenda</MenuItem>
                            <MenuItem value="generic">Generic document</MenuItem>
                        </Select>
                    </FormControl>
                </Grid>
                <span style={{ fontSize: '10px', marginLeft: '10px' }}>
                    The document type for this file is set to "Resolution". Changing it to a different value will
                    reset all the current styles and may render some parts invalid for the new document type.
                </span>
            </Grid>

            <Grid
                container
                direction="row"
                justifyContent="center"
                alignItems="flex-start"
                spacing={1}
            >
                <Grid item xs={6}>
                    <p style={{ marginLeft: '3px', fontSize: '17px' }}>Language</p>
                </Grid>
                <Grid item xs={6}>
                    <FormControl
                        sx={{ m: 1, minWidth: 120 }}
                        size="small"
                    >
                        <InputLabel id="demo-select-small">- Choose -</InputLabel>
                        <Select
                            labelId="demo-select-small"
                            id="demo-select-small"
                            value={resolution}
                            label="resolutions"
                            onChange={handleChangeDocumentType}
                        >
                            <MenuItem value="">
                                <em>None</em>
                            </MenuItem>
                            <MenuItem value="english">English</MenuItem>
                            <MenuItem value="french">French</MenuItem>
                            <MenuItem value="russian">Russian</MenuItem>
                            <MenuItem value="spanish">Spanish</MenuItem>
                            <MenuItem value="arabic">Srabic</MenuItem>
                            <MenuItem value="chinese">Chinese</MenuItem>
                        </Select>
                    </FormControl>
                </Grid>
            </Grid>
            <hr />
            <label htmlFor="identifier" style={{ marginTop: "10px", fontSize: '17px' }}>Official identifier</label>
            <input
                type="text"
                id="identifier"
                name="identifier"
                value={identifier}
                onChange={handleInputChange}
                style={{
                    marginTop: '5px',
                    width: '95%',
                    height: '30px',
                    border: '1px solid black',
                    borderRadius: '5px',
                    paddingLeft: '5px',
                    fontSize: '16px',
                    marginBottom: '10px'
                }}
            />
            <br />
            <label htmlFor="alias" style={{ fontSize: '17px' }}>Alias</label>
            <br />
            <input
                type="text"
                id="alias"
                name="alias"
                value={identifier}
                onChange={handleInputChange}
                style={{
                    marginTop: '5px',
                    width: '82%',
                    height: '30px',
                    border: '1px solid black',
                    borderRadius: '5px',
                    paddingLeft: '5px',
                    fontSize: '16px',
                }}
            />
            <IconButton>
                <AddIcon />
            </IconButton>
            <hr />
            <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '10px', fontSize: '20px' }}>
                Current draft
            </div>
            <Grid
                container
                direction="row"
                justifyContent="center"
                alignItems="flex-start"
                spacing={1}
            >
                <Grid item xs={5}>
                    <p style={{ marginLeft: '3px', fontSize: '17px' }}>Stage</p>
                </Grid>
                <Grid item xs={7}>
                    <FormControl
                        sx={{ m: 1, minWidth: 120 }}
                        size="small"
                        style={{ width: '95%' }}
                    >
                        <InputLabel id="demo-select-small">Informal</InputLabel>
                        <Select
                            labelId="demo-select-small"
                            id="demo-select-small"
                            value={resolution}
                            label="informal"
                            onChange={handleChangeDocumentType}
                        >
                            <MenuItem value="">
                                <em>None</em>
                            </MenuItem>
                            <MenuItem value="first">First draft</MenuItem>
                            <MenuItem value="under">Under revision</MenuItem>
                            <MenuItem value="approved">Approved</MenuItem>
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
                style={{ marginTop: '5px' }}
            >
                <Grid item xs={5}>
                    <p style={{ marginLeft: '3px', fontSize: '17px' }}>Completed</p>
                </Grid>
                <Grid item xs={7}>
                    <LocalizationProvider dateAdapter={AdapterDayjs}>
                        <DatePicker
                            value={valDate}
                            onChange={(newValue) => setValDate(newValue)}
                            className={classes.datePicker}
                        />
                    </LocalizationProvider>
                </Grid>
            </Grid>
            {valDate != "" ?
                <Grid container justifyContent="flex-end">
                    <Button 
                        variant="text" 
                        size="small" 
                        startIcon={<DeleteIcon />}
                        onClick={() => setValDate("")}
                        color="inherit"
                    >
                        Clear
                    </Button>
                </Grid>
            : null}
            <Grid
                container
                direction="row"
                justifyContent="center"
                alignItems="flex-start"
                spacing={1}
                style={{ marginTop: '15px' }}
            >
                <Grid item xs={5}>
                    <label htmlFor="drafter" style={{ fontSize: '17px' }}>Drafter</label>
                </Grid>
                <Grid item xs={7}>
                    <input
                        type="text"
                        id="drafter"
                        name="drafter"
                        value={identifier}
                        onChange={handleInputChange}
                        style={{
                            width: '95%',
                            height: '30px',
                            border: '1px solid black',
                            borderRadius: '5px',
                            paddingLeft: '5px',
                            fontSize: '16px',
                        }}
                    />
                </Grid>
            </Grid>

            <Grid
                container
                direction="row"
                justifyContent="center"
                alignItems="flex-start"
                spacing={1}
                style={{ marginTop: '15px' }}
            >
                <Grid item xs={5}>
                    <p style={{ fontSize: '17px' }}>Office</p>
                </Grid>
                <Grid item xs={7}>
                    <FormControl
                        sx={{ m: 1, minWidth: 120 }}
                        size="small"
                        style={{ width: '95%' }}
                    >
                        <InputLabel id="demo-select-small">External draft</InputLabel>
                        <Select
                            labelId="demo-select-small"
                            id="demo-select-small"
                            value={resolution}
                            label="external draft"
                            onChange={handleChangeDocumentType}
                        >
                            <MenuItem value="">
                                <em>None</em>
                            </MenuItem>
                            <MenuItem value="drafting">Drafting office</MenuItem>
                            <MenuItem value="publication">Publication office</MenuItem>
                        </Select>
                    </FormControl>
                </Grid>
            </Grid>
        </div>
    )
}