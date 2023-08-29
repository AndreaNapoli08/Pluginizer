// licenza d'uso riservata ad Andrea Napoli e all'università si Bologna
import { Button, FormControl, FormControlLabel, FormLabel, Grid, InputLabel, MenuItem, Radio, RadioGroup, Select, SelectChangeEvent } from '@mui/material';
import * as React from 'react';
import { useState } from 'react';
import { styled } from '@mui/material/styles';

export const Structures = () => {
    const [structure, setStructure] = useState('');

    const handleChangeBody = async (event: SelectChangeEvent) => {
        setStructure(event.target.value);
    }

    const LowercaseButton = styled(Button)({
        textTransform: 'lowercase',
    });

    return (
        <div>
            <p style={{ fontSize: '10px', marginBottom: '10px' }}>Selection is inside structure: "body"</p>
            <Grid
                container
                direction="row"
                justifyContent="center"
                alignItems="flex-start"
                spacing={1}
            >
                <Grid item xs={6}>
                    <p style={{ marginLeft: '3px', fontSize: '17px' }}>Convert to:</p>
                </Grid>
                <Grid item xs={6}>
                    <FormControl
                        sx={{ m: 1, minWidth: 120 }}
                        size="small"
                    >
                        <InputLabel id="demo-select-small">Structure</InputLabel>
                        <Select
                            labelId="demo-select-small"
                            id="demo-select-small"
                            value={structure}
                            label="structure"
                            onChange={handleChangeBody}
                        >
                            <MenuItem value="">
                                <em>None</em>
                            </MenuItem>
                            <MenuItem value="preface">Preface</MenuItem>
                            <MenuItem value="preamble">Preamble</MenuItem>
                            <MenuItem value="introduction">Introduction</MenuItem>
                            <MenuItem value="body">Body</MenuItem>
                            <MenuItem value="conclusions">Conclusions</MenuItem>
                        </Select>
                    </FormControl>
                </Grid>
                <span style={{ fontSize: '10px', marginLeft: '10px' }}>
                    Converting to a different structure type may render inner sections invalid.
                </span>
            </Grid>

            <Grid
                container
                direction="row"
                justifyContent="center"
                alignItems="flex-start"
                spacing={1}
                style={{ marginTop: '10px' }}
            >
                <Grid item xs={7}>
                    <LowercaseButton variant="outlined" color="inherit">Insert structure</LowercaseButton>
                </Grid>
                <Grid item xs={5}>
                    <LowercaseButton variant="outlined" color="inherit">Dissolve</LowercaseButton>
                </Grid>
            </Grid>

            <Grid
                container
                direction="row"
                justifyContent="center"
                alignItems="flex-start"
                spacing={1}
                style={{ marginTop: '10px' }}
            >
                <Grid item xs={6}>
                    <LowercaseButton variant="outlined" color="inherit">Split at cursor</LowercaseButton>
                </Grid>
                <Grid item xs={6}>
                    <LowercaseButton variant="outlined" color="inherit">Join to previous</LowercaseButton>
                </Grid>
            </Grid>

            <hr />

            <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '10px', fontSize: '20px' }}>
                Current draft
            </div>

            <FormControl>
                <RadioGroup
                    aria-labelledby="demo-radio-buttons-group-label"
                    defaultValue="female"
                    name="radio-buttons-group"
                >
                    <Grid
                        container
                        direction="row"
                        justifyContent="center"
                        alignItems="flex-start"
                        spacing={1}
                    >
                        <Grid item xs={4}>
                            <img src="assets/none.png" width={60} style={{ marginTop: '5px' }} title="none" />
                        </Grid>
                        <Grid item xs={8}>
                            <FormControlLabel value="none" control={<Radio />} label="None" />
                        </Grid>
                    </Grid>

                    <Grid
                        container
                        direction="row"
                        justifyContent="center"
                        alignItems="flex-start"
                        spacing={1}
                    >
                        <Grid item xs={4}>
                            <img src="assets/bullet.png" width={60} style={{ marginTop: '5px' }} title="none" />
                        </Grid>
                        <Grid item xs={8}>
                            <FormControlLabel value="bullet" control={<Radio />} label="Bullet" />
                        </Grid>
                    </Grid>

                    <Grid
                        container
                        direction="row"
                        justifyContent="center"
                        alignItems="flex-start"
                        spacing={1}
                    >
                        <Grid item xs={4}>
                            <img src="assets/only_num.png" width={60} style={{ marginTop: '5px' }} title="none" />
                        </Grid>
                        <Grid item xs={8}>
                            <FormControlLabel value="only num" control={<Radio />} label="Only num" />
                        </Grid>
                    </Grid>

                    <Grid
                        container
                        direction="row"
                        justifyContent="center"
                        alignItems="flex-start"
                        spacing={1}
                    >
                        <Grid item xs={4}>
                            <img src="assets/indented.png" width={60} style={{ marginTop: '5px' }} title="none" />
                        </Grid>
                        <Grid item xs={8}>
                            <FormControlLabel value="indented" control={<Radio />} label="Only num - indented" />
                        </Grid>
                    </Grid>

                    <Grid
                        container
                        direction="row"
                        justifyContent="center"
                        alignItems="flex-start"
                        spacing={1}
                    >
                        <Grid item xs={4}>
                            <img src="assets/heading.png" width={60} style={{ marginTop: '5px' }} title="none" />
                        </Grid>
                        <Grid item xs={8}>
                            <FormControlLabel value="heading" control={<Radio />} label="Only heading" />
                        </Grid>
                    </Grid>

                    <Grid
                        container
                        direction="row"
                        justifyContent="center"
                        alignItems="flex-start"
                        spacing={1}
                    >
                        <Grid item xs={4}>
                            <img src="assets/inline.png" width={60} style={{ marginTop: '5px' }} title="none" />
                        </Grid>
                        <Grid item xs={8}>
                            <FormControlLabel value="inline" control={<Radio />} label="Num+heading - in line" />
                        </Grid>
                    </Grid>

                    <Grid
                        container
                        direction="row"
                        justifyContent="center"
                        alignItems="flex-start"
                        spacing={1}
                    >
                        <Grid item xs={4}>
                            <img src="assets/stacked.png" width={60} style={{ marginTop: '5px' }} title="none" />
                        </Grid>
                        <Grid item xs={8}>
                            <FormControlLabel value="stacked" control={<Radio />} label="Num + heading - stacked" />
                        </Grid>
                    </Grid>
                </RadioGroup>
            </FormControl>

        </div>
    )
}