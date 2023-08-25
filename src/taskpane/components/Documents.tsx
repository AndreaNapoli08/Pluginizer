import { FormControl, Grid, InputLabel, MenuItem, Input, IconButton, Button } from '@mui/material';
import Select, { SelectChangeEvent } from '@mui/material/Select';
import * as React from 'react';
import { useState, useEffect } from 'react';
import AddIcon from '@mui/icons-material/Add';
import RemoveIcon from '@mui/icons-material/Remove';
import { makeStyles } from '@mui/styles';
import DeleteIcon from '@mui/icons-material/Delete';

export const Documents = () => {
    const [resolution, setResolution] = useState("decisions");
    const [language, setLanguage] = useState("");
    const [stage, setStage] = useState("first")
    const [office, setOffice] = useState("");
    const [identifier, setIdentifier] = useState('');
    const [valDate, setValDate] = useState('');
    const [drafter, setDrafter] = useState("");
    const [inputFields, setInputFields] = useState([]);
    const [showConfirm, setShowConfirm] = useState(false);
    const [showAddAlias, setShowAddAlias] = useState(true);
    const [currentInputValue, setCurrentInputValue] = useState('');
    const [currentIndex, setCurrentIndex] = useState('');
    let contextDocument;
    const NAMESPACE_URI = "prova";

    const deleteInformation = async (type) => {
        if (type === "calendar") {
            setValDate("");
        }
        // Elimina informazione attuale
        Office.context.document.customXmlParts.getByNamespaceAsync(NAMESPACE_URI, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const xmlParts = result.value;
                for (const xmlPart of xmlParts) {
                    xmlPart.getXmlAsync(asyncResult => {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            const xmlData = asyncResult.value;
                            if (xmlData.includes(`${type}=`)) {
                                xmlPart.deleteAsync();
                            }
                        } else {
                            console.error("Errore nel recupero dei contenuti personalizzati");
                        }

                    });
                }
            } else {
                console.error("Errore nel recupero dei contenuti personalizzati");
            }
        });

        await contextDocument.sync();
    }

    const deleteInformationAlias = async (index) => {
        // Elimina informazione attuale
        Office.context.document.customXmlParts.getByNamespaceAsync(NAMESPACE_URI, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const xmlParts = result.value;
                for (const xmlPart of xmlParts) {
                    xmlPart.getXmlAsync(asyncResult => {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            const xmlData = asyncResult.value;
                            console.log(xmlData)
                            console.log(index);
                            if (xmlData.includes(`index="${index}"`) || xmlData.includes(`index=""`)) {
                                console.log("Eliminazioneeee :", xmlData)
                                xmlPart.deleteAsync();
                            }
                        } else {
                            console.error("Errore nel recupero dei contenuti personalizzati");
                        }

                    });
                }
            } else {
                console.error("Errore nel recupero dei contenuti personalizzati");
            }
        });

        await contextDocument.sync();
    }

    const insertInformation = async (xmlData) => {
        // inserimento nuova informazione
        Office.context.document.customXmlParts.addAsync(xmlData, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Dati personalizzati aggiunti con successo");
            } else {
                console.error("Errore durante l'aggiunta dei dati personalizzati");
            }
        });
        await contextDocument.sync();
    }

    const getInformation = async (type) => {
        const values = [...inputFields];
        Office.context.document.customXmlParts.getByNamespaceAsync(NAMESPACE_URI, async (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const xmlParts = result.value;
                for (const xmlPart of xmlParts) {
                    await xmlPart.getXmlAsync(asyncResult => {    // questa istruzione non aspetta il completamento di ciascuna chiamata
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            const xmlData = asyncResult.value;
                            if (xmlData.includes(`${type}=`)) {
                                const parser = new DOMParser();
                                const xmlDoc = parser.parseFromString(xmlData, "text/xml");
                                const dataElement = xmlDoc.querySelector(`data[${type}]`);
                                if (dataElement) {
                                    let jsonData = JSON.parse(dataElement.textContent);
                                    switch (type) {
                                        case "resolutions":
                                            setResolution(jsonData.resolutions);
                                            break;
                                        case "language":
                                            setLanguage(jsonData.language);
                                            break;
                                        case "identifier":
                                            setIdentifier(jsonData.identifier);
                                            break;
                                        case "alias":
                                            values.push({ value: jsonData.alias });
                                            console.log(values);
                                            setInputFields(values);
                                        case "stage":
                                            setStage(jsonData.stage);
                                            break;
                                        case "calendar":
                                            setValDate(jsonData.calendar);
                                            break;
                                        case "office":
                                            setOffice(jsonData.office);
                                            break;
                                    }
                                }
                            }
                        } else {
                            console.error("Errore nel recupero dei contenuti personalizzati");
                        }
                    });
                }
            } else {
                console.error("Errore nel recupero dei contenuti personalizzati");
            }
        });
    }

    useEffect(() => {
        // Funzione asincrona separata per eseguire Word.run
        const runWordAsync = async () => {
            try {
                await Word.run(async (context) => {
                    contextDocument = context;
                    if (Office.context.platform !== Office.PlatformType.OfficeOnline) {
                        let selection = context.document.getSelection();
                        selection.load("text");
                        await context.sync();
                        selection.insertField("Start", Word.FieldType.author)
                        selection.load("text");
                        await context.sync();
                        setDrafter(selection.text);
                        selection.clear();
                    }
                    getInformation("resolutions");
                    getInformation("language");
                    getInformation("identifier");
                    getInformation("alias");
                    getInformation("stage");
                    getInformation("calendar")
                    getInformation("office");
                });
            } catch (error) {
                console.error(error);
            }
        };
        runWordAsync();
    }, []);

    const handleInputChange2 = (index, event) => {
        const values = [...inputFields];
        values[index].value = event.target.value;
        setCurrentInputValue(event.target.value);
        setCurrentIndex(index);
        setInputFields(values);
    };

    const confirm = () => {
        setShowConfirm(false);
        setShowAddAlias(true);
        let jsonData = {
            alias: currentInputValue,
            index: currentIndex
        };
        const xmlData = `<root xmlns="${NAMESPACE_URI}"><data index="${currentIndex}" alias="${currentInputValue}">${JSON.stringify(jsonData)}</data></root>`;
        insertInformation(xmlData);
    }

    const handleAddFields = () => {
        setShowConfirm(true);
        const values = [...inputFields];
        values.push({ value: '' });
        setInputFields(values);
        setShowAddAlias(false);
    };

    const handleRemoveFields = (index) => {
        const values = [...inputFields];
        values.splice(index, 1);
        setInputFields(values);
        deleteInformationAlias(index);
    };

    const useStyles = makeStyles({
        datePicker: {
            borderRadius: '5px',
            border: '1px solid #ccc', // Colore di default del bordo
            transition: 'border-color 0.3s ease-in-out', // Aggiungi una transizione al cambio di colore del bordo
            '& input[type="date"]': {
                border: '1px solid #ccc', // Colore di default del bordo dell'input
                transition: 'border-color 0.3s ease-in-out', // Aggiungi una transizione al cambio di colore del bordo
                '&:focus': {
                    border: '1px solid black', // Cambio di colore del bordo dell'input quando è in focus
                },
            },
        },
    });
    const classes = useStyles();

    const handleChangeDocumentType = async (event: SelectChangeEvent) => {
        let jsonData = {
            resolutions: event.target.value
        };
        const uniqueId = Date.now();
        const xmlData = `<root xmlns="${NAMESPACE_URI}"><data id="${uniqueId}" resolutions="${event.target.value}">${JSON.stringify(jsonData)}</data></root>`;
        deleteInformation("resolutions");
        insertInformation(xmlData);
        setResolution(event.target.value);
    };

    const handleChangeLanguage = async (event: SelectChangeEvent) => {
        let jsonData = {
            language: event.target.value
        };
        const uniqueId = Date.now();
        const xmlData = `<root xmlns="${NAMESPACE_URI}"><data id="${uniqueId}" language="${event.target.value}">${JSON.stringify(jsonData)}</data></root>`;
        deleteInformation("language");
        insertInformation(xmlData);
        setLanguage(event.target.value);
    };

    const handleInputChange = (event) => {
        let jsonData = {
            identifier: event.target.value
        };
        const uniqueId = Date.now();
        const xmlData = `<root xmlns="${NAMESPACE_URI}"><data id="${uniqueId}" identifier="${event.target.value}">${JSON.stringify(jsonData)}</data></root>`;
        deleteInformation("identifier");
        insertInformation(xmlData);
        setIdentifier(event.target.value);
    };

    const handleChangeStage = async (event: SelectChangeEvent) => {
        let jsonData = {
            stage: event.target.value
        };
        const uniqueId = Date.now();
        const xmlData = `<root xmlns="${NAMESPACE_URI}"><data id="${uniqueId}" stage="${event.target.value}">${JSON.stringify(jsonData)}</data></root>`;
        deleteInformation("stage");
        insertInformation(xmlData);
        setStage(event.target.value);
    };

    const handleChangeDate = async (newValue) => {
        let jsonData = {
            calendar: newValue
        };
        const uniqueId = Date.now();
        const xmlData = `<root xmlns="${NAMESPACE_URI}"><data id="${uniqueId}" calendar="${newValue}">${JSON.stringify(jsonData)}</data></root>`;
        deleteInformation("calendar");
        insertInformation(xmlData);
        setValDate(newValue);
    }

    const handleChangeOffice = async (event: SelectChangeEvent) => {
        let jsonData = {
            office: event.target.value
        };
        const uniqueId = Date.now();
        const xmlData = `<root xmlns="${NAMESPACE_URI}"><data id="${uniqueId}" office="${event.target.value}">${JSON.stringify(jsonData)}</data></root>`;
        deleteInformation("office");
        insertInformation(xmlData);
        setOffice(event.target.value);
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
                            value={language}
                            label="resolutions"
                            onChange={handleChangeLanguage}
                        >
                            <MenuItem value="">
                                <em>None</em>
                            </MenuItem>
                            <MenuItem value="english">English</MenuItem>
                            <MenuItem value="french">French</MenuItem>
                            <MenuItem value="russian">Russian</MenuItem>
                            <MenuItem value="spanish">Spanish</MenuItem>
                            <MenuItem value="arabic">Arabic</MenuItem>
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
            {inputFields.map((inputField, index) => (
                <div key={index}>
                    <label htmlFor={`alias-${index}`} style={{ fontSize: '17px' }}>Alias</label>
                    <br />
                    <input
                        type="text"
                        id={`alias-${index}`}
                        name={`alias-${index}`}
                        value={inputField.value}
                        onChange={(event) => handleInputChange2(index, event)}
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
                    <IconButton onClick={() => handleRemoveFields(index)}>
                        <RemoveIcon />
                    </IconButton>
                </div>
            ))}
            {showAddAlias ? <Button color="inherit" onClick={handleAddFields}>
                Add alias
            </Button> : null}

            {showConfirm ? <Button color="primary" onClick={confirm}>
                Confirm
            </Button> : null}

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
                            value={stage}
                            label="informal"
                            onChange={handleChangeStage}
                        >
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
            >
                <Grid item xs={5}>
                    <p style={{ marginLeft: '3px', fontSize: '17px' }}>Completed</p>
                </Grid>
                <Grid item xs={7}>
                    <input
                        title="calendar"
                        type="date"
                        value={valDate}
                        onChange={(event) => handleChangeDate(event.target.value)}
                        className={classes.datePicker}
                        style={{ marginLeft: "7px", marginTop: "9px", height: "38px", width: "150px", fontSize: "15px", textAlign: "center" }}
                    />
                </Grid>
            </Grid>
            {valDate !== "" ? (
                <Grid container justifyContent="flex-end">
                    <Button
                        variant="text"
                        size="small"
                        startIcon={<DeleteIcon />}
                        onClick={() => deleteInformation("calendar")}
                        color="inherit"
                    >
                        Clear
                    </Button>
                </Grid>
            ) : null}
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
                        disabled={true}
                        type="text"
                        id="drafter"
                        name="drafter"
                        value={drafter}
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
                            value={office}
                            label="external draft"
                            onChange={handleChangeOffice}
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