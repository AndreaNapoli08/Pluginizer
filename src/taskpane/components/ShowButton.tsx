import * as React from 'react';
import { Button } from '@fluentui/react-components';
import { DefaultButton } from '@fluentui/react/lib/Button';
import {
  bundleIcon,
  CalendarMonthFilled,
  CalendarMonthRegular,
} from "@fluentui/react-icons";

const CalendarMonth = bundleIcon(CalendarMonthFilled, CalendarMonthRegular);

export interface IWordSelectionState extends React.ComponentState {
  selectedText: string;
  dis: boolean;
}

export class WordSelection extends React.Component<{}, IWordSelectionState> {

  public constructor(props: {}) {
    super(props);
    this.state = {
      selectedText: '',
      dis:true,
      isHovered: false, 
      buttonColor: "transparent"
    };
    this.handleMouseEnter = this.handleMouseEnter.bind(this);
    this.handleMouseLeave = this.handleMouseLeave.bind(this);
  }

  handleMouseEnter() {
    this.setState({ isHovered: true, buttonColor:"lightgrey" });
  }

  handleMouseLeave() {
    this.setState({ isHovered: false, buttonColor:"transparent" });
  }

  componentDidMount() {
    
    Word.run(async (context) => {
      await this.setDefaultFont();
      // Ottenere il testo selezionato dal documento
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        () => {
          Word.run(async (context) => {
            const newSelection = context.document.getSelection();
            newSelection.load("text");
            await context.sync();
            const newSelectedText = newSelection.text;
            if(newSelectedText.length === 0){
              this.setState({
                selectedText: "Nessun testo selezionato",
                dis:true
              });
            }else{
              this.setState({
                selectedText: newSelectedText,
                dis:false
              });
            } 
          });
        });
      return context.sync();
    });
  }
  
  public boldText = async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text, font");
      await context.sync();
      if (selection.isNullObject) {
        return;
      }
      if (selection.font.bold) {
        selection.font.bold = false;
      } else {
        selection.font.bold = true;
      }
      await context.sync();
    });
  } 

  public italicText = async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text, font");
      await context.sync();
      if (selection.isNullObject) {
        return;
      }
      if (selection.font.italic) {
        selection.font.italic = false;
      } else {
        selection.font.italic = true;
      }
      await context.sync();
    });
  } 

  public async setDefaultFont() {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("font");
      await context.sync();
      body.font.underline = "None";
      await context.sync();
      console.log("font default: ", body.font.underline)
    });
  }

  public underlineText = async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text, font");
      await context.sync();
      console.log(selection.font.underline)
      if (selection.font.underline === "Mixed" || selection.font.underline === "None") {
        selection.font.underline = "Single";
        await context.sync();
        console.log(selection.font.underline)
      } else {  
        selection.font.underline = "None";
        await context.sync();
        console.log(selection.font.underline)
      }
    });
  } 
  
  public render() {
    return (
      <div>
        <p>Selected text: {this.state.selectedText}</p>
        <DefaultButton 
          disabled={this.state.dis} 
          style = {{
            marginRight: "10px",
          }}
          onClick={ this.boldText }>
            <b>G</b>
        </DefaultButton>
        <DefaultButton 
          disabled={this.state.dis} 
          style = {{
            marginRight: "10px",
          }}
          onClick={ this.italicText }>
            <i>I</i>
        </DefaultButton>
        <DefaultButton 
          disabled={this.state.dis} 
          onClick={ this.underlineText }>
            <u>SS</u>
        </DefaultButton>
      </div>
    );
  }
}