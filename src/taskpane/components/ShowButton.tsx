import * as React from 'react';
import { Button } from '@fluentui/react-components';
import {
  bundleIcon,
  CalendarMonthFilled,
  CalendarMonthRegular,
} from "@fluentui/react-icons";


const CalendarMonth = bundleIcon(CalendarMonthFilled, CalendarMonthRegular);

export interface IWordSelectionState extends React.ComponentState {
  selectedText: string;
  dis: boolean;
  buttonColor: string;
}

export class WordSelection extends React.Component<{}, IWordSelectionState> {

  public constructor(props: {}) {
    super(props);
    this.state = {
      selectedText: '',
      dis:true,
      buttonColor: 'transparent'
    };
  }

  componentDidMount() {
    Word.run(async (context) => {
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
        this.setState({
          buttonColor: "transparent"
        });
      } else {
        selection.font.bold = true;
        this.setState({
          buttonColor: "red"
        });
      }
      await context.sync();
    });
  } 
  
  public render() {
    return (
      <div>
        <p>Selected text: {this.state.selectedText}</p>
        <Button 
          id="ciao"
          icon={<CalendarMonthRegular />} 
          style={{ 
            backgroundColor: this.state.buttonColor, 
            border: "1px solid",
            
          }} 
          disabled={this.state.dis} onClick={ this.boldText }>Grassetto
        </Button>
      </div>
    );
  }
}


/*

PER METTERE L'EFFETTO HOVER SUL BOTTONE
class MyButton extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      buttonColor: 'red',
      isHovered: false
    };
    this.handleMouseEnter = this.handleMouseEnter.bind(this);
    this.handleMouseLeave = this.handleMouseLeave.bind(this);
  }

  handleMouseEnter() {
    this.setState({ isHovered: true });
  }

  handleMouseLeave() {
    this.setState({ isHovered: false });
  }

  render() {
    const { isHovered } = this.state;
    const backgroundColor = isHovered ? 'gray' : this.state.buttonColor;
    
    return (
      <Button
        icon={<CalendarMonthRegular />}
        style={{ backgroundColor, border:"1px solid" }}
        disabled={this.state.dis}
        onClick={this.boldText}
        onMouseEnter={this.handleMouseEnter}
        onMouseLeave={this.handleMouseLeave}
      >
        Grassettoo
      </Button>
    );
  }
}

*/