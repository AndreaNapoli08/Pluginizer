import * as React from "react";
import Header from "./Header";
import {Menu} from './Menu'

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export default class App extends React.Component<AppProps> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
  }

  render() {
    return (
      <div>
        <Header logo="assets/logo.png" title={this.props.title} message="Benvenuto" />
        <Menu />
      </div>
    );
  }
}
