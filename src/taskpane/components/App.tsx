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
