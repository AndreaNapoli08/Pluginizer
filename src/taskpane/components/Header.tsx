// licenza d'uso riservata ad Andrea Napoli e all'università si Bologna
import * as React from "react";
import { ModalInfo } from './ModalInfo';
export interface HeaderProps {
  title: string;
  logo: string;
  message: string;
  message2: string;
}

export default class Header extends React.Component<HeaderProps> {
  render() {
    const { title, logo, message, message2 } = this.props;

    return (
      <section className="ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500">
        <img width="150" height="150" src={logo} alt={title} title={title} />
        <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{message} <br/> {message2}</h1>
        <ModalInfo />
      </section>
    );
  }
}
