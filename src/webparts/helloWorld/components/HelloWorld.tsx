import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface IHelloWordStates {
  items: any[],
  nome: string,
  idade: string,
}

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWordStates, {}> {

  constructor(props) {
    super(props);

    this.state = {
      items: [],
      nome: "Anderson",
      idade: ""
    }
  }



  async componentDidMount() {
    const newItems: any[] = await sp.web.lists.getById(this.props.lists).items.filter(`Title eq ${this.props.Title}`).get();
    this.setState({ items: newItems });
    console.log(newItems);
  }

  public render(): React.ReactElement<IHelloWorldProps> {

    return (
      <div>
        <table className="table">
          <thead>
            <tr>
              <th scope="col">ID</th>
              <th scope="col">Title</th>
              <th scope="col">Created</th>
              <th scope="col">Actions</th>
            </tr>
          </thead>
          <tbody>
            {
              this.state.items.map((item) => {
                return(
                  <tr>
                    <th scope="row">{item.ID}</th>
                    <td>{item.Title}</td>
                    <td>{item.Created}</td>
                    <td>
                      <button className="btn btn-warning btn-sm">Editar</button>
                      <button className="btn btn-danger btn-sm">Delete</button>
                    </td>
                  </tr>
                )
              })
            }

          </tbody>
        </table>
      </div>
    );
  }
}
