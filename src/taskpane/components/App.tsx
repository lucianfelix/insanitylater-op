import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { createProposal } from "./client";

/* global require */

// const Office = Office || {};

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: []
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration"
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality"
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro"
        }
      ]
    });
  }

  click = async () => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Text,
      { asyncContext: "This is passed to the callback" },
      function(asyncResult: Office.AsyncResult<string>) {
        const body = asyncResult.value;

        createProposal(body).then((newBody) => {
          Office.context.mailbox.item.body.setAsync(
            newBody,
            { asyncContext: "This is passed to the callback" },
            function() {}
          );
        });
        // console.log(result.value);
        // console.log(result.asyncContext);
      }
    );
    /**
     * Insert your Outlook code here
     */
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header
          logo={require("./../../../assets/logo-filled.png")}
          title={this.props.title}
          message="Serenity Now"
        />

        <p className="ms-font-l">Discover what AI can do for you today!.</p>


        {/*Radio buttons: "PersonaExec" with options: Manager, Engineer, Big Lebowski*/}
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <input type="radio" value="Manager" name="Persona" /> Manager
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <input type="radio" value="Engineer" name="Persona" /> Engineer
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <input
                  type="radio"
                  value="Big Lebowski"
                  name="Persona"
                />Big Lebowski
              </div>
            </div>
          </div>
        </div>

        {/*Radio buttons: "Tone" with options: Professional, Bastard, Fellow Kids, Upbeat*/}
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <input type="radio" value="Manager" name="Tone" /> Professional
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <input type="radio" value="Manager" name="Tone" /> Polite
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <input type="radio" value="Engineer" name="Tone" /> Disappointed
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <input type="radio" value="Engineer" name="Tone" /> Bastard
              </div>
            </div>
          </div>
        </div>





        {/*Two buttons: Generate, Reset*/}
        {/*Generate button: calls API, returns text, sets body to text*/}
        {/*Reset button: clears body*/}
        {/*<DefaultButton*/}
        {/*  className="ms-welcome__action"*/}
        {/*  iconProps={{ iconName: "ChevronRight" }}*/}
        {/*  onClick={this.click}*/}
        {/*>*/}
        {/*  Generate*/}
        {/*</DefaultButton>*/}


        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
          Insanity Later
        </DefaultButton>
      </div>
    );
  }
}
