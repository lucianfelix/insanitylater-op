import * as React from "react";
import {
  ComboBox,
  DefaultButton,
  PrimaryButton, ProgressIndicator,
  SelectableOptionMenuItemType,
  Stack,
  TextField
} from "@fluentui/react";
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

export default function App(_props: AppProps) {
  const [busy, setBusy] = React.useState<boolean>(false);
  const [original, setOriginal] = React.useState<string | undefined>();
  const [personaSelectedKey, setPersonaSelectedKey] = React.useState<string>("Software Engineer");
  const personaOptions = [
    { key: "Header1", text: "SFW", itemType: SelectableOptionMenuItemType.Header },
    { key: "CEO", text: "CEO", disabled: false },
    { key: "Vice President", text: "Vice President", disabled: false },
    { key: "Manager", text: "Manager" },
    { key: "Director", text: "Director" },
    { key: "Software Engineer", text: "Software Engineer" },
    { key: "Disgruntled worker", text: "Disgruntled worker" },
    { key: "HR Person", text: "HR Person" },
    { key: "divider", text: "-", itemType: SelectableOptionMenuItemType.Divider },
    { key: "Header2", text: "NSFW", itemType: SelectableOptionMenuItemType.Header },
    { key: "Big Lebowski", text: "Big Lebowski" }
  ];

  const [toneSelectedKey, setToneSelectedKey] = React.useState<string>("Constructive");
  const toneOptions = [
    { key: "Header1", text: "SFW", itemType: SelectableOptionMenuItemType.Header },
    { key: "Constructive", text: "Constructive" },
    { key: "Candid", text: "Candid" },
    { key: "Genuine", text: "Genuine" },
    { key: "Professional", text: "Professional" },
    { key: "Polite", text: "Polite" },
    { key: "Competent", text: "Competent (requires Premium💰)", disabled: false },
    { key: "divider", text: "-", itemType: SelectableOptionMenuItemType.Divider },
    { key: "Header2", text: "NSFW", itemType: SelectableOptionMenuItemType.Header },
    { key: "Jackass", text: "Jackass" },
    { key: "Accusing", text: "Accusing" },
    { key: "Aggressive", text: "Aggressive" },
    { key: "Disappointed", text: "Disappointed" },
    { key: "Mocking", text: "Mocking" },
    { key: "Brazen", text: "Brazen" },
    { key: "Idiotic", text: "Idiotic" },
    { key: "Humiliating", text: "Humiliating" }
  ];


  const click = async () => {
    let sourceToUse;

    try {
      setBusy(true);
      if (original) {
        sourceToUse = original;
      } else {
        sourceToUse = await getBody();
        setOriginal(sourceToUse);
      }

      const proposal = await createProposal(sourceToUse, personaSelectedKey, toneSelectedKey);
      await setBody(proposal);
    } finally {
      setBusy(false);
    }
  };

  const reset = async () => {
    // const sourceToUse = original? original : await getBody();
    // const proposal = await createProposal(sourceToUse, personaSelectedKey, toneSelectedKey);
    original && await setBody(original);
    setOriginal(undefined);
  };

  return (
    <Stack tokens={{ childrenGap: 10 }}>
      <ComboBox
        selectedKey={personaSelectedKey}
        label="Persona"
        autoComplete="on"
        allowFreeform={false}
        options={personaOptions}
        onChange={(_event, option, _index, _value) => setPersonaSelectedKey("" + option.key)}
      />

      <ComboBox
        selectedKey={toneSelectedKey}
        label="Tone"
        autoComplete="on"
        allowFreeform={false}
        options={toneOptions}
        onChange={(_event, option, _index, _value) => setToneSelectedKey("" + option.key)}
      />

      {busy && <ProgressIndicator label="Mincing words.." />}

      {!busy && <PrimaryButton onClick={click}>
        Insanity Later
      </PrimaryButton>}

      {!busy && original && <DefaultButton onClick={reset}>
        Reset
      </DefaultButton>}
    </Stack>
  );
}

function getBody(): Promise<string> {
  return new Promise((resolve, _reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Text,
      { asyncContext: "This is passed to the callback" },
      function(asyncResult: Office.AsyncResult<string>) {
        const body = asyncResult.value;
        resolve(body)
      }
    )
  });
}

function setBody(text: string): Promise<void> {
  return new Promise((resolve, _reject) => {
    Office.context.mailbox.item.body.setAsync(
      text,
      { asyncContext: "This is passed to the callback" },
      function() {
        resolve();
      }
    )
  });
}