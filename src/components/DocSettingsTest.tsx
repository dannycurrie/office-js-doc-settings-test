import * as React from 'react';
import { Button } from 'office-ui-fabric-react';
import styled from 'styled-components';

const SettingsInfo = styled.div`
  display: flex;
  justify-content: center;
  font-size: Large;
  padding: 10px;
`;

const settingsKey = 'key';
const settingsValue = 'value';

const addToSettings = async (key, value) => {
  if (Office.context.document) {
    const { settings } = Office.context.document;
    settings.set(key, value);
    await settings.saveAsync();
    return settings.get(key);
  }
  return null;
};

const removeFromSettings = async key => {
  if (Office.context.document) {
    const { settings } = Office.context.document;
    settings.remove(key);
    await settings.saveAsync();
  }
  return null;
};

const refreshSettings = () => {
  const { settings } = Office.context.document;
  return new Promise(resolve =>
    settings.refreshAsync(({ value }) => resolve(value.get(settingsKey)))
  );
};

const DocSettingsTest: React.FC<{}> = () => {
  const [settings, setSettings] = React.useState(null);

  React.useEffect(() => {
    if (Office.context.document) {
      const { settings } = Office.context.document;
      console.log('retrieving settings');
      const value = settings.get(settingsKey);
      console.log('value: ', value);
      setSettings(value);
    }
  });

  const add = async () =>
    setSettings(await addToSettings(settingsKey, settingsValue));
  const clear = async () => setSettings(await removeFromSettings(settingsKey));
  const refresh = async () => setSettings(await refreshSettings());

  //setSettings(await refreshSettings());

  return (
    <React.Fragment>
      <SettingsInfo>{`Current settings: ${settings || 'empty'}`}</SettingsInfo>
      <Button onClick={add}>Add to Settings</Button>
      <Button onClick={clear}>Clear Settings</Button>
      <Button onClick={refresh}>Refresh Settings</Button>
    </React.Fragment>
  );
};

export default DocSettingsTest;
