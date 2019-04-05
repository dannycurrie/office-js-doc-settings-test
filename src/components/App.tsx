import * as React from 'react';
import Header from './Header';
import Progress from './Progress';
import DocSettingsTest from './DocSettingsTest';

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
  }

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo='assets/logo-filled.png'
          message='Please sideload your addin to see app body.'
        />
      );
    }

    return (
      <div className='ms-welcome'>
        <Header
          logo='assets/logo-filled.png'
          title={this.props.title}
          message='Welcome'
        />
        <DocSettingsTest />
      </div>
    );
  }
}
