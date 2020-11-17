/*
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

/// <reference types="office-js" />
/* global Excel */ //Required for these to be found.  see: https://github.com/OfficeDev/office-js-docs-pr/issues/691

import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import Header from './Header';
import HeroList from './HeroList';
import Progress from './Progress';

const logo = require('../assets/logo-filled.png');

export default class App extends React.Component {
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
          icon: 'Ribbon',
          primaryText: 'Achieve more with Office integration',
        },
        {
          icon: 'Unlock',
          primaryText: 'Unlock features and functionality',
        },
        {
          icon: 'Design',
          primaryText: 'Create and visualize like a pro',
        },
      ],
    });
  }

  click = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load('address');

        // Update the fill color
        range.format.fill.color = 'yellow';

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={logo}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo={logo} title={this.props.title} message="Welcome JavaScript" />
        <HeroList
          message="Discover what Office .NET Core 3.1 Add-ins can do for you today!"
          items={this.state.listItems}
        >
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: 'ChevronRight' }}
            onClick={this.click}
          >
            Run
          </Button>
        </HeroList>
      </div>
    );
  }
}
