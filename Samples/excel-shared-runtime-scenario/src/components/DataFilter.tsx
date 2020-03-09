import { Component } from 'react';
import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import { generateCustomFunction } from '../../utilities/office-apis-helpers';

interface DataFilterProps {
    selectedOption: string
}

export default class DataFilter extends Component<{}, DataFilterProps> {

    componentWillMount() {
        this.setState({
            selectedOption: 'Groceries'
        });
      }

    //Handler for when the selected option is changed
    handleOptionChange(changeEvent: any) {
        this.setState({
            selectedOption: changeEvent.target.value
        });
    }

    //Handler for when button is pressed
    handleFilteredDataButtonPress(selectedOption: string) {
        generateCustomFunction(selectedOption);
    }

    render() {
        return <div className='ms-welcome'>
            <label>
                <input type='radio' value='Communications'
                    checked={this.state.selectedOption === 'Communications'}
                    onChange={this.handleOptionChange.bind(this)} />
                Communication
      </label>
            <label>
                <input type='radio' value='Groceries'
                    checked={this.state.selectedOption === 'Groceries'}
                    onChange={this.handleOptionChange.bind(this)} />
                Groceries
      </label>

            <div className='ms-welcome__main'>
                <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.handleFilteredDataButtonPress.bind(this, this.state.selectedOption)}>Insert filtered data</Button>
            </div>
        </div>;
    }

}




