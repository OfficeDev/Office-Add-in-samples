import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import { generateCustomFunction } from '../../utilities/office-apis-helpers';


 export interface CustomFunctionGenerateProps {

     login: () => {};
 }


export interface CustomFunctionGenerateState {
    selectedOption?: string;
}


export default class CustomFunctionGenerate extends React.Component<CustomFunctionGenerateProps, CustomFunctionGenerateState> {
    constructor(props, context) {
        super(props, context);

  //      this.boundSetState = this.setState.bind(this);

        this.state = {selectedOption: 'communications'};
    }

//    boundSetState: (newState: any) => {};

    render() {
        const { login } = this.props;

        console.log(login);
        return (
            <div className='ms-welcome'>
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
                    <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.handleCustomFunctionPress.bind(this, this.state.selectedOption)}>Insert data function</Button>
                </div>
            </div>
        );
    }

    handleOptionChange (changeEvent: any) {
        this.setState({
          selectedOption: changeEvent.target.value
        });
      }

      handleCustomFunctionPress (selectedOption: string) {
          generateCustomFunction(selectedOption);
      }
}
