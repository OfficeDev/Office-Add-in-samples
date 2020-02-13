import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import { generateCustomFunction } from '../../utilities/office-apis-helpers';


 export interface CustomFunctionGenerateProps {

     login: () => {};
 }

 
export interface CustomFunctionGenerateState {
    selectedOption?: string;
}


export default class ConnectButton extends React.Component<CustomFunctionGenerateProps, CustomFunctionGenerateState> {
    constructor(props, context) {
        super(props, context);

        this.boundSetState = this.setState.bind(this);

        this.state = {selectedOption: 'communication'};
    }

    boundSetState: () => {};
    
    render() {
        const { login } = this.props;

        console.log(login);
        return (
            <div className='ms-welcome'>
                  <label>
        <input type="radio" value="communication" 
                      checked={this.state.selectedOption === 'communication'}
                      onChange={this.handleOptionChange} />
        Communication
      </label>
      <label>
        <input type="radio" value="groceries" 
                      checked={this.state.selectedOption === 'groceries'}
                      onChange={this.handleOptionChange} />
        Groceries
      </label>
      
                <div className='ms-welcome__main'>
                    <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={generateCustomFunction}>Insert data function</Button>
                </div>
            </div>
        );
    }

    handleOptionChange (changeEvent: any) {
        this.setState({
          selectedOption: changeEvent.target.value
        });
      }
}
