import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import { connectService } from '../../utilities/office-apis-helpers';


 export interface ConnectButtonProps {
 
     login: () => {};
 }

export default class ConnectButton extends React.Component<ConnectButtonProps> {
    render() {
        const { login } = this.props;

        console.log(login);
        return (
            <div className='ms-welcome'>
     
                <div className='ms-welcome__main'>
                    <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={connectService}>Connect to service</Button>
                </div>
            </div>
        );
    }
}
