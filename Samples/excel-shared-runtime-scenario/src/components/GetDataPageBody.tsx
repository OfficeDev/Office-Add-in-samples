import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';

export interface GetDataPageBodyProps {
    getFileNames: () => {};
    logout: () => {};
}

export default class GetDataPageBody extends React.Component<GetDataPageBodyProps> {
    render() {
        const { getFileNames, logout } = this.props;

        return (
            <div className='ms-welcome__main'>
                <h2 className='ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20'>You can add data from Microsoft Graph to the document on this page.</h2>
                <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={getFileNames}>Get File Names</Button>
                <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={logout}>Sign out from Office 365</Button>
            </div>
        );
    }
}
