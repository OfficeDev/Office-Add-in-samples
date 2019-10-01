/*
    This component is needed to wrap the Fabric React MessageBar when the latter appears at the
    top of the task pane in an Office Add-in, because the taskpane in an Office Add-in has a
    personality menu that covers a small rectangle of the upper right corner of the task pane.
    This rectangle covers the "dismiss X" on the right end of the MessageBar unless extra padding
    is added.
*/

import * as React from 'react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';

export interface OfficeAddinMessageBarProps {
    onDismiss: () => void;
    message: string;
}

export default class OfficeAddinMessageBar extends React.Component<OfficeAddinMessageBarProps> {

    constructor(props: OfficeAddinMessageBarProps) {
        super(props);
        this.officeAddinTaskpaneStyle = { paddingRight: '20px' };
      }

    private officeAddinTaskpaneStyle: any;

    render() {
        return (
            <div style={this.officeAddinTaskpaneStyle}>
                <MessageBar messageBarType={MessageBarType.error} isMultiline={true} onDismiss={this.props.onDismiss} dismissButtonAriaLabel='Close'>
                {this.props.message}.{' '}</MessageBar>
            </div>

        );
    }
}
