import * as React from 'react';
//import { Spinner, SpinnerType } from 'office-ui-fabric-react';
import Header from './Header';
import ConnectButton from './ConnectButton';
import Progress from './Progress';
//import StartPageBody from './StartPageBody';
//import GetDataPageBody from './GetDataPageBody';
//import SuccessPageBody from './SuccessPageBody';
import OfficeAddinMessageBar from './OfficeAddinMessageBar';
import { getGraphData } from '../../utilities/microsoft-graph-helpers';
import { writeFileNamesToWorksheet, getGlobal, ensureStateInitialized } from '../../utilities/office-apis-helpers';
import { btnSignIn } from '../commands/commands';
//import CustomFunctionGenerate from './CustomFunctionGenerate';
import DataFilter from './DataFilter';

export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
    isStartOnDocOpen: boolean;
    isSignedIn: boolean;
}

export interface AppState {
    authStatus?: string;
    fileFetch?: string;
    headerMessage?: string;
    errorMessage?: string;
}


export default class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);

        // Bind the methods that we want to pass to, and call in, a separate
        // module to this component. And rename setState to boundSetState
        // so code that passes boundSetState is more self-documenting.
        this.boundSetState = this.setState.bind(this);
        this.setToken = this.setToken.bind(this);
        this.displayError = this.displayError.bind(this);
        //this.login = this.login.bind(this);
        const theToken = localStorage.getItem('mytoken');
        console.log(btnSignIn);
        console.log('token from session storage is: ' + theToken);

        if (theToken != null) {
            // Initialize state for signed in
            console.log('signed in');
            this.state = {
                authStatus: 'loggedIn',
                fileFetch: 'notFetched',
                headerMessage: 'Welcome',
                errorMessage: ''
            };
            this.setToken(theToken);
        } else {
            // Initialize state for not signed in
            console.log('signed out');
            this.state = {
                authStatus: 'notLoggedIn',
                fileFetch: 'notFetched',
                headerMessage: 'Welcome',
                errorMessage: ''
            };
        }
    }

    /*
        Properties
    */

    // The access token is not part of state because React is all about the
    // UI and the token is not used to affect the UI in any way.
    accessToken: string;

    /*
        Methods
    */

    boundSetState: () => {};

    setToken = (accesstoken: string) => {
        console.log('setting token');
        this.accessToken = accesstoken;
    }

    displayError = (error: string) => {
        this.setState({ errorMessage: error });
    }

    // Runs when the user clicks the X to close the message bar where
    // the error appears.
    errorDismissed = () => {
        this.setState({ errorMessage: '' });

        // If the error occured during a "in process" phase (logging in or getting files),
        // the action didn't complete, so return the UI to the preceding state/view.
        this.setState((prevState) => {
            if (prevState.authStatus === 'loginInProcess') {
                return { authStatus: 'notLoggedIn' };
            }
            else if (prevState.fileFetch === 'fetchInProcess') {
                return { fileFetch: 'notFetched' };
            }
            return null;
        });
    }

    dummy1 = async () => {
        
    }

    dummy2 = async () => {
        
    }

    getFileNames = async () => {
        this.setState({ fileFetch: 'fetchInProcess' });
        getGraphData(

            // Get the `name` property of the first 3 Excel workbooks in the user's OneDrive.
            "https://graph.microsoft.com/v1.0/me/drive/root/microsoft.graph.search(q = '.xlsx')?$select=name&top=3",
            this.accessToken
        )
            .then(async (response) => {
                await writeFileNamesToWorksheet(response, this.displayError);
                this.setState({
                    fileFetch: 'fetched',
                    headerMessage: 'Success'
                });
            })
            .catch((requestError) => {
                // If this runs, then the `then` method did not run, so this error must be
                // from the Axios request in getGraphData, not the Office.js in 
                // writeFileNamesToWorksheet
                this.displayError(requestError);
            });
    }

    render() {

        const { title, isOfficeInitialized } = this.props;

        if (!isOfficeInitialized) {
            return (
                <Progress
                    title={title}
                    logo='assets/Onedrive_Charts_icon_80x80px.png'
                    message='Please sideload your add-in to see app body.'
                />
            );
        }

        // Set the body of the page based on where the user is in the workflow.
        let body;
        //let statusBody = ( <StatusBody isSignedIn={true} isStartOnDocOpen={true} />);

        const g = getGlobal() as any;
        //g.state.setTaskpaneStatus(true);
        if (g.state.isConnected) {
            //connected UI
            // filter text button
            // preview data view
            // insert cf button
            body = (<DataFilter />);
        } else {
            //disconnected UI
            //just a connect button
            body = (<ConnectButton login={this.dummy2} />);
        }
       

        return (
            <div>
                {this.state.errorMessage ?
                    (<OfficeAddinMessageBar onDismiss={this.errorDismissed} message={this.state.errorMessage + ' '} />)
                    : null}

                <div className='ms-welcome'>
                    <Header logo='assets/Onedrive_Charts_icon_80x80px.png' title={this.props.title} message={this.state.headerMessage} />
                    {body}
                </div>
              
            </div>
        );
    }

    componentDidMount() {
        ensureStateInitialized();
        let g = getGlobal() as any;

        g.state.updateRct = (data: string) => {
            // `this` refers to our react component
            this.setState({ authStatus: data });
        };
    }



}
