import { Input, Modal } from "antd";
import React from "react";

export interface ApiKeyConfigProps {
    isOpen: boolean;
    apiKey: string;
    endpoint: string;
    deployment: string;
    setKey: (key: string) => void;
    setEndpoint: (endpoint: string) => void;
    setDeployment: (deployment: string) => void;
    setOpen: (isOpen: boolean) => void;
}

export interface ApiKeyConfigState {
    inputKey: string;
    inputEndpoint: string;
    inputDeployment: string;
}

export default class AIKeyConfigDialog extends React.Component<ApiKeyConfigProps, ApiKeyConfigState> {
    constructor(props, context) {
        super(props, context);
    }

    handleOk = () => {
        if (this.state != null && (this.state.inputKey != null && this.state.inputKey.length > 0)) {
            this.props.setKey(this.state.inputKey);
        }
        if (this.state != null && (this.state.inputEndpoint != null && this.state.inputEndpoint.length > 0)) {
            this.props.setEndpoint(this.state.inputEndpoint);
        }
        if (this.state != null && (this.state.inputDeployment != null && this.state.inputDeployment.length > 0)) {
            this.props.setDeployment(this.state.inputDeployment);
        }
        this.props.setOpen(false);
    };

    handleCancel = () => {
        this.props.setOpen(false);
    };

    inputApiKeyChange = (e) => {
        this.setState({ inputKey: e.target.value });
    }

    inputEndpointChange = (e) => {
        this.setState({ inputEndpoint: e.target.value });
    }

    inputDeploymentChange = (e) => {
        this.setState({ inputDeployment: e.target.value });
    }

    render() {
        return <>
            <Modal
                title="Connect to Azure OpenAI Service"
                open={this.props.isOpen}
                onOk={this.handleOk}
                onCancel={this.handleCancel}>
                <label>Endpoint</label><Input defaultValue={this.props.apiKey} onChange={this.inputEndpointChange} />
                <label>Deployment</label><Input defaultValue={this.props.deployment} onChange={this.inputDeploymentChange} />
                <label>API Key</label><Input defaultValue={this.props.apiKey} onChange={this.inputApiKeyChange} />
            </Modal>
        </>;
    }
}