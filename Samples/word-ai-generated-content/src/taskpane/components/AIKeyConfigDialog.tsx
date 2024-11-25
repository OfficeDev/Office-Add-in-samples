import { Input, Modal } from "antd";
import React, { useState } from "react";

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

export default function AIKeyConfigDialog(props: ApiKeyConfigProps) {
  const [inputKey, setInputKey] = useState(props.apiKey);
  const [inputEndpoint, setInputEndpoint] = useState(props.endpoint);
  const [inputDeployment, setInputDeployment] = useState(props.deployment);

  const handleOk = () => {
    if (inputKey.length > 0) {
      props.setKey(inputKey);
    }
    if (inputEndpoint.length > 0) {
      props.setEndpoint(inputEndpoint);
    }
    if (inputDeployment.length > 0) {
      props.setDeployment(inputDeployment);
    }
    props.setOpen(false);
  };

  const handleCancel = () => {
    props.setOpen(false);
  };

  return (
    <>
      <Modal title="Connect to Azure OpenAI Service" open={props.isOpen} onOk={handleOk} onCancel={handleCancel}>
        <label>Endpoint</label>
        <Input value={inputEndpoint} onChange={(e) => setInputEndpoint(e.target.value)} />
        <label>Deployment</label>
        <Input value={inputDeployment} onChange={(e) => setInputDeployment(e.target.value)} />
        <label>API Key</label>
        <Input value={inputKey} onChange={(e) => setInputKey(e.target.value)} />
      </Modal>
    </>
  );
}
