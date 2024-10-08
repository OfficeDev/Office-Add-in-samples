import React, { useRef } from "react";
import { insertFile } from "../office-document";
import { Button, makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  input: {
    display: "none",
  },
});

const FileUploader = () => {
  const styles = useStyles();

  const hiddenFileInput = useRef(null);

  const handleClick = () => {
    hiddenFileInput.current.click();
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      var reader = new FileReader();
      reader.onload = async function () {
        const startIndex = reader.result.toString().indexOf("base64,");
        await insertFile(reader.result.toString().substr(startIndex + 7));
      };
      reader.readAsDataURL(e.target.files[0]);
    }
  };

  return (
    <>
      <div>
        <Button appearance="primary" disabled={false} size="large" onClick={handleClick}>
          Import
        </Button>
        <label htmlFor="file" className={styles.input}>
          Choose a file:
        </label>
        <input id="file" type="file" ref={hiddenFileInput} onChange={handleFileChange} className={styles.input} />
      </div>
    </>
  );
};

export default FileUploader;
