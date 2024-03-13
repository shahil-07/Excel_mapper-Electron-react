import React, { useEffect, useState } from "react";
import "./App.css";

const { ipcRenderer } = window.require("electron");

function App() {
  const [selectedFolder, setSelectedFolder] = useState("");
  const [isEmpbDataChecked, setIsEmpbDataChecked] = useState(false);
  const [isDimensionChecked, setIsDimensionChecked] = useState(false);
  const [runStatus, setRunStatus] = useState(0);

  const handleBrowseClick = () => {
    ipcRenderer.send("select-folder");
    console.log("Folder selection request sent to main process.");
  };

  const handleEmpbCheckboxChange = (e) => {
    setIsEmpbDataChecked(e.target.checked);
  };

  const handleDimensionCheckboxChange = (e) => {
    setIsDimensionChecked(e.target.checked);
  };

  const handleRunPythonScript = () => {
    if (!selectedFolder) {
      window.alert("Please select a folder before running the script.");
      return;
    }

    if (!isDimensionChecked && !isEmpbDataChecked) {
      window.alert("Please select at least one checkbox.");
      return;
    }

    ipcRenderer
      .invoke("run", isDimensionChecked, isEmpbDataChecked)
      .then((scriptOutput) => {
        console.log("Python Output:", scriptOutput);
      })
      .catch((error) => {
        console.error(error);
      });
  };

  useEffect(() => {
    ipcRenderer.on(
      "selected-folder",
      (event, folderPath, folderNameToSearch) => {
        if (folderPath) {
          setSelectedFolder(folderPath);
          console.log("Selected folder:", folderPath);
        }
        if (folderNameToSearch) {
          console.log("Folder name to search:", folderNameToSearch);
        }
      }
    );

    ipcRenderer.on("selected-file", (event, fileType, filePath) => {
      if (fileType === "file1") {
        // Handle file1
        console.log("Selected file1:", filePath);
      } else if (fileType === "file2") {
        // Handle file2
        console.log("Selected file2:", filePath);
      } else if (fileType === "file3") {
        // Handle file3
        console.log("Selected file3:", filePath);
      } else if (fileType === "file4") {
        // Handle file4
        console.log("Selected file4:", filePath);
      }
    });

    ipcRenderer.on("run-status", function (evt, message) {
      console.log(message);
      if (message === 1) {
        setRunStatus(1);
        alert("Run successful");
      } else if (message === 2) {
        setRunStatus(2);
      } else if (message === -1) {
        setRunStatus(-1);
        alert("Something went wrong!");
      } else {
        setRunStatus(0);
      }
    });

    // ipcRenderer.on('script-error', (event, error) => {
    //   console.error('Script error:', error);
    //   alert(`${error}`);
    // });

    // Clean up the event listener when the component unmounts
    return () => {
      ipcRenderer.removeAllListeners("selected-folder");
      ipcRenderer.removeAllListeners("selected-file");
      // ipcRenderer.removeAllListeners('script-error');
    };
  }, []);

  const isRunning = runStatus === 2;
  return (
    <div className="App">
      <div className="header"> Excel Mapper </div>
      <div className="job-folder-input">
        <div className="upload-label">Select Job Folder :</div>
        <input
          type="text"
          value={selectedFolder.split(/[\\/]/).pop()}
          readOnly
        />
        <button onClick={handleBrowseClick}>Browse</button>
      </div>

      <div className="checkbox">
        <input
          type="checkbox"
          checked={isEmpbDataChecked}
          onChange={handleEmpbCheckboxChange}
        />
        <div className="upload-label">EMPB DATA</div>
      </div>
      <div className="checkbox">
        <input
          type="checkbox"
          checked={isDimensionChecked}
          onChange={handleDimensionCheckboxChange}
        />
        <div className="upload-label">DIMENSION</div>
      </div>
      <div className="footer-actions">
        <button
          disabled={isRunning}
          className={isRunning ? "running-action" : "run-action"}
          onClick={() => handleRunPythonScript()}
        >
          {isRunning ? "Running" : "Run"}
        </button>
      </div>
    </div>
  );
}

export default App;
