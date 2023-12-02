import React, { useEffect, useState } from 'react';
import './App.css';

const { ipcRenderer } = window.require('electron');

const UPLOADER_MAP = {
  TEMPLATE: 'template',
  EMPB_DATA: 'empb-data',
  CELL_MAPPINGS: 'cell-mappings',
  DIMENSION_SHEET: 'dimension-sheet'
};

function App() {
  const [templatePath, setTemplatepath] = useState(null)
  const [empbDataPath, setEmpbDatapath] = useState(null)
  const [cellMappingsPath, setCellMappingspath] = useState(null)
  const [dimensionSheetPath, setDimensionSheetpath] = useState(null)
  const [runStatus, setRunStatus] = useState(0)
  const [showDimension, setShowDimension] = useState(false);

  const [templateFileList, setTemplateFileList] = useState([]);
  const [dimensionFileList, setDimensionFileList] = useState([]);
  const [empbDataFileList, setEmpbDataFileList] = useState([]);
  const [cellMappingsFileList, setCellMappingsFileList] = useState([]);
  
  
  const getDropdownValues = (subfolderName, setFileList) => {
    if (!subfolderName) {
      console.error('subfolderName is undefined or null');
      return;
    }

    // console.log('Fetching files from subfolder:', subfolderName);  

    ipcRenderer
        .invoke('fetch-files-from-subfolder', subfolderName)
        .then((files) => {
            console.log(`Fetched ${subfolderName} files:`, files);
            // Verify that 'files' contains valid file paths
           // console.log('Files array:', files);
            if (files.length > 0) {
              // Set the file list for the dropdown based on fileType
              setFileList(files);
            }
        })
        .catch((error) => {
            console.error(`Error fetching files from ${subfolderName} folder:`, error);
        });
  };

  const handleRunPythonScript = () => {
    console.log(templatePath)

    // Check if the "Dimension" option is selected
    if (showDimension) {
      if (!templatePath) {
        alert('Please select a Template before running the script.');
        return; // Exit the function if the required folder is not selected
      }
      if (!dimensionSheetPath) {
        alert('Please select a Dimension sheet before running the script.');
        return; // Exit the function if the required folder is not selected
      }
      // Run the dimension script
      ipcRenderer
        .invoke('run', JSON.stringify([templatePath, dimensionSheetPath]))
        .then((scriptOutput) => {
          console.log('Python Output:', scriptOutput);
        })
        .catch((error) => {
          console.error(error);
        });
    } else {
      // Run the default script
      if (!templatePath) {
        alert('Please select a Template before running the script.');
        return; // Exit the function if the required folder is not selected
      }
      if (!empbDataPath) {
        alert('Please select a empb Data before running the script.');
        return; // Exit the function if the required folder is not selected
      }
      if (!cellMappingsPath) {
        alert('Please select a cell Mappings before running the script.');
        return; // Exit the function if the required folder is not selected
      }

      ipcRenderer
        .invoke('run', JSON.stringify([templatePath, empbDataPath, cellMappingsPath, null]))
        .then((scriptOutput) => {
          console.log('Python Output:', scriptOutput);
        })
        .catch((error) => {
          console.error(error);
        });
     }
  };
  
  // Define a debounce function
  const debounce = (func, delay) => {
    let timeout;
    return function (...args) {
      clearTimeout(timeout);
      timeout = setTimeout(() => {
        func(...args);
      }, delay);
    };
  };
  
  const debouncedHandleUploadButton = debounce((type) => {
    ipcRenderer.invoke('save-files', type).then((result) => {
      console.log(result);
    });
  }, 500); 

  const handleOpenOutputFile = () => {
      ipcRenderer.invoke('open-file-explorer');
  };

  useEffect(() => {
    ipcRenderer.on('file-saved', function (evt, message) {
      console.log(message);
      const { path, type } = message
      if (type === UPLOADER_MAP.TEMPLATE) {
        setTemplatepath(path)
      } else if (type === UPLOADER_MAP.EMPB_DATA) {
        setEmpbDatapath(path)
      } else if (type === UPLOADER_MAP.CELL_MAPPINGS) {
        setCellMappingspath(path)
      } else if (type === UPLOADER_MAP.DIMENSION_SHEET) {
        setDimensionSheetpath(path)
      }
    });

    ipcRenderer.on('run-status', function (evt, message) {
      console.log(message);
      if(message === 1){
        setRunStatus(1)
        alert('Run successful');
      }else if(message === 2){
        setRunStatus(2)
      }else if(message === -1){
        setRunStatus(-1)
        alert('Something went wrong!');
      }else{
        setRunStatus(0)
      }
    });
    
    getDropdownValues(UPLOADER_MAP.TEMPLATE, setTemplateFileList);
    getDropdownValues(UPLOADER_MAP.EMPB_DATA, setEmpbDataFileList);
    getDropdownValues(UPLOADER_MAP.CELL_MAPPINGS, setCellMappingsFileList);
    getDropdownValues(UPLOADER_MAP.DIMENSION_SHEET, setDimensionFileList);
    
    return () => {
      ipcRenderer.removeAllListeners('file-saved');
      ipcRenderer.removeAllListeners('run-status')
    };
  }, []);
  
  const isRunning = runStatus === 2
  return (
    <div className="App">
      <div className='header'>
        Excel Mapper
        {/* <div className='open-output-area'> */}
        <span className='open-output-action' onClick={handleOpenOutputFile}>Open Output File</span>
      {/* </div> */}
      </div>
      {/* Radio button to toggle "Template" and "Dimension" */}
      <div className='toggle-dimension'>
        <input
          type="radio"
          name="upload-type"
          value="Default"
          checked={!showDimension}
          onClick={() => setShowDimension(false)}
        />
        <div className='upload-label'>Default</div>
      
        <input
          type="radio"
          name="upload-type"
          value="showDimension"
          checked={showDimension}
          onClick={() => setShowDimension(true)}
        />
        <div className='upload-label'>Dimension</div>
      </div>
      <div className='upload-area'>
        <div className='upload-sect'>
          <div className='uploader-label'>Template</div>
          <button
            className="upload-btn"
            onClick={() => debouncedHandleUploadButton(UPLOADER_MAP.TEMPLATE)}
          >
            {templatePath ? 'Uploaded' : 'Add file'}
          </button>
          <div className="divider">OR</div>
        <select
            className="dropdown-btn"
            onChange={(e) => {
              setTemplatepath(e.target.value);
              // const selectedValue = e.target.value; 
              // setTemplatepath(selectedValue); 
            }}
              value={templatePath || ''}
          >
            <option value="">Select a Template</option>
            {templateFileList.map((filePath) => (
              <option key={filePath} value={filePath} className="dropdown-option" >
                {filePath.split(/[\\/]/).pop()}
                {/* {filePath.split('\\').pop()} */}
              </option>
            ))}
          </select>
        </div>
        {/* Conditionally render "Empb Data" and "Cell Mappings" based on state */}
        {showDimension ? (
          <div className='upload-sect'>
            <div className='uploader-label'>Dimension</div>
            <button 
              className="upload-btn" 
              onClick={() => debouncedHandleUploadButton(UPLOADER_MAP.DIMENSION_SHEET)}
            >
              {dimensionSheetPath ? 'Uploaded' : 'Add file'}
            </button>
            <div className="divider">OR</div>
            <select
              className="dropdown-btn"
              onChange={(e)=> {
                setDimensionSheetpath(e.target.value);
              }}
              value={dimensionSheetPath  || ''}
            >
               <option value="">Select a Dimension</option>
               {dimensionFileList.map((filePath) => (
                   <option key={filePath} value={filePath} className="dropdown-option">
                       {filePath.split(/[\\/]/).pop()}
                   </option>
               ))}
              </select>
          </div>
        ) : (
          <>
            <div className='upload-sect'>
              <div className='uploader-label'>Empb Data</div>
              <button 
                className="upload-btn"
                onClick={() =>  debouncedHandleUploadButton(UPLOADER_MAP.EMPB_DATA)}
              >
                  {empbDataPath ? 'Uploaded' : 'Add file'}
                </button>
                <div className="divider">OR</div>
                <select
                  className="dropdown-btn"
                  onChange={(e)=> {
                    setEmpbDatapath(e.target.value); 
                  }}
                  value={empbDataPath  || ''}
                >
                  <option value="">Select a Empb data </option>
                  {empbDataFileList.map((filePath) => (
                      <option key={filePath} value={filePath} className="dropdown-option">
                          {filePath.split(/[\\/]/).pop()}
                      </option>
                  ))}
              </select>
            </div>
            <div className='upload-sect'>
              <div className='uploader-label'>Cell Mappings</div>
              <button 
                className="upload-btn" onClick={() => 
                debouncedHandleUploadButton(UPLOADER_MAP.CELL_MAPPINGS)}
              >
                {cellMappingsPath ? 'Uploaded' : 'Add file'}
              </button>
              <div className="divider">OR</div>
              <select
                className="dropdown-btn"
                onChange={(e)=> {
                  setCellMappingspath(e.target.value); 
                }}
                value={cellMappingsPath  || ''}
              >
                  <option value="">Select a Cell mapping </option>
                  {cellMappingsFileList.map((filePath) => (
                      <option key={filePath} value={filePath} className="dropdown-option">
                          {filePath.split(/[\\/]/).pop()}
                      </option>
                  ))}
              </select>
            </div>
          </>
        )}
      </div>
      
      {/* {
        runStatus === 1 && <div>Run successful</div>
      }
      {
        runStatus === -1 && <div>Something went wrong!</div>
      } */}
      <div className='footer-actions'>
        <button disabled={isRunning} className={isRunning?'running-action':'run-action'} onClick={() => handleRunPythonScript()}>{isRunning ?'Running': 'Run'}</button>
      </div>
      {/* <pre>{outputFilePath}</pre> */}
    </div>
  );
}

export default App;