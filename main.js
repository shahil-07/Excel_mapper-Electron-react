const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
// const { spawn} = require('child_process');
const { execFile } = require("child_process");

let mainWindow;
const isDev = false;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
  });
  mainWindow.webContents.on("did-finish-load", () => {
    mainWindow.webContents.send("app-ready"); // Notify the renderer process that the app is ready
  });

  const indexPath = isDev
    ? "http://localhost:3000"
    : `file://${path.join(__dirname, "./build/index.html")}`;

  mainWindow.loadURL(indexPath);
  mainWindow.on("closed", () => (mainWindow = null));
}

app.on("ready", createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("activate", () => {
  if (mainWindow === null) {
    createWindow();
  }
});

global.jobName = "";
let file1 = ""; //template
let file3 = ""; //cell mapping
let file2 = ""; //empb data
let file4 = ""; //balloon
let selectedFolderPath = "";

ipcMain.on("select-folder", (event) => {
  //create a file dialog to select folder
  dialog
    .showOpenDialog(mainWindow, {
      properties: ["openDirectory"],
    })
    .then((result) => {
      if (!result.canceled) {
        //send the selected folder path back to renderer process.
        selectedFolderPath = result.filePaths[0];
        console.log(`Selected folder path: ${selectedFolderPath}`);

        file2 = path.join(selectedFolderPath, "empb_data.xlsx");
        file4 = path.join(selectedFolderPath, "Balloon.xlsx");

        if (fs.existsSync(file2)) {
          event.sender.send("selected-file", "file2", file2);
          console.log("file 2:", file2);
        } else {
          console.error("empb_data Excel file not found.");
        }

        if (fs.existsSync(file4)) {
          event.sender.send("selected-file", "file4", file4);
          console.log("file 4:", file4);
        } else {
          console.error("Ballon.xlsx file not found.");
        }

        if (fs.existsSync(file2)) {
          try {
            const workbook = XLSX.readFile(file2);
            const worksheet = workbook.Sheets["ManualData"];

            let folderNameToSearch = null;
            let foundSpecification = false;

            for (let cell in worksheet) {
              const cellValue = worksheet[cell].v;
              if (foundSpecification) {
                folderNameToSearch = cellValue;
                break;
              } else if (cellValue === "Specification") {
                foundSpecification = true;
              }
            }

            if (foundSpecification) {
              event.sender.send(
                "selected-folder",
                selectedFolderPath,
                folderNameToSearch
              );
              console.log("Folder name to search:", folderNameToSearch);

              // Additional code for searching "Job Name" in MasterData sheet
              const masterDataSheet = workbook.Sheets["MasterData"];

              for (const cell in masterDataSheet) {
                const cellValue = masterDataSheet[cell].v;
                if (cellValue === "Job Name") {
                  const cellAddress = XLSX.utils.decode_range(cell).e;
                  const nextCellAddress = XLSX.utils.encode_cell({
                    r: cellAddress.r,
                    c: cellAddress.c + 1,
                  });
                  global.jobName = masterDataSheet[nextCellAddress].v;
                  break;
                }
              }

              if (global.jobName) {
                const parts = global.jobName.split("_");
                if (parts.length > 0) {
                  global.jobName = parts[0];
                  event.sender.send("job-name-found", global.jobName);
                  console.log("Job Name:", global.jobName);
                } else {
                  console.error("Job Name not found in the MasterData sheet.");
                }
              } else {
                console.error("Job Name not found in the MasterData sheet.");
              }
              // const baseDirectory = path.dirname(selectedFolderPath);
              const targetFolderName = folderNameToSearch;

              // Function to recursively search for the folder and files
              function searchForFolderAndFiles(directory) {
                const files = fs.readdirSync(directory);

                for (const file of files) {
                  const filePath = path.join(directory, file);

                  if (fs.statSync(filePath).isDirectory()) {
                    if (path.basename(filePath) === targetFolderName) {
                      console.log("Found folder:", filePath);

                      // Now, search for and filter files inside the folder
                      const folderFiles = fs.readdirSync(filePath);
                      for (const folderFile of folderFiles) {
                        for (const validPrefix1 of ["NCS"]) {
                          if (folderFile.startsWith(validPrefix1)) {
                            file1 = path.join(filePath, folderFile);
                            console.log(
                              "file1:",
                              path.join(filePath, folderFile)
                            );
                            event.sender.send(
                              "selected-file",
                              "file1",
                              path.join(filePath, folderFile)
                            );
                          }
                        }
                        for (const validPrefix3 of ["Cell_mapping"]) {
                          if (folderFile.startsWith(validPrefix3)) {
                            file3 = path.join(filePath, folderFile);
                            console.log(
                              "file3:",
                              path.join(filePath, folderFile)
                            );
                            event.sender.send(
                              "selected-file",
                              "file3",
                              path.join(filePath, folderFile)
                            );
                          }
                        }
                      }
                    } else {
                      searchForFolderAndFiles(filePath);
                    }
                  }
                }
              }

              // Get the root drive of the selected folder
              const rootDrive = path.parse(selectedFolderPath).root;
              // Specify the directory to start searching from
              const searchStartDirectory = path.join(
                rootDrive,
                "UMG EMPB",
                "Customer Specification"
              );
              searchForFolderAndFiles(searchStartDirectory);
            } else {
              console.error("Specification not found in the Excel file.");
            }
          } catch (error) {
            console.error("Error while reading Excel file:", error);
          }
        } else {
          console.error("EMPB DATA Excel file not found.");
        }
      } else {
        console.log("Folder selection dialog was canceled.");
      }
    })
    .catch((error) => {
      console.error("Error while opening folder dialog:", error);
    });
});

// run status
// 1 for DONE
// 0 for STOPPED
// 2 for running
let outputFilePath = "";

ipcMain.handle("run", async (event, isDimensionChecked, isEmpbDataChecked) => {
  mainWindow.webContents.send("run-status", 2);

  const scriptPath1 = path.join(__dirname, "./build/dist/Excel_mapper.exe");
  const scriptPath2 = path.join(
    __dirname,
    "./build/dist/Excel_mapper_dimension.exe"
  );
  // console.log('__dirname:', __dirname);

  const outputFileName = `${global.jobName}_EMPB_XXX_XX.xlsx`;
  outputFilePath = path.join(selectedFolderPath, outputFileName);
  if (isEmpbDataChecked && isDimensionChecked) {
    // Both dimension and empb data are checked, run both scripts one by one
    try {
      // Execute the first script
      // Pass the full path to the output file as an argument
      const args1 = [file1, file2, file3, selectedFolderPath, outputFilePath];
      await runScript(scriptPath1, args1);

      // Use the same output file path as an argument for the second script
      const args2 = [
        outputFilePath,
        file4,
        selectedFolderPath,
        `${global.jobName}_EMPB_XXX_XX.xlsx`,
      ];

      // Run the second script using the output file path as an argument
      await runScript(scriptPath2, args2);

      // Notify the renderer process that both scripts have completed successfully
      mainWindow.webContents.send("run-status", 1);
    } catch (error) {
      // Handle any errors that occur during script execution
      console.error("Error running scripts:", error);
      mainWindow.webContents.send("run-status", -1);
    }
  } else if (isEmpbDataChecked) {
    // Only empb data is checked, run the empb data script
    const args1 = [
      file1,
      file2,
      file3,
      selectedFolderPath,
      `${global.jobName}_EMPB_XXX_XX.xlsx`,
    ];
    runScript(scriptPath1, args1)
      .then(() => mainWindow.webContents.send("run-status", 1))
      .catch(() => mainWindow.webContents.send("run-status", -1));
  } else if (isDimensionChecked) {
    // Only dimension is checked, run the dimension script
    const args2 = [
      outputFilePath,
      file4,
      selectedFolderPath,
      `${global.jobName}_EMPB_XXX_XX.xlsx`,
    ];
    runScript(scriptPath2, args2)
      .then(() => mainWindow.webContents.send("run-status", 1))
      .catch(() => mainWindow.webContents.send("run-status", -1));
  } else {
    // Neither is checked, nothing to run
    mainWindow.webContents.send("run-status", -1);
  }
});

function runScript(scriptPath, args) {
  return new Promise((resolve, reject) => {
    const pythonProcess = execFile(scriptPath, args);

    pythonProcess.stdout.on("data", (data) => {
      console.log(data.toString());
      if (data.toString().indexOf("DONE") > -1) {
        resolve();
      }
    });

    pythonProcess.stderr.on("data", (data) => {
      console.log(`stderr: ${data}`);
      // mainWindow.webContents.send('script-error', data.toString());
      reject();
    });
  });
}
