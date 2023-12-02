const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const fs = require('fs');
const path = require('path');
const { execFile } = require('child_process');

let mainWindow;
const isDev = false

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false,
        },
    });
    mainWindow.webContents.on('did-finish-load', () => {
        mainWindow.webContents.send('app-ready'); // Notify the renderer process that the app is ready
    });

    const indexPath = isDev
        ? 'http://localhost:3000'
        : `file://${path.join(__dirname, './build/index.html')}`;

    mainWindow.loadURL(indexPath);
    mainWindow.on('closed', () => (mainWindow = null));

    create_necessary_folders();
}

app.on('ready', createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', () => {
    if (mainWindow === null) {
        createWindow();
    }
});


function create_necessary_folders() {
    const baseFolderPath = path.join(app.getPath('documents'), 'ExcelMapper');

    if (fs.existsSync(baseFolderPath)) {
        // Code to execute after the folder is there.
    } else {
        fs.mkdir(baseFolderPath, (err) => {
            if (err) {
                console.error('Error creating base folder:', err);
                return;
            }

            const uploadsFolderPath = path.join(baseFolderPath, 'uploads');

            fs.mkdir(uploadsFolderPath, (uploadsErr) => {
                if (uploadsErr) {
                    console.error('Error creating "uploads" folder:', uploadsErr);
                } else {
                    console.log('Created "uploads" folder at', uploadsFolderPath);

                    const subfolderNames = ['template', 'empb-data', 'cell-mappings', 'dimension-sheet'];

                    subfolderNames.forEach((subfolder) => {
                        const subfolderPath = path.join(uploadsFolderPath, subfolder);
                        fs.mkdir(subfolderPath, (subfolderErr) => {
                            if (subfolderErr) {
                                console.error(`Error creating ${subfolder} folder:`, subfolderErr);
                            } else {
                                console.log(`Created ${subfolder} folder at ${subfolderPath}`);
                            }
                        });
                    });
                }
            });
        });
    }
}

function fetchFilesFromSubfolder(subfolderName) {
    const subfolderPath = path.join(app.getPath('documents'), 'ExcelMapper', 'uploads', subfolderName);
    // console.log('Fetching files from:', subfolderPath);
    try {
        const files = fs.readdirSync(subfolderPath);
        return files.map((fileName) => path.join(subfolderPath, fileName));
    } catch (err) {
        console.error(`Error fetching files from ${subfolderName} folder:`, err);
        return [];
    }
}


// The function triggered by your button
function uploadImageFile(type, event) {
    console.log('Upload Type:', type);

    // opens a window to choose file
    dialog.showOpenDialog({ properties: ['openFile'] }).then(result => {

        // checks if window was closed
        if (result.canceled) {
            console.log("No file selected!")
        } else {

            // get first element in array which is path to file selected
            const filePath = result.filePaths[0];

            // get file name
            const fileName = path.basename(filePath);

            // const appPath = app.getAppPath();
            create_necessary_folders()
            const uploadPath = path.join(app.getPath("documents"), 'ExcelMapper', 'uploads', type, fileName)
            // const uploadPath = path.join(appPath, 'public', 'uploads', type, fileName)
            // const relativePath = path.relative(appPath, uploadPath)
            // console.log(uploadPath, appPath, relativePath)

            // copy file from original location to app data folder
            fs.copyFile(filePath, uploadPath, (err) => {
                if (err) throw err;
                // console.log(fileName + ' uploaded.');
                mainWindow.webContents.send('file-saved', { path: uploadPath, type });
            });
        }
    });
}

ipcMain.handle('fetch-files-from-subfolder', async (event, subfolderName) => {
    const files = fetchFilesFromSubfolder(subfolderName);
    // console.log('Fetched files:', files);
    return files;
});

ipcMain.handle('save-files', async (event, type) => {
    uploadImageFile(type, event)
});


// run status
// 1 for DONE
// 0 for STOPPED
// 2 for running 

ipcMain.handle('run', async (event, scriptPath) => {
    mainWindow.webContents.send('run-status', 2);

    const [template, empb_data, cell_mappings, dimension_sheet] = JSON.parse(scriptPath)

    // Specify the app's path
    // const appPath = app.getAppPath();

    // Build the complete path to the Python script
    let fullPathToScript;

    if(dimension_sheet !== null){
        fullPathToScript = path.join(__dirname, './build/dist/Excel_mapper_dimension.exe')
    } else {
        fullPathToScript = path.join(__dirname, './build/dist/Excel_mapper.exe')
    }    
    // const fullPathToScript = path.join(app.getPath("documents"), 'ExcelMapper', 'Excel_mapper.exe');

    const pythonProcess = execFile(fullPathToScript, [template, empb_data, cell_mappings, dimension_sheet]);

    pythonProcess.stdout.on('data', (data) => {
        console.log(data.toString())
        if(data.toString().indexOf("DONE") > -1 ){
            mainWindow.webContents.send('run-status', 1);
        }
    });

    pythonProcess.stderr.on('data', (data) => {
        console.log(`stderr: ${data}`)
        mainWindow.webContents.send('run-status', -1);
    })

    // console.log(template, empb_data, cell_mappings, dimension_sheet)

    // pythonProcess.on('close', (code) => {
    //     if (code === 0) {
    //         event.reply('python-script-done', scriptOutput);
    //     } else {
    //         event.reply('python-script-error', "Python script failed with exit code: " + code);
    //     }
    // });
});

ipcMain.handle('open-file-explorer', (event, filePath) => {
    console.log('Opening file explorer at path:', filePath);
    try {
        shell.openPath(path.join(app.getPath("documents")));
    } catch (error) {
        console.error('Error opening file explorer:', error);
    }
});
