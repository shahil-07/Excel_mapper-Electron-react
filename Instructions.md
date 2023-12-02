python exe process

1. make sure install pyinstaller and xlwings and other 3rd party libraryies needed
2. Go to the public folder where the .py file exist and run command
   pyinstaller -F .\Excel_mapper.py
3. A build and dist folder is created.
4. use the path to the exe in the dist folder to point in the main.js with arguments

React part

1. first perform yarn build and update the location of index.html in main.js to point to new build folder of react

Electron part
yarn make

the exe will be avialble in out folder
