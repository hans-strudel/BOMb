var path = require('path'),
	fs = require('fs');
module.paths.push(path.resolve('node_modules'));
module.paths.push(path.resolve('../node_modules'));
module.paths.push(path.resolve(__dirname, '..', '..', '..', '..', 'resources', 'app', 'node_modules'));
module.paths.push(path.resolve(__dirname, '..', '..', '..', '..', 'resources', 'app.asar', 'node_modules'));
const {app} = require('electron');  // Module to control application life.
const {BrowserWindow} = require('electron');  // Module to create native browser window.
// Keep a global reference of the window object, if you don't, the window will
// be closed automatically when the JavaScript object is garbage collected.
let mainWindow = null;
// Quit when all windows are closed.
app.on('window-all-closed', function() {
	// On OS X it is common for applications and their menu bar
	// to stay active until the user quits explicitly with Cmd + Q
	if (process.platform != 'darwin') {
		app.quit();
	}
});
// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
app.on('ready', function() {
	// Create the browser window.
	mainWindow = new BrowserWindow({
		width: 400,
		height: 400,
		x: 800, // dual monitor 2nd screen
		y: 400,
		resizable: false,
		maximizable: false,
		autoHideMenuBar: true
	});
	// and load the index.html of the app.
	mainWindow.loadURL('file://' + __dirname + '/index.html');
	// Open the DevTools.
	fs.stat('dev.txt', function(err,res){ // check if dev.txt exists
		if (res){
			mainWindow.openDevTools();
		}
	})
	// Emitted when the window is closed.
	mainWindow.on('closed', function() {
		// Dereference the window object, usually you would store windows
		// in an array if your app supports multi windows, this is the time
		// when you should delete the corresponding element.
		mainWindow = null;
	});
});
