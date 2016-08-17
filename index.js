
const v = '1.0.0'
const AUTHOR = 'HansS';

var remote = require('electron').remote,
	dialog = require('electron').remote.dialog,
	fs = require('fs'),
	path = require('path'),
	xls = require('xlsjs');
	
var fnInfo,
	currFile;
var typeOfConv;
var data;
var toplevelAssy,
	toplevelRev;

var fileTypes = ['.CSV','.XLSX','.XLS'];

var inp = document.getElementById('fInputBUT');
inp.onclick = getFile;

var conv = document.getElementById('type');
conv.onchange = function(){
	console.log(conv.value)
	typeOfConv = conv.value;
	if (conv.selectedIndex < 3){
		document.getElementById('settings').className = 'disabled';
	} else {
		document.getElementById('settings').className = '';
	}
	if (currFile && (conv.selectedIndex !== 0)){
		switch (conv.value){
			case 'bel':
				loadBEL(currFile);
				break;
		}
	}
}

inp.ondragover = function () {
	this.className = 'hover';
	return false;
}
inp.ondragleave = inp.ondragend = function () {
	this.className = '';
	return false;
};
inp.ondrop = function (e) {
	this.className = '';
	e.preventDefault();
	var file = e.dataTransfer.files[0];
	currFile = file.path;
	loadFileData(file.path)
	return false;
};

function displayError(str){
	document.getElementById('error').innerHTML = str
}

window.onerror = function(m,f,l,c,e){
	displayError(e)
	if (global.process.listeners('uncaughtException').length > 0) {
		global.process.emit('uncaughtException', e)
		return true
    } else {
		return false
    }
}

function getFile(){
	dialog.showOpenDialog({}, (fn) => { //array
		if (fn){
			currFile = fn[0];
			loadFileData(fn[0]) // send the first file	
		}
	})
}

function loadFileData(fn){
	if (fn){
		fnInfo = path.parse(fn)
		document.getElementById('fInfo').innerHTML = '<span class="fName">' + fnInfo.base + '</span><br />'
		if (!(fileTypes.indexOf(fnInfo.ext.toUpperCase()) > -1)){
			displayError('Bad File Type')
		}
		switch (conv.value){
			case 'bel':
				loadBEL(currFile);
				break;
		}
	} else {
		// no filenames selected
	}
}

var fOut = document.getElementById('fOutputBUT')
fOut.onclick = function(){
	if (fnInfo || data){
		console.log(typeOfConv)
		switch (typeOfConv){
			case 'bel':
				outputBEL(data, document.getElementById('justitems').checked);
				break;
			case 'm2a':
				mysis2arena(currFile);
				break;
			case 'a2m':
				arena2mysis(currFile);
				break;
		}
	} else {
		displayError('no file loaded!')
	}
}

var cells = {}
var FIRST = true
var arenaHeaders = 'Level,Item_number,item_name,revision,Quantity,bom_notes,Unit_Of_Measure,reference_designator'

var boms = []
var itemheaders = {
	'miitem': 'itemId,descr,type,custFld1,xdesc,ref,pick,sales,uOfM,poUOfM,uConvFact,glId,segId,cycle,revId,track,lead,minLvl,maxLvl,ordLvl,ordQty,variance,lotSz,cLast,cStd,cAvg,cLand,unitWgt,locId,suplId,mfgId,lotMeth,glAcct,apDist,a0Status,a0Func,a0Vol,a0Off,a1Status,a1Func,a1Vol,a1Off,a2Status,a2Func,a2Vol,a2Off,status,bHuman,lstUseDt,lstPIDt,a0Start,a0End,a1Start,a1End,a2Start,a2End,totQStk,totQWip,totQRes,totQOrd,totIqStk,totIqWip,totIqRes,totIqOrd,itemCost',
	'miqmfg': 'itemId,mfgId,mfgName,mfgProdCode',
}
var bomheaders = {
	'mibomh': 'bomItem,bomRev,autoBuild,author,revCmnt,rollup,mult,yield,lstMainDt,descr,ecoNum,ovride,docPath,assyLead,ecoDocPath,qPerLead,revDate,effStartDate,effEndDate,revDt,effStartDt,effEndDt,totQWip,totQRes,maxLead,opCnt',
	'mibomd': 'bomItem,bomRev,bomEntry,lineNbr,partId,qty,opCode,srcLoc,custFld1,custFld2,dType,revId,lead,cmnt,altItems',
	'mibord': 'bomItem,bomRev,opCode,lineNbr,wcId,toolId,batchSize,cycleTime,setupTime,preopTime,postTime,cmnt,overlap,milestone,ctlGroupId'
}
function mysis2arena(file){
	data = {}
	var wk = xls.readFile(file)
	
	var cells = wk.Sheets[wk.SheetNames[0]]
	var rows = cells['!range']['e']['r']
	var obj = {}
	//console.log(cells["A1"],rows)
	
	var es = {v:""}
	
	for (i = 2; i <= rows; i++){
		
		obj.name = (cells["A"+i] || es).v
		obj.rev = (cells["B"+i] || es).v
		obj.desc = (cells["X"+i] || es).v
		
		if (obj.name.indexOf("-" + obj.rev + "-") > -1){
			obj.Fname = obj.name.replace("-" + obj.rev + "-", "-")
		} else {
			obj.Fname = obj.name
		}
		boms.push({"name":obj.name, "Fname":obj.Fname, "rev":obj.rev, "desc":obj.desc})
	}
	var cells = wk.Sheets[wk.SheetNames[1]]
	var rows = cells['!range']['e']['r'] 
	var output
	start = 2
	dialog.showOpenDialog({'properties':['openDirectory']}, (dir)=>{
		if (dir){
			console.log(dir)
			OUTPUTDIR = dir[0]
			boms.forEach(function(elem,ind,arr){

			output = '0,' + (elem.Fname || elem.name) + ',"' + elem.desc + '",' + elem.rev + ',1,X,each\r\n'
			
			for (start = 2;start<rows+1;start++){
				
				if (elem.name != cells["A"+start].v || elem.rev != cells["B"+start].v){
					//console.log(elem)
				} else {
					//console.log((elem.name) + ': ' + start)
					output += '1,' + (cells["F"+start] || es).v + ',"' + (cells["M"+start] || es).v.replace(/"/g, "''") 
					+ '"' + ',A,' + (cells["G"+start] || es).v + ',"' + (cells["T"+start] || es).v.replace('\r\n', '') + '",' + 
					(cells["S"+start] || es).v + ',"' + (cells["U"+start] || es).v + '"\r\n'
					//console.log(output)
				}
			}
			fs.writeFileSync(OUTPUTDIR + '\\' + (elem.name) + ' rev ' + (elem.rev) + '.csv', arenaHeaders + '\r\n' + output)
			})
		} else {
			console.log('NO DIR SELECTED')
		}
	})
}

function CSVtoArray(text) {
    var re_valid = /^\s*(?:'[^'\\]*(?:\\[\S\s][^'\\]*)*'|"[^"\\]*(?:\\[\S\s][^"\\]*)*"|[^,'"\s\\]*(?:\s+[^,'"\s\\]+)*)\s*(?:,\s*(?:'[^'\\]*(?:\\[\S\s][^'\\]*)*'|"[^"\\]*(?:\\[\S\s][^"\\]*)*"|[^,'"\s\\]*(?:\s+[^,'"\s\\]+)*)\s*)*$/;
    var re_value = /(?!\s*$)\s*(?:'([^'\\]*(?:\\[\S\s][^'\\]*)*)'|"([^"\\]*(?:\\[\S\s][^"\\]*)*)"|([^,'"\s\\]*(?:\s+[^,'"\s\\]+)*))\s*(?:,|$)/g;
    // Return NULL if input string is not well formed CSV string.
    //if (!re_valid.test(text)) return null;
    var a = [];                     // Initialize array to receive values.
    text.replace(re_value, // "Walk" the string using replace with callback.
        function(m0, m1, m2, m3) {
            // Remove backslash from \' in single quoted values.
            if      (m1 !== undefined) a.push(m1.replace(/\\'/g, "'"));
            // Remove backslash from \" in double quoted values.
            else if (m2 !== undefined) a.push(m2.replace(/\\"/g, '"'));
            else if (m3 !== undefined) a.push(m3);
            return ''; // Return empty string.
        });
    // Handle special case of empty last value.
    if (/,\s*$/.test(text)) a.push('');
    return a;
}

function arena2mysis(file){
	data = {}
	FIRST = true
	dialog.showOpenDialog({'properties':['openDirectory']}, (dir)=>{
		if (dir){
			console.log(dir)
			folder = dir[0] + '\\'
			if (path.parse(file).ext.toUpperCase() == '.XLS'){
				console.log('done')
				var wk = xls.readFile(file)
				cells = wk.Sheets[wk.SheetNames[0]]
			} else if (path.parse(file).ext.toUpperCase() == '.CSV'){
				var lines = fs.readFileSync(file, 'utf8').split('\n')
				rows = lines.length
				for (var i = 0; i < lines.length-1;i++){
					currentLine = lines[i].replace(/'/g, '')
					currentLine = CSVtoArray(currentLine.replace(/"([^,]+)"/g, function(v){
						return v.replace(/\"/g, "'")
					}))
					for (var j = 0; j < lines[0].split(',').length; j++){
						cells[String.fromCharCode(65 + j) + (i + 1)] = {'v':String(currentLine[j]).trim()}
						
					}
				}
				cells['!ref'] = 0
			}
			var headers = {}	
			var lvlCol,
				headerLen = 0
				
			Object.keys(cells).every((e,i,a)=>{
				if (headers[e[0]]){
					data[headers[e[0]]] = data[headers[e[0]]] || []
					data[headers[e[0]]].push(cells[e].v)
				} else {
					headers[e[0]] = cells[e].v
					headerLen++
				}
				OUTPUT = ''
				if (cells[e].v == 'level') lvlCol = e[0]
				if (lvlCol && (i)%headerLen == 0 && i > 2*headerLen){ // make sure it doesnt trigger on headers
					if (cells[lvlCol + ((Math.floor((i)/headerLen))+1)].v == 0){
						//console.log(lvlCol, i, ((Math.floor((i)/headerLen))+1))
						delete headers['!'] // remove the '!' field
						//console.log(lvlCol + ((Math.floor((i)/headerLen)) + 1))
						//console.log(headers, cells, data)
						buildBomImport(folder)
						buildItemImport(folder)
						FIRST = false
						data = {}
					}
				}
				return true
			})
			buildBomImport(folder)  // get last item
			buildItemImport(folder)
			data = {}
		} else {
			console.log('no dir selected')
		}
	})
}

function buildItemImport(folder){
	var i
	for (x in itemheaders){
		out = ''
		if (FIRST) out += itemheaders[x] + '\r\n'
		switch (x){
			case 'miitem':
				for (i=0;i<data['item_number'].length;i++){
					out += '"' + data['item_number'][i] + '",' +
							'"' + data['item_name'][i] + '",' +
							((data['item_number'][i].indexOf('-TK') > -1)?'2':'0') + ',' +
							'FALSE' + '\r\n'
				}
				break;
			case 'miqmfg':
				for (i=0;i<data['item_number'].length;i++){
					var zyx = "manufacturer_item_number_1"
					out += '"' + data['item_number'][i] + '",' +
							1 + ',' +
							'"' + data['manufacturer_1'][i] + '",' +
							'"' + data[zyx][i] + '",' + '\r\n'
				}
				break;
		}
		//fs.unlinkSync(ITEMOUT + x.toUpperCase() + '.csv')

		fs.appendFileSync(folder + x.toUpperCase() + '.csv', out)
	}
}

function buildBomImport(folder){
	var i
	for (x in bomheaders){
		out = ''
		if (FIRST) out += bomheaders[x] + '\r\n'
		switch (x){
			case 'mibomh':
				//console.log(keys)
				out += '"' + data['item_number'][0] + '",' + data['revision'][0] + ',2,' + AUTHOR + ',TURNKEY\r\n'
				break;
			case 'mibomd':
				var len = data['item_number'].length
				if (len == 1) len += 1
				for (i=1;i<len;i++){
					out += '"' + data['item_number'][0] + '",' +
							(data['revision'][0] || 'A') + ',' + String(i) + ',' + String(i) +
							',"' + (data['item_number'][i] || '') + '",' +
							'"' + (data['quantity'][i] || '1') + '",' +
							'KITTING, HOME,"' + (data['bom_notes'][i] || '') + '","' + (data['reference_designator'][i] || '') + '"\r\n'
				}
				break;
			case 'mibord':
				cnt = 1
				out += data['item_number'][0] + ',' + data['revision'][0] + ',KITTING,' + cnt++ + ',STOCK,,,,450,900,,,,,' + '\r\n'
				if (data['bom_notes'].indexOf('SMT') > -1){
					out += data['item_number'][0] + ',' + data['revision'][0] +',SMT PLACEMENT,' + cnt++ + ',SMT,,,,900,900,,,,,' + '\r\n' +
							data['item_number'][0] + ',' + data['revision'][0] + ',SMT INSPECTION,' + cnt++ + ',SMT-INS,,,,450,450,,,,,' + '\r\n' +
							data['item_number'][0] + ',' + data['revision'][0] + ',TOUCH UP,' + cnt++ + ',T-UP,,,,450,450,,,,,' + '\r\n'
				}
				if (data['bom_notes'].indexOf('ML') > -1 || data['bom_notes'].indexOf('M-L') > -1){
					out += data['item_number'][0] + ',' + data['revision'][0] + ',MANUAL LOADING,' + cnt++ + ',ML,,,,900,450,,,,,' + '\r\n'
					out += data['item_number'][0] + ',' + data['revision'][0] + ',WAVESOLDER,' + cnt++ + ',W-S,,,,450,210,,,,,' + '\r\n'
				}				
				if (data['bom_notes'].indexOf('T/UP') > -1){
					out += data['item_number'][0] + ',' + data['revision'][0] + ',TOUCH UP,' + cnt++ + ',T-UP,,,,450,450,,,,,' + '\r\n'
				}
				if (data['bom_notes'].indexOf('2HW') > -1){
					out += data['item_number'][0] + ',' + data['revision'][0] + ',2ND HARDWARE,' + cnt++ + ',2HW,,,,450,900,,,,,' + '\r\n'
				}	
				if (data['bom_notes'].indexOf('T/UP') > -1 || data['bom_notes'].indexOf('2HW') > -1){
					out += data['item_number'][0] + ',' + data['revision'][0] + ',FINAL QC,' + cnt++ + ',FQC,,,,210,450,,,,,' + '\r\n'
				}
				out += data['item_number'][0] + ',' + data['revision'][0] + ',FINAL QC,' + cnt++ + ',FQC,,,,210,450,,,,,' + '\r\n'
				//if (data['bom_notes'].indexOf('TEST') > -1){
					out += data['item_number'][0] + ',' + data['revision'][0] + ',PCBA TEST,' + cnt++ + ',TESTING,,,,900,450,,,,,' + '\r\n'
				//}
				if (data['bom_notes'].indexOf('3HW') > -1){
					out += data['item_number'][0] + ',' + data['revision'][0] + ',3RD HARDWARE,' + cnt++ + ',3HW,,,,900,900,,,,,' + '\r\n'
				}
				out += data['item_number'][0] + ',' + data['revision'][0] + ',XFER TO FGI,' + cnt++ + ',FGI,,,,120,120,,,,,\r\n'
				
		}
		//fs.unlinkSync(BOMOUT + x.toUpperCase() + '.csv')
		fs.appendFileSync(folder + x.toUpperCase() + '.csv', out)
	}
}


function loadBEL(file){
	var headers = ['Company','Mtl Seq','Mtl Part Num','Mtl Rev','Material Description',
				'Qty Per','Manufacturer','Mfg Part Num','Reference Designators']
	data = {};
	var wk = xls.readFile(file); // pass the file in as arg
	var cells = wk.Sheets[wk.SheetNames[0]]; // first sheet
	var rows = cells['!range']['e']['r'];
	var dataLoc = {};
	var seqRow = [],
		prevSeq = [],
		level = 1,
		manuCount = 0,
		itemCount = 0;
	
	data['level'] = [];
	data['manuCount'] = [];
	data['Reference Designators'] = [];
	var col = '';
	for (var c in cells){
		if (cells[c].v == 'GFR'){
			continue;
		}
		if (String(cells[c].v).indexOf('Indented') > -1){ // try to get assy & rev
			var revLoc = cells[c].v.indexOf('Rev: ')
			toplevelAssy = cells[c].v.substring(18, revLoc).trim();
			toplevelRev = cells[c].v.substring(revLoc+5, cells[c].v.length);
	
			document.getElementById('toplevelAssy').value = 'BEL' + toplevelAssy + '-TK';
			document.getElementById('toplevelRev').value = toplevelRev + '-A';
		}
		col = (c.match(/^[A-Z]+/) || ['undefined'])[0] || c.slice(0,-1);
		if (seqRow.indexOf(col) > -1){
			
				level = seqRow.indexOf(col);
		}
		if (cells[c].v == 'Mtl Seq'){
			level++;
			seqRow[level] = col;
		}
		var head = headers.indexOf(cells[c].v);
		if (typeof dataLoc[level] === 'undefined'){
			dataLoc[level] = [];
		}
		if (head > -1){ // its a header
		
			if (typeof dataLoc[level][col] === 'undefined'){
				dataLoc[level][col] = [];
			}
			dataLoc[level][col] = headers[head];
		} else {
			if (typeof data[dataLoc[level][col]] === 'undefined'){
				data[dataLoc[level][col]] = [];
			}
			if (typeof data[dataLoc[level][col]] === 'undefined'){
				data[dataLoc[level][col]] = [];
			}
			
			if (!nxtIsPart){
				if (dataLoc[level][col] == 'Manufacturer'){
					// log(cells[c].v)
					manuCount++;
				}
				data[dataLoc[level][col]].push(cells[c].v);
			} else {
				
				if (typeof data[headers[2]] === 'undefined'){
					data[headers[2]] = [];
				}
				data[headers[2]].push(cells[c].v);
				// log('part num', cells[c].v)
				data['level'].push(level-1);
				data['manuCount'].push(manuCount);
				// log('No Refs!')
				if (data['Reference Designators'].length < itemCount-1){
					data['Reference Designators'].push('')
				}
				manuCount = 0;
				nxtIsPart = false;
			}
			if (seqRow[level] == col){
				// log(cells[c].v)
				var nxtIsPart = true;
				itemCount++;
			}
		}
	}
	data['manuCount'].push(manuCount); // grab the last entry
	data['manuCount'].splice(0,1); // and remove the first entry
	fs.writeFileSync('test.json', JSON.stringify(data));
	delete data['undefined'];
}

function outputBEL(data, onlyItems){
	var newHeaders = 'Level,Item_number,item_name,revision,Quantity,bom_notes,reference_designator';
	var tlA = document.getElementById('toplevelAssy').value;
	var tlR = document.getElementById('toplevelRev').value;
	var tlD = document.getElementById('toplevelDesc').value || 'ASSY';
	var output = newHeaders;
	var numOfManu = 0;
	data['manuCount'].forEach(function(e){
		if (e > numOfManu) numOfManu = e;
	})
	for (var i = 0; i < numOfManu; i++){
		output += ',Manufacturer ' + (i+1) + ',Manufacturer Item Num ' + (i+1);
	}
	output += '\r\n';
	var len = data['level'].length;
	var z = 0;
	output += '0,' + tlA + ',"' + tlD + '",' + tlR + '\r\n';
	if (!onlyItems){
		for (var i = 0; i < len; i++){
			output += '"' + data['level'][i] + '",';
			output += '"BEL' + data['Mtl Part Num'][i] + '",';
			output += '"' + data['Material Description'][i].replace(/\"/g, "''") + '",';
			output += '"' + data['Mtl Rev'][i] + '-A",';
			output += '"' + data['Qty Per'][i] + '",';
			output += '"' + ' ' + '",'; // location
			output += '"' + (data['Reference Designators'][i] || '') + '",';
			for (var j = 0; j < data['manuCount'][i]; j++){
				// log(z+j)
				output += '"' + data['Manufacturer'][z+j].replace(/\"/g, "''") + '",';
				output += '"' + data['Mfg Part Num'][z+j].replace(/\"/g, "''") + '",';
			}
			z += j;
			output += '\r\n';
		}
	} else {
		output = 'item_number,n,Manufacturer,Manufacturer Item Num\r\n'
		for (var i = 0; i < len; i++){

			for (var j = 0; j < data['manuCount'][i]; j++){
			if (data['Mtl Part Num'].indexOf(data['Mtl Part Num'][i]) >= i){
				output += '"BEL' + data['Mtl Part Num'][i] + '",';
				output += j+1 + ',';
				output += '"' + data['Manufacturer'][z+j].replace(/\"/g, "''") + '",';
				output += '"' + data['Mfg Part Num'][z+j].replace(/\"/g, "''") + '"\r\n';
			}
				
			}
			z += j;
		}
	}

	dialog.showSaveDialog({
	"defaultPath": ((onlyItems)?'ITEMS-':'') + 'BEL' + toplevelAssy + '-TK Rev ' + toplevelRev,
	"filters":[{"name": "CSV", "extensions": ["csv"]}]}, (fn)=>{
		if (fn){
			fs.writeFileSync(fn, output);
		}
	})
}