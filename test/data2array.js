var exec = require('child_process').exec,
    _ = require('lodash'),
    XLSX = require('xlsx'),
    data = require('./data.json').value,
    ReportSegment_ID = 1006,
    Wave_ID = 1,
    splits = ['SegmentLevel', 'Total'],
    wb = new Workbook()
    ;

    for(i=0; i<splits.length; i++) { 
        (function(i) {
            exec("curl --ntlm -u : localhost:20247/odata/Result(ReportSegment_ID="+ReportSegment_ID+",Wave_ID="+Wave_ID+",Split='"+splits[i]+"')", function(err, stdout, stderr) {
                if(err) throw err;
                var data = JSON.parse(stdout)
                var ws = sheet_from_array_of_arrays( convertObjectDataToTable(data.value));

                /* TEST: add worksheet to workbook */
                wb.SheetNames.push(splits[i]);
                wb.Sheets[splits[i]] = ws;

                /* write file */
                XLSX.writeFile(wb, 'Results_'+ReportSegment_ID+'.xlsx');

            });
        }(i));
    }

function convertObjectDataToTable(data) {
    data = (data.indexOf) ? data : [data];
  //  var header = _.keys(data[0])
    var table = [],
        i;
    table.push(_.keys(data[0]));
    for(i=0; i<data.length; i++) {
        table.push(_.values(data[i]));
    }
    return table;
}

/* dummy workbook constructor */
function Workbook() {
	if(!(this instanceof Workbook)) return new Workbook();
	this.SheetNames = [];
	this.Sheets = {};
}
var wb = new Workbook();

/* TODO: date1904 logic */
function datenum(v, date1904) {
	if(date1904) v+=1462;
	var epoch = Date.parse(v);
	return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

/* convert an array of arrays in JS to a CSF spreadsheet */
function sheet_from_array_of_arrays(data, opts) {
	var ws = {};
	var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
	for(var R = 0; R != data.length; ++R) {
		for(var C = 0; C != data[R].length; ++C) {
			if(range.s.r > R) range.s.r = R;
			if(range.s.c > C) range.s.c = C;
			if(range.e.r < R) range.e.r = R;
			if(range.e.c < C) range.e.c = C;
			var cell = {v: data[R][C] };
			if(cell.v == null) continue;
			var cell_ref = XLSX.utils.encode_cell({c:C,r:R});

			/* TEST: proper cell types and value handling */
			if(typeof cell.v === 'number') cell.t = 'n';
			else if(typeof cell.v === 'boolean') cell.t = 'b';
			else if(cell.v instanceof Date) {
				cell.t = 'n'; cell.z = XLSX.SSF._table[14];
				cell.v = datenum(cell.v);
			}
			else cell.t = 's';
			ws[cell_ref] = cell;
		}
	}

	/* TEST: proper range */
	if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
	return ws;
}

function Workbook() {
	if(!(this instanceof Workbook)) return new Workbook();
	this.SheetNames = [];
	this.Sheets = {};
}

