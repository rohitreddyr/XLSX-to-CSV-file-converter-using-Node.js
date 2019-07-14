const xlsx = require('xlsx');
const fs= require('fs');
const jsonexport= require('jsonexport');
const csv=require('fast-csv');
const dateformat=require('dateformat');


function readfile(file){
	var wb = xlsx.readFile(file,{cellDates:true, cellNF: false, cellText:true});
	var jsonData=[];
	var data=[];
	var sheetname= wb.SheetNames;
	var sheet=wb.Sheets;
	for(i=0;i<sheetname.length;i++){
		var ws=wb.Sheets[sheetname[i]];
		var data=xlsx.utils.sheet_to_json(ws,{dateNF:"yyyy-mmm-dd"});
		for(j=0;j<data.length;j++){
			data[j].Sheet=sheetname[i];
			data[j].File=file;
			data[j].Created=dateformat(data[j].Created,'yyyy-mmm-dd')
			data[j].Updated=dateformat(data[j].Updated,'yyyy-mmm-dd')
			jsonData=jsonData.concat(data[j]);
		}
	}
	return jsonData;
}

function files(){
	var jsonData=readfile("examplexlsxfile.xlsx");
	
	
	return jsonData;
	
}

function names(data, filename){
	var wb=xlsx.readFile(filename);
	var sheetname= wb.SheetNames;
	var sheet=wb.Sheets;
	for(k=0;k<sheetname.length;k++){
		var ws=wb.Sheets[sheetname[k]]
		var name=xlsx.utils.sheet_to_json(ws);
		for(i=0;i<data.length;i++){
			for(j=0;j<name.length;j++){
				//Normalisation of data
				if(data[i]["Assigned to"].toLowerCase()==name[j]["Alias"]){
					data[i]["Assigned to"]=name[j]["Name"];
				}
				if(data[i]['Priority']==name[j]["Alias"]){
					data[i]['Priority']=name[j]["Name"];
				}
				if(data[i]['Status'].toLowerCase()==name[j]['Alias']){
					data[i]['Status']=name[j]['Name'];
				}
				
			}
		}
	}
	removeDuplicate(data);
	return data;
}

function removeByIndex(arr, index) {
	arr.splice(index,1);
}

function removeDuplicate(jsonData) {
	var arr = [];
	for (i = 0; i < jsonData.length; i++) {
		for (j = i + 1; j < jsonData.length; j++) {
			if (jsonData[i]['Issue Number'] == jsonData[j]['Issue Number']) {
				if (jsonData[i]['Status'] == jsonData[j]['Status']) {
					//repeated information in two months data
					jsonData[i]['Status'] == 'Closed' ?
						jsonData.splice(i, 1) :
						jsonData.splice(j, 1);
				}
				jsonData[i]['Status'] == 'Open' ?
					jsonData.splice(j, 1) :
					jsonData.splice(i, 1);
			}
			if (
				(jsonData[i]['Status'] == 'Open' &&
					jsonData[j]['Status'] == 'Closed') ||
				(jsonData[i]['Status'] == 'Closed' &&
					jsonData[j]['Status'] == 'Open') &&
				jsonData[i]['Updated'] > jsonData[j]['Updated']
			) {
				jsonData.splice(i, 1);
			}
		}
	}
}
var nwb=xlsx.utils.book_new()
var combdata=files();
var finaldata= names(combdata, "Names.xlsx");
//delete finaldata[137]
var ws1=fs.createWriteStream('outputfile.csv');
	csv.write(finaldata,{headers:true},{dateNF:"DD-MM-YYYY"}).pipe(ws1);
	console.log("Done!");
	
