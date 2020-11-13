const moment = require('moment');
const fs = require('fs');

const dir = './log';
const fileName = '/excelLog.log';

const writeLogs=async(msg, fileNm,cd)=> {
	try{
		fs.mkdirSync(dir);
	}catch(e){
		if ( e.code != 'EEXIST' ) throw e; // 존재할경우 패스처리함.
	}
	
	if(cd==1){//END CODE
		fs.appendFileSync(dir+fileName, "\n");
	}
	try{
		fs.appendFileSync(dir+fileName, '['+moment().format("YYYY/MM/DD HH:mm:ss")+'] '+msg+" ("+fileNm+")"+"\n");
	}catch(e){
		if ( e.code != 'EEXIST' ) fs.writeFileSync( dir+fileName, '\ufeff' + '['+ moment().format("YYYY/MM/DD HH:mm:ss")+']'+msg+" ("+fileNm+")"+"\n", {encoding: 'utf8'});
	}
	if(cd==-1){//END CODE
		fs.appendFileSync(dir+fileName, "\n");
	}
}

module.exports.writeLogs = writeLogs;