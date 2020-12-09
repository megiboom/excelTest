var Excel = require('exceljs');
const moment = require('moment');
const fs = require('fs');
const logs = require('./logs');

var path = require('path');
var scriptName = path.basename(__filename);

//const dir = "C:/Users/Public/전력/"
const dir = "../log"

var workbook = new Excel.Workbook();

var color=['FFC6E0B4','FFFFFF99','FFFFE699','FFC6E0B4','FFCCFFCC','FFBDD7EE']

function numToSSColumn(num){
  var s = '', t;
  var n = num;
  while (num > 0) {
    t = (num - 1) % 26;
    s = String.fromCharCode(65 + t) + s;
    num = (num - t)/26 | 0;
  }
  return s || undefined;
}

async function makeExcel(data){
	var msg = "makeExcel Start"
    await logs.writeLogs(msg,scriptName);

    await workbook.xlsx.readFile('../template/template.xlsx')
    .then(function() {
	    msg = "readFile success(template.xlsx)"
        logs.writeLogs(msg,scriptName);

        var worksheet = workbook.getWorksheet(1);

        var rowNum = 3;
        var cellNum = 2;

        worksheet.getCell(2,1).value=moment().format("YYYY-MM-DD")
        data.forEach((v_group,j)=>{
            var moduleLength = 10;
            var moduleLength = v_group.moduleList.length;

            /* 대분류 시작 */
            //시작행, 시작열, 끝행, 끝열
            worksheet.mergeCells(rowNum, cellNum, rowNum, cellNum+3*moduleLength-1);
            worksheet.getCell(rowNum,cellNum).alignment = {
                vertical: 'middle', 
                horizontal: 'center'
            }

            const randomColor=color[j%color.length];
            
            worksheet.getCell(rowNum,cellNum).fill = {
                type:'pattern',
                pattern:'solid',
                fgColor:{argb:randomColor}
            }
            worksheet.getCell(rowNum,cellNum).border = {top: {style:'medium'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};

            var row = worksheet.getRow(rowNum);        
            row.getCell(cellNum).value=v_group.groupName;  //대분류 명명
            row.getCell(cellNum).font={bold:true}
            /* 대분류 끝 */

            /* 중분류 시작 */
            v_group.moduleList.forEach((v_module,i)=>{

                worksheet.mergeCells(rowNum+1, cellNum+3*i, rowNum+1, (cellNum+3*i)+2);
                worksheet.getCell(rowNum+1,cellNum+3*i).alignment = {
                    vertical: 'middle', 
                    horizontal: 'center'
                }
                worksheet.getCell(rowNum+1,cellNum+3*i).fill = {
                    type:'pattern',
                    pattern:'solid',
                    fgColor:{argb:randomColor}
                }
                worksheet.getCell(rowNum+1,cellNum+3*i).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
            
                row = worksheet.getRow(rowNum+1);
                row.getCell(cellNum+3*i).value=v_module.moduleName;
                row.getCell(cellNum+3*i).font={bold:true}

                /* 소분류 시작 */
                row = worksheet.getRow(rowNum+2);
                row.getCell(cellNum+3*i).value='전류[A]'
                worksheet.getCell(rowNum+2,cellNum+3*i).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'hair'}};
                worksheet.getCell(rowNum+2,cellNum+3*i).alignment = {
                    vertical: 'middle', 
                    horizontal: 'center'
                }
                worksheet.getCell(rowNum+2,cellNum+3*i).fill = {
                    type:'pattern',
                    pattern:'solid',
                    fgColor:{argb:randomColor}
                }

                row.getCell(cellNum+3*i+1).value='지침'
                worksheet.getCell(rowNum+2,cellNum+3*i+1).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'thin'},right: {style:'hair'}};
                worksheet.getCell(rowNum+2,cellNum+3*i+1).alignment = {
                    vertical: 'middle', 
                    horizontal: 'center'
                }
                worksheet.getCell(rowNum+2,cellNum+3*i+1).fill = {
                    type:'pattern',
                    pattern:'solid',
                    fgColor:{argb:randomColor}
                }

                row.getCell(cellNum+3*i+2).value='전력량'
                worksheet.getCell(rowNum+2,cellNum+3*i+2).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'thin'},right: {style:'thin'}};
                worksheet.getCell(rowNum+2,cellNum+3*i+2).alignment = {
                    vertical: 'middle', 
                    horizontal: 'center'
                }
                worksheet.getCell(rowNum+2,cellNum+3*i+2).fill = {
                    type:'pattern',
                    pattern:'solid',
                    fgColor:{argb:randomColor}
                }

                //합계
                row = worksheet.getRow(rowNum+28);
                var colNum = numToSSColumn(cellNum+3*i+2)
                row.getCell(cellNum+3*i+2).value={ formula: "SUM("+colNum+(rowNum+3)+":"+colNum+(rowNum+27)+")"}
                row.getCell(cellNum+3*i+2).numFmt = '#,##0.0';
                row.getCell(cellNum+3*i+2).font={bold:true,color: { argb: 'FFFF0000' }}
                row.getCell(cellNum+3*i).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'hair'}};
                row.getCell(cellNum+3*i+1).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'hair'}};
                row.getCell(cellNum+3*i+2).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'thin'}};
                //평균
                row = worksheet.getRow(rowNum+29);
                var colNum = numToSSColumn(cellNum+3*i+2)
                row.getCell(cellNum+3*i+2).value={ formula: "Average("+colNum+(rowNum+3)+":"+colNum+(rowNum+27)+")"}
                row.getCell(cellNum+3*i+2).numFmt = '#,##0.0';
                row.getCell(cellNum+3*i+2).font={bold:true,color: { argb: 'FF800080' }}
                row.getCell(cellNum+3*i).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'hair'}};
                row.getCell(cellNum+3*i+1).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'hair'}};
                row.getCell(cellNum+3*i+2).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'thin'}};
                //최대전력
                row = worksheet.getRow(rowNum+30);
                var colNum = numToSSColumn(cellNum+3*i+2)
                row.getCell(cellNum+3*i+2).value={ formula: "MAX("+colNum+(rowNum+3)+":"+colNum+(rowNum+27)+")"}
                row.getCell(cellNum+3*i+2).numFmt = '#,##0.0';
                row.getCell(cellNum+3*i+2).font={bold:true,color: { argb: 'FF0000FF' }}
                row.getCell(cellNum+3*i).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'hair'}};
                row.getCell(cellNum+3*i+1).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'hair'}};
                row.getCell(cellNum+3*i+2).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'thin'}};
                // 부하량
                row = worksheet.getRow(rowNum+31);
                row.getCell(cellNum+3*i+2).font={bold:true,color: { argb: 'FFFF00FF' }}
                row.getCell(cellNum+3*i).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'medium'},right: {style:'hair'}};
                row.getCell(cellNum+3*i+1).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'medium'},right: {style:'hair'}};
                row.getCell(cellNum+3*i+2).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'medium'},right: {style:'thin'}};
                /* 소분류 끝 */


                var rows = worksheet.getRows(rowNum+3,rowNum+22);
                rows.forEach((value,index) => {
                    if(index != 0){
                        value.getCell(cellNum+3*i).value = v_module.data[index-1].current; //전류
                        value.getCell(cellNum+3*i+2).value = v_module.data[index-1].activePowerQty-v_module.data[index-1].activePowerQtyBeg; //전력량
                    }
                    
                    value.getCell(cellNum+3*i).border = {top: {style:'hair'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'hair'}};
                    value.getCell(cellNum+3*i).numFmt='#,##0'         
                    
                    value.getCell(cellNum+3*i+1).value = v_module.data[index].activePowerQtyBeg; //지침
                    value.getCell(cellNum+3*i+1).border = {top: {style:'hair'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'hair'}};
                    value.getCell(cellNum+3*i+1).numFmt='#,##0.00'         

                    value.getCell(cellNum+3*i+2).border = {top: {style:'hair'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'thin'}};
                    value.getCell(cellNum+3*i+2).numFmt='#,##0.0'  
                    
                    value.commit();
                })
            })
            row.commit();
            /* 중분류 끝 */
            cellNum=cellNum+(3*moduleLength);
        })
        
	    msg = "makeExcel Complete(setData)"
        logs.writeLogs(msg,scriptName);

        return "complete";
    }).then(()=>{
        const fileName = '/'+moment().format("YYYYMMDD")+"일자 전력데이터"
	    try{
	    	fs.mkdirSync(dir);
	        msg = "make dir : "+dir
            logs.writeLogs(msg,scriptName);
	    }catch(e){
	    	if ( e.code != 'EEXIST' ) throw e; // 존재할경우 패스처리함.
	    }
        workbook.xlsx.writeFile(dir+fileName+'.xlsx');

        msg = "makeExcel Complete : "+dir+fileName+'.xlsx';
        logs.writeLogs(msg,scriptName);

        return dir+fileName+'.xlsx';
    }).catch((e)=>{
        msg = "makeExcel Fail : "+e;
        logs.writeLogs(msg,scriptName);
        console.error(e);
        return e;
    })
    return "complete"
}
module.exports.makeExcel = makeExcel;