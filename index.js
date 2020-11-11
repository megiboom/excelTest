var Excel = require('exceljs');
var workbook = new Excel.Workbook();
var testData = require('./test.json')

workbook.xlsx.readFile('template.xlsx')
    .then(function() {
        var worksheet = workbook.getWorksheet(1);

        var rowNum = 2;
        var cellNum = 3;
        var moduleLength = 4;
        var groupLength = 2;

        for(var j =0; j<groupLength; j++){
            
            cellNum=cellNum+(3*moduleLength)*j;
            /* 대분류 시작 */
            //시작행, 시작열, 끝행, 끝열
            worksheet.mergeCells(rowNum, cellNum, rowNum, cellNum+3*moduleLength-1);
            worksheet.getCell(rowNum,cellNum).alignment = {
                vertical: 'middle', 
                horizontal: 'center'
            }
            worksheet.getCell(rowNum,cellNum).fill = {
                type:'pattern',
                pattern:'solid',
                fgColor:{argb:'CCFFEE'}
            }
            worksheet.getCell(rowNum,cellNum).border = {top: {style:'medium'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};

            var row = worksheet.getRow(rowNum);        
            row.getCell(cellNum).value='154kV S/S (154kV)'  //대분류 명명
            row.getCell(cellNum).font={bold:true}
            /* 대분류 끝 */

            /* 중분류 시작 */
            for(var i=0; i<moduleLength;i++){//0 (3,3,2,5) 1 (3,6,2,8) 2 (3,9,2,11) 3>5>8>11

                worksheet.mergeCells(rowNum+1, cellNum+3*i, rowNum+1, (cellNum+3*i)+2);
                worksheet.getCell(rowNum+1,cellNum+3*i).alignment = {
                    vertical: 'middle', 
                    horizontal: 'center'
                }
                worksheet.getCell(rowNum+1,cellNum+3*i).fill = {
                    type:'pattern',
                    pattern:'solid',
                    fgColor:{argb:'CCFFEE'}
                }
                worksheet.getCell(rowNum+1,cellNum+3*i).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
            
                row = worksheet.getRow(rowNum+1);
                row.getCell(cellNum+3*i).value='(1)154kV GIS MAIN반(☞280,000)'
                row.getCell(cellNum+3*i).font={bold:true}

                /* 소분류 시작 */
                row = worksheet.getRow(rowNum+2);
                row.getCell(cellNum+3*i).value='전류[A]'
                worksheet.getCell(rowNum+2,cellNum+3*i).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
                worksheet.getCell(rowNum+2,cellNum+3*i).fill = {
                    type:'pattern',
                    pattern:'solid',
                    fgColor:{argb:'CCFFEE'}
                }
                row.getCell(cellNum+3*i+1).value='지침'
                worksheet.getCell(rowNum+2,cellNum+3*i+1).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
                worksheet.getCell(rowNum+2,cellNum+3*i+1).fill = {
                    type:'pattern',
                    pattern:'solid',
                    fgColor:{argb:'CCFFEE'}
                }
                row.getCell(cellNum+3*i+2).value='전력량'
                worksheet.getCell(rowNum+2,cellNum+3*i+2).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
                worksheet.getCell(rowNum+2,cellNum+3*i+2).fill = {
                    type:'pattern',
                    pattern:'solid',
                    fgColor:{argb:'CCFFEE'}
                }

                //합계
                row = worksheet.getRow(rowNum+28);
                var colNum = (cellNum+3*i+2 + 9).toString(36).toUpperCase();
                row.getCell(cellNum+3*i+2).value={ formula: "SUM("+colNum+5+","+colNum+29+")"}
                row.getCell(cellNum+3*i+2).numFmt = '#,##';
                row.getCell(cellNum+3*i+2).font={bold:true,color: { argb: 'FFFF0000' }}
                row.getCell(cellNum+3*i).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
                row.getCell(cellNum+3*i+1).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
                row.getCell(cellNum+3*i+2).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
                //평균
                row = worksheet.getRow(rowNum+29);
                var colNum = (cellNum+3*i+2 + 9).toString(36).toUpperCase();
                row.getCell(cellNum+3*i+2).value={ formula: "Average("+colNum+5+","+colNum+29+")"}
                row.getCell(cellNum+3*i+2).numFmt = '#,##';
                row.getCell(cellNum+3*i+2).font={bold:true,color: { argb: 'FF800080' }}
                row.getCell(cellNum+3*i).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
                row.getCell(cellNum+3*i+1).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
                row.getCell(cellNum+3*i+2).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
                //최대전력
                row = worksheet.getRow(rowNum+30);
                var colNum = (cellNum+3*i+2 + 9).toString(36).toUpperCase();
                row.getCell(cellNum+3*i+2).value={ formula: "MAX("+colNum+5+","+colNum+29+")"}
                row.getCell(cellNum+3*i+2).numFmt = '#,##';
                row.getCell(cellNum+3*i+2).font={bold:true,color: { argb: 'FF0000FF' }}
                row.getCell(cellNum+3*i).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
                row.getCell(cellNum+3*i+1).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
                row.getCell(cellNum+3*i+2).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
                // 부하량
                row = worksheet.getRow(rowNum+31);
                row.getCell(cellNum+3*i+2).font={bold:true,color: { argb: 'FFFF00FF' }}
                row.getCell(cellNum+3*i).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'medium'},right: {style:'thin'}};
                row.getCell(cellNum+3*i+1).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'medium'},right: {style:'thin'}};
                row.getCell(cellNum+3*i+2).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'medium'},right: {style:'thin'}};
                /* 소분류 끝 */


                var rows = worksheet.getRows(5,25);
            
                rows.forEach((value) => {
                    value.getCell(cellNum+3*i).value = 100000; //전류
                    value.getCell(cellNum+3*i).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
                
                    //지침
                    value.getCell(cellNum+3*i+1).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};

                    value.getCell(cellNum+3*i+2).value = 100000.3; //전력량
                    value.getCell(cellNum+3*i+2).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
                
                    value.commit();
                })
            }
            row.commit();
            /* 중분류 끝 */
            
        }

        
        return workbook.xlsx.writeFile('new.xlsx');
    })