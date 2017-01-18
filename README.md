# excel-export-ext #

一个简单导出数据为Excel文件的node.js模块

## 说明

	原excel-export模块地址：https://github.com/functionscope/Node-Excel-Export
	对excel-export模块进行了扩展，在导出数据时，不在内存中拼接所有数据，而是将数据
		写入到临时文件中，最后添加到压缩文件形成xlsx文件。

## Using excel-export-ext ##

    var express = require('express');
	var nodeExcel = require('excel-export-ext');
	var app = express();

	app.get('/Excel', function(req, res){
	  	var conf ={};
		conf.stylesXmlFile = "styles.xml";
	  	conf.cols = [{
			caption:'string',
            type:'string',
            beforeCellWrite:function(row, cellData){
				 return cellData.toUpperCase();
			},
            width:28.7109375
		},{
			caption:'date',
			type:'date',
			beforeCellWrite:function(){
				var originDate = new Date(Date.UTC(1899,11,30));
				return function(row, cellData, eOpt){
              		if (eOpt.rowNum%2){
                		eOpt.styleIndex = 1;
              		}  
              		else{
                		eOpt.styleIndex = 2;
              		}
                    if (cellData === null){
                      eOpt.cellType = 'string';
                      return 'N/A';
                    } else
                      return (cellData - originDate) / (24 * 60 * 60 * 1000);
				} 
			}()
		},{
			caption:'bool',
			type:'bool'
		},{
			caption:'number',
			 type:'number'				
	  	}];
	  	conf.rows = [
	 		['pi', new Date(Date.UTC(2013, 4, 1)), true, 3.14],
	 		["e", new Date(2012, 4, 1), false, 2.7182],
            ["M&M<>'", new Date(Date.UTC(2013, 6, 9)), false, 1.61803],
            ["null date", null, true, 1.414]  
	  	];
		
		var filePath = "test.xlsx";
		nodeExcel.execute(conf, filePath, function () {
			res.download(filePath, filePath, function () {
				fs.unlinkSync(filePath);
			});
		});
	});

	app.listen(3000);
	console.log('Listening on port 3000');

