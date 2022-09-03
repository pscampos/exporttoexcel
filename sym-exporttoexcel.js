(function (PV) {
	"use strict";

	function symbolVis() { };
	PV.deriveVisualizationFromBase(symbolVis);

	var definition = { 
		typeName: "exporttoexcel",
		displayName: 'Export to Excel',
		iconUrl: 'Scripts/app/editor/symbols/ext/icons/exporttoexcel.png',
		visObjectType: symbolVis,
		datasourceBehavior: PV.Extensibility.Enums.DatasourceBehaviors.Single,
		getDefaultConfig: function(){ 
			return { 
				DataShape: 'Timeseries',
				Height: 150,
				Width: 150, 
				HeaderColor: "blue",
				
			} 
		},
		configOptions: function(){
			return[
				{
					title: "Format Symbol",
					mode: "format"
					
				}
			];
		}
	}
	
	symbolVis.prototype.init = function(scope, elem) { 
		
		scope.exportTableToExcel = function(){
			var downloadLink;
			var dataType = 'application/vnd.ms-excel';
			var tableSelectList = document.getElementsByTagName("table");
			
			var dateNow = new Date();
			var dateNowName = formatDateTime(dateNow);
			
			for(var i=0; i<tableSelectList.length; i=i+2){
				
				if(tableSelectList[i+1] != undefined){
					
					var tableHTML = tableSelectList[i].outerHTML.replace(/ /g, '%20');
					tableHTML = tableHTML + tableSelectList[i+1].outerHTML.replace(/ /g, '%20');
					
					var countTable = (i / 2) + 1;
					
					// Specify file name
					var filename = document.getElementById("disp-name").innerText+'_Table'+countTable+'-'+dateNowName+'.xls';
					
					// Create download link element
					downloadLink = document.createElement("a");
					
					document.body.appendChild(downloadLink);
					
					if(navigator.msSaveOrOpenBlob){
						var blob = new Blob(['\ufeff', tableHTML], {
							type: dataType
						});
						navigator.msSaveOrOpenBlob( blob, filename);
					}else{
						// Create a link to the file
						downloadLink.href = 'data:' + dataType + ', ' + tableHTML;
					
						// Setting the file name
						downloadLink.download = filename;
						
						//triggering the function
						downloadLink.click();
					}
				}
			}
		};
		
		function formatDateTime(sDate) {
			var lDate = new Date(sDate)

			var hh = lDate.getHours() < 10 ? '0' + 
				lDate.getHours() : lDate.getHours();
			var mi = lDate.getMinutes() < 10 ? '0' + 
				lDate.getMinutes() : lDate.getMinutes();
			var ss = lDate.getSeconds() < 10 ? '0' + 
				lDate.getSeconds() : lDate.getSeconds();

			var d = lDate.getDate();
			var dd = d < 10 ? '0' + d : d;
			var yyyy = lDate.getFullYear();
			var mon = eval(lDate.getMonth()+1);
			var mm = (mon<10?'0'+mon:mon);
			
			return yyyy+'_'+mm+'_'+dd+'-'+hh+'_'+mi+'_'+ss;
		}
		
	};
	
	PV.symbolCatalog.register(definition); 

})(window.PIVisualization); 
