var OfficeAppName;

Office.initialize = function (reason) {
// Add any initialization logic to this function.
}

var MyArray = [['Berlin'],['Munich'],['Duisburg']];

function writeData() {
printData("write:data");
    Office.context.document.setSelectedDataAsync(MyArray, { coercionType: Office.CoercionType.Matrix });
}

function ReadData() {
Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix, 
        { valueFormat: "unformatted", filterType: "all" },
        function (asyncResult) {
			printData('returned from hell');
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                printData(error.name + ": " + error.message);
            } 
            else {
                // Get selected data.
                var dataValue = asyncResult.value; 
                printData('Selected data is ' + dataValue);
            }            
        });

//    Office.context.document.getSelectedDataAsync("matrix", function (result) {
//        if (result.status === "succeeded"){
//            printData(result.value);
//        } else{
//            printData(result.error.name + ":" + err.message);
//        }
//    });
}

function printData(data) {
	var printOut = "";

	for (var x = 0 ; x < data.length; x++) {
		for (var y = 0; y < data[x].length; y++) {
			printOut += data[x][y] + ",";
		}
	}
	document.getElementById("results").innerText += printOut;
}


function getFileContentTheNewWay1(){
    var fileContent;
    Office.context.document.getFileAsync ("compressed", function (result) {
        var myFile = result.value;
        myFile.getSliceAsync(0, function (result) {
            if (result.status == "succeeded")
                fileContent = OSF.OUtil.encodeBase64(result.value.data);
        });
    });
}

function getFileContentTheNewWay2(){
    var fileContent;
    Office.context.document.getFileAsync ("text", function (result) {
        var myFile = result.value;
        myFile.getSliceAsync(0, function (result) {
            if (result.status == "succeeded")
                fileContent = result.value.data;
        });
    });
}
