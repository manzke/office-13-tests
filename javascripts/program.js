Office.initialize = function (reason) {
// Add any initialization logic to this function.
printData("initialize: "+reason);
}

var MyArray = [['Berlin'],['Munich'],['Duisburg']];

function writeData() {
printData("write:data");
    Office.context.document.setSelectedDataAsync(MyArray, { coercionType: Office.CoercionType.Matrix });
}
//{ valueFormat: "unformatted", filterType: "all" },
function ReadData() {
Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix, 
        
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
	document.getElementById("results").innerText += data;
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

Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler(), function(result){} 
);

// Event handler function.
function myHandler(eventArgs){
printdata('Document Selection Changed');
}