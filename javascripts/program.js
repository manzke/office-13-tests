/// <reference path="Scripts/Office.js" />
/// <reference path="Scripts/MicrosoftAjax.js" />

var OfficeAppName;

Office.initialize = function (reason) {
// Add any initialization logic to this function.
    if (Office.context.document.customXmlParts) {
        OfficeAppName = 'Word';
    }
    else {
        OfficeAppName = 'Excel';
    }
}

var MyArray = [['Berlin'],['Munich'],['Duisburg']];

function writeData() {
    Office.context.document.setSelectedDataAsync(MyArray, { coercionType: 'matrix' });
}

function ReadData() {
    Office.context.document.getSelectedDataAsync("matrix", function (result) {
        if (result.status === "succeeded"){
            printData(result.value);
        }

        else{
            printData(result.error.name + ":" + err.message);
        }
    });
}

      function printData(data) {
    {
        var printOut = "";

        for (var x = 0 ; x < data.length; x++) {
            for (var y = 0; y < data[x].length; y++) {
                printOut += data[x][y] + ",";
            }
        }
       document.getElementById("results").innerText = printOut;
    }
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
