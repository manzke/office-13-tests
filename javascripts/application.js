var application = new function(){
				var _doc;
				
				this.init = function () {
					Office.initialize = function (reason) {
						displayText("init called");
						// Store a reference to the current document.
						_doc = Office.context.document; 
						// Check whether text is already selected.
						onAfterOfficeInit();
					};
				}
				
				function onAfterOfficeInit(){
					_doc.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChange)
				}
				
				function onSelectionChange(evt){
					_doc.getSelectedDataAsync(Office.CoercionType.Text, displayText);
				}
				
				function displayText(result){
					document.getElementById("results").innerText += 'received: '+result.value;
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

function submitPres() {
    try {

        // Try to get the entire PowerPoint presentation.
        Office.context.document.getFileAsync("compressed", function(result) {

            // Check to see whether the attempt to get the file was successful.
            if (result.status == "succeeded") {

                // Get the File object from the result.
                var myFile = result.value;
                statusInfo.innerHTML = "Getting file of " + myFile.size + 
                    " bytes<br/>";

                // Iterate over each slice in the file and access each chunk.
                for (var i = 0; i < myFile.sliceCount; i++) {

                    myFile.getSliceAsync(i, function(result) {

                        // If the call returns successfully, we have access
                        // to the Slice object.
                        if (result.status == "succeeded") {

                           // Send the slice to the web service.
                           statusInfo.innerHTML += "Sending piece " + i + 
                               " of " + myFile.sliceCount + "<br/>";
                           sendSlice(result.value);
                        }
                    });
                }

                // Close the file when we’re done with it.
                myFile.closeAsync(function (result) {

                    // If the result returns as a success, the
                    // file has been successfully closed.
                    if (result.status == "succeeded") {
                        statusInfo.innerHTML += "File closed.<br/>";
                    }
                    else {
                        statusInfo.innerHTML += 
                            "File couldn't be closed.<br/>";
                    }
                });
            }
        });
    }
    catch (err) {
        statusInfo.innerHTML += (err.name, err.message);
    }
}