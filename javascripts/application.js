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
