//tinyMCEPopup.requireLangPack();

var GoogleChartDialog = {
	init : function() {
		var f = document.forms[0];
		// Get the selected contents as text and place it in the input
		//f.someval.value = tinyMCEPopup.editor.selection.getContent({format : 'html'});
		//f.somearg.value = tinyMCEPopup.getWindowArg('some_custom_arg');
	},

	insert : function(id) {
		// Insert the contents from the input into the document
		//tinyMCEPopup.editor.execCommand('mceInsertRawHTML', false, '<object width="700px" height="400px" data="http://sysdev/keypoint/database/graph.asp?id=' + id + '" type="text/html"></object>');
		//tinyMCEPopup.close();
		//function buttonClick() {
		var parentWin = (!window.frameElement && window.dialogArguments) || opener || parent || top;
		var parentEditor = parentWin.my_namespace_activeEditor;
		parentEditor.execCommand('mceInsertRawHTML', false, '<object width="700px" height="400px" data="http://sysdev/keypoint/database/graph.asp?id=' + id + '" type="text/html"></object>');
		tinyMCEPopup.close();
		//}
	}
};

//tinyMCEPopup.onInit.add(GoogleChartDialog.init, GoogleChartDialog);
