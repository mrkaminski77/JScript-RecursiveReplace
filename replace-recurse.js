/*
replace-recurse.js
By David Leyden (January 2016)

recursively search a folder for matching files and replace matching strings with replacement strings

*/

function replacer(objFile,pattern,strReplacement) {
	try {
		var objTextStream = objFile.OpenAsTextStream(1, -2);
		var strText = objTextStream.ReadAll();
		objTextStream.Close();
		If strText.search(pattern) != -1 {
			strText = strText.replace(pattern,strReplacement)
			objTextStream = objFile.OpenAsTextStream(2,-2);
			objTextStream.Write(strText);
			objTextStream.Close();
		}
		return 1;
	}
	catch (e) {
		throw(e); // error is currently thrown instead of handled...
	}
}

function recurser(objFolder,filePattern,textPattern,strReplacement) {
	var fileList = new Enumerator(objFolder.files);
	var folderList = new Enumerator(objFolder.SubFolders);
	
	for (; !fileList.atEnd(); fileList.moveNext()) {
		if (fileList.item().Name.search(filePattern) != -1) {
			replacer(fileList.item(), textPattern, strReplacement);
		}
	}
	for (; !folderList.atEnd(); folderList.moveNext()) {
		WScript.Echo(folderList.item());
		recurser(folderList.item(),filePattern,textPattern,strReplacement);
	}
}

var objFSO = new ActiveXObject("Scripting.FileSystemObject");
var strRoot = WScript.Arguments(0);
var filePattern = new RegExp(WScript.Arguments(1),'g');
var textPattern = new RegExp(WScript.Arguments(2),'g');
var strReplacement = WScript.Arguments(3);
var objRoot = objFSO.GetFolder(strRoot);

recurser(objRoot, filePattern, textPattern, strReplacement);
