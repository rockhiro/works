//指定フォルダ内の画像を加工して指定フォルダ先へリサイズして出力する
#target photoshop;
preferences.rulerUnits = Units.PIXELS;

reg = /(.*)(?:\.([^.]+$))/;
//--------------------- 指定したフォルダ内のpsdファイルからjpgを書き出す -----------------
//フォルダー指定
foldername = Folder.selectDialog("psdがあるフォルダを指定してください。");
folderObj = new File(foldername);
fsname = folderObj.fsName;

foldername2 = Folder.selectDialog("jpg出力先を指定してください。");
folderObj2 = new File(foldername2)

//ファイルを開く
folderRef = new Folder(fsname);
fileList = folderRef.getFiles("*.psd");
for (i = 0; i < fileList.length; i++) {
    fileObj = new File(fileList[i].fsName);
    open(fileObj);
    //Lサイズの作成
    actH = 250 + (250 - activeScale("W"));
    activeDocument.resizeImage(250, 250);
    activeDocument.resizeCanvas(activeScale("W"), actH, AnchorPosition.MIDDLECENTER);
    activeDocument.resizeImage(180, 180 / activeScale("W") * actH);
    activeDocument.resizeCanvas(180, 250, AnchorPosition.MIDDLECENTER);
    saveAsJPG(folderObj2.fsName + "/" + activeDocument.name.match(reg)[1] + "_l.jpg");

    //Mサイズの作成
    activeDocument.resizeImage(100, 139);
    saveAsJPG(folderObj2.fsName + "/" + activeDocument.name.match(reg)[1] + "_m.jpg");

    //Sサイズの作成
    activeDocument.resizeImage(77, 107);
    saveAsJPG(folderObj2.fsName + "/" + activeDocument.name.match(reg)[1] + "_s.jpg");

    activeDocument.close(SaveOptions.DONOTSAVECHANGES);
}
alert("リサイズが完了しました。");

function saveAsJPG(jpgName) {
    jpgObj = new File(jpgName);
    jpgOpt = new JPEGSaveOptions();
    jpgOpt.embedColorProfile = true;
    jpgOpt.formatOptions = FormatOptions.STANDARDBASELINE;
    jpgOpt.matte = MatteType.NONE;
    jpgOpt.quality = 5;
    activeDocument.saveAs(jpgObj, jpgOpt, true, Extension.LOWERCASE);
}

function activeScale(direction) {
    activeDocument.activeLayer = activeDocument.layers["obi"];
    layObj = activeDocument.activeLayer.bounds;
    x1 = parseInt(layObj[0]);
    y1 = parseInt(layObj[1]);
    x2 = parseInt(layObj[2]);
    y2 = parseInt(layObj[3]);
    W = x2 - x1;
    H = y2 - y1;
    switch (direction) {
        case "W":
            return W;
            break;
        case "H":
            return H;
            break;
        default:
            alert("方向をWHで決めてください");
            break;
    }
}
