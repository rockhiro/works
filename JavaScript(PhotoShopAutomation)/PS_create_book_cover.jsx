//指定フォルダ内の画像ファイルに指定した画像をレイヤーとして合成してPSDにする
//--------------------------帯ファイルの指定
caution = File.openDialog("セット商品用帯をPhotoshopで開いておいてください。", "*.psd");

if (caution != null) {                           //フォルダを指定したかどうか
    #target photoshop;
    preferences.rulerUnits = Units.PIXELS;       //ピクセル単位で処理をする
    reg = /(.*)(?:\.([^.]+$))/;                  //拡張子を消すための正規表現
    fileObj = new File(caution);
    open(caution)                                //帯を開く

    //--------------------- 指定したフォルダ内のjpgファイルからpsdを書き出す -----------------
    //フォルダー指定
    foldername = Folder.selectDialog("jpgがあるフォルダを指定してください。");
    folderObj = new File(foldername);
    fsname = folderObj.fsName;

    foldername2 = Folder.selectDialog("psd出力先を指定してください。");
    folderObj2 = new File(foldername2)

    //ファイルを開く
    folderRef = new Folder(fsname);              //フォルダ内のjpg取得
    fileList = folderRef.getFiles("*.jpg");
    for (i = 0; i < fileList.length; i++) {      //フォルダ内のjpgに同じ処理を施す
        fileObj = new File(fileList[i].fsName);
        open(fileObj);                           //jpgを開く
        docObj = app.documents;                  //帯レイヤーを合成する
        activeDocument = docObj[0];
        activeDocument.layerSets["obi"].duplicate(docObj[1]);

        activeDocument = docObj[1];              //書影をPSDファイル化して保存
        saveAsPSD(folderObj2.fsName + "/" + activeDocument.name.match(reg)[1] + ".psd");

        activeDocument.close(SaveOptions.DONOTSAVECHANGES);
    }
    activeDocument.close(SaveOptions.DONOTSAVECHANGES);
    alert("処理が完了しました");

    function saveAsPSD(PSDName) {                //PSDファイルに保存するための関数
        PSDObj = new File(PSDName);
        PSDOpt = new PhotoshopSaveOptions();
        PSDOpt.alphaChannels = true;
        PSDOpt.fannotations = false
        PSDOpt.embedColorProfile = false;
        PSDOpt.layers = true;
        PSDOpt.sporColors = false;
        activeDocument.saveAs(PSDObj, PSDOpt, true, Extension.LOWERCASE);
    }
} else {
    alert("処理を中止しました");
}
