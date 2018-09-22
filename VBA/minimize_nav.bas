'ツールリボンとナビゲーションウィンドウを消す
'=>AccessをUIのみの表示にしてアプリっぽさを出す

'リボン最小化とナビゲーションウィンドウを閉じる
If CommandBars.GetPressedMso("MinimizeRibbon") = False Then
	CommandBars.ExecuteMso "MinimizeRibbon"
End If
DoCmd.NavigateTo "acNavigationCategoryObjectType", ""