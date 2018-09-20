:バッチからAccessVBA実行
@echo off
"C:\Program Files\Microsoft Office\Office16\MSACCESS.EXE"  "\\hogehoge/hoge.accdb" /cmd "hoge_cmd"

:バッチからAccessの最適化
@echo off
":\Program Files\Microsoft Office\Office16\MSACCESS.EXE”  “\\hogehoge/hoge.accdb" /compact



:VBA側
:'クライアント側を開いたら6時間後に強制終了 UIの設定関連を保存させないようにするため、accde形式にしておくこと
:me.open()　me.TimerInterval = 21600000

:me.Form_Timmer()
:DoCmd.Quit acQuitSaveNone
