(*
sendEmailToDoc rev1
システムの値をmailtoのURLにエンコードして送り
機器情報をメールにて送信してもらう（ような）場面を
想定して作成
【重要】テキストエディタをリッチテキストモードに変更します


*)

-------設定項目
---宛先toメールアドレス
set theToMailAdd to "to@foo.hoge.com" as text

---宛先ccメールアドレス
set theCcMailAdd to "cc@foo.hoge.com" as text

---宛先bccメールアドレス
set theBccMailAdd to "" as text
---メール件名
set theEmailSubject to "【機器情報】対応機器情報を送ります" as text
---ここまで設定項目



--------------------------------------
tell application "TextEdit"
	try
		quit saving no
	end try
end tell

-----リッチテキストモードに変更
do shell script "defaults write com.apple.TextEdit AutosaveDelay -int 0"
do shell script "defaults write com.apple.TextEdit RichText -int 1"
do shell script "defaults write com.apple.TextEdit PlainTextEncoding -int 4"
do shell script "defaults write com.apple.TextEdit PlainTextEncodingForWrite -int 4"

do shell script "defaults write com.apple.TextEdit CheckSpellingWhileTyping -bool FALSE"
do shell script "defaults write com.apple.TextEdit CorrectSpellingAutomatically -bool FALSE"
do shell script "defaults write com.apple.TextEdit SmartCopyPaste -bool FALSE"
do shell script "defaults write com.apple.TextEdit UseXHTMLDocType -bool FALSE"

do shell script "defaults write com.apple.TextEdit PMPrintingExpandedStateForPrint -bool TRUE"
do shell script "defaults write com.apple.TextEdit IgnoreHTML -bool TRUE"
do shell script "defaults write com.apple.TextEdit DataDetectors -bool TRUE"
do shell script "defaults write com.apple.TextEdit UseTransitionalDocType -bool TRUE"

----Safariを起動しておく
tell application "Safari"
	launch
	activate
	try
		tell document 1
			close
		end tell
	end try
end tell
---テキストエディタを起動しておく
tell application "TextEdit"
	launch
	activate
	try
		tell document 1
			close
		end tell
	end try
	quit saving no
end tell
--------------------------------------
---システムインフォから各種値を呼び出しておく
set mySystemInfo to (get system info)
---OSバージョン
set theSystemVer to (system version of (mySystemInfo)) as text
---コンピューター名
set theComputerName to (computer name of (mySystemInfo)) as text
---ホスト名
set theHostName to (host name of (mySystemInfo)) as text
---IPアドレス
set theIPv4Address to (IPv4 address of (mySystemInfo)) as text
---Macアドレス
set thePrimaryEthernetAddress to (primary Ethernet address of (mySystemInfo)) as text
---ユーザー名
set theLoginUserName to (long user name of (mySystemInfo)) as text
---アカウント名
set theLoginUserName to (theLoginUserName & "｜" & (short user name of (mySystemInfo)) as text) as text
---日付けを設定
set theDate to do shell script "date '+%Y年%m月%d日 %k時%M分%S秒'"
---機器のシリアルを取得
set theSerialNumber to do shell script "ioreg -l | awk -F\\\" ' /IOPlatformSerialNumber/ { print $4 } ' "


---本文前半部分
set theInfoData to "\r\r以下の文章をメールで送信してください\r\r-----ここから\r\r"

---本文後半部分
set theOutPutData to "\r\r-----ここまで\r"

---メール本文
set theOutPutData to theInfoData & ¬
	"調査日：　" & theDate & "\r" & ¬
	"コンピューター名：　" & theComputerName & "\r" & ¬
	"ホスト名：　" & theHostName & "\r" & ¬
	"IPアドレス：　" & theIPv4Address & "\r" & ¬
	"Macアドレス：　" & thePrimaryEthernetAddress & "\r" & ¬
	"OS ver：　" & theSystemVer & "\r" & ¬
	"シリアル：　" & theSerialNumber & "\r" & ¬
	"機器利用者：　" & theLoginUserName & "\r" & ¬
	theOutPutData & "\r"



---保存ファイル名に付加する日付
set theFileNameDate to do shell script "date '+%Y%m%d'"
---ファイル名を定義（拡張子無し）
set theFileName to (theFileNameDate & "_機器情報") as text
---書類フォルダまでのアリアス
set aliasDocumentsFolder to path to documents folder from user domain as alias
---書類フォルダまでのUNIXパス
set theUnixPath to POSIX path of (aliasDocumentsFolder) as text
---保存先までのフルパス（拡張子を付加）
set theSaveFile to (aliasDocumentsFolder & theFileName & ".rtf") as text


------結果をテキストエディタに表示する
tell application "TextEdit"
	close
	make new document at front with properties {text:theOutPutData, name:theFileName, path:theUnixPath}
	tell document 1
		try
			save in file theSaveFile
		end try
		set theSaveFile to path
	end tell
	
end tell
tell application "TextEdit" to activate

---件名をエンコードする
set theSubject to my encodeURL(theEmailSubject) as text
---本文をエンコードする
set theBody to encodeURL(theOutPutData) as text
---リンク用の文字列を定義する
set theUrl to ("mailto:" & theToMailAdd & "?subject=" & theSubject & "&cc=" & theCcMailAdd & "&bcc=" & theBccMailAdd & "&body=" & theBody & "") as text
---メールリンクをsafariで開く
tell application "Safari"
	make new document with properties {name:"メールを送信してください"}
	activate
	tell document 1
		try
			tell window 1
				open location theUrl
				close
			end tell
		end try
	end tell
end tell



---URLエンコードのサブルーチン
on encodeURL(str)
	set scpt to "php -r 'echo urlencode(\"" & str & "\");'"
	return do shell script scpt
end encodeURL


