chcp 932
$input = read-host "`nThumbs.dbが入っているフォルダのパスを入力"
#write $input\Thumbs.db
if(Test-Path $input\Thumbs.db){
	cd $input
	ls -force
	write "`n　■■Thumbs.dbが見つかりました！■■　`n"
	$input = read-host "　削除しますか？`n　[y]はい(y)　[n]いいえ(n)"
	if($input -eq "y"){
		write "`nyes"
		taskkill /f /im explorer.exe
		sleep 2
		rm Thumbs.db -force
		sleep 1
		start-process explorer
		write "`n　■■Thumbs.dbを削除しました！■■　`n"
	}elseif($input -eq "n"){
		write "`nno"
	}else{
		write "`nerror"
		write "　■■yかnを入力してください！■■　`n"
	}
}else{
	write "`n`n　　ディレクトリ：$input`n`n`n`n　■■Thumbs.dbは見つかりませんでした！■■　`n"
}
read-host "`n　何かキーを押して終了"
