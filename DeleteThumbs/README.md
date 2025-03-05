# DeleteThumbs
ネットワーク上のフォルダを削除しようとしても以下のようなポップアップが出て困ったため、解決するために作成した。  
![image](https://github.com/user-attachments/assets/4e1e83e2-6e8a-478b-ace8-88ce7b2e98b1)  

# そもそもなぜ消せなかったのか
■ダイアログを見ると「エクスプローラー」が「Thumbs.db」を開いているから削除できないとのこと。  
■エクスプローラーとは、言うまでもなくこいつのこと。  
![image](https://github.com/user-attachments/assets/15cc1a0c-0c2b-4a98-bf8f-545b18e9a372)  
  
■Thumbs.dbとは？  
　○フォルダの中の画像を縮小表示したとき、その縮小画像を記憶しておくために作成されるもの…らしい。  
　○要するに、サムネの表示スピードを速めるため勝手に生成されるファイルということ。  
　　※参考：https://azby.fmworld.net/usage/windows_tips/20070221/  
  
■こいつらが何かの拍子でバグる  
　→Thumbs.dbを消すためにはエクスプローラーを起動する⇔エクスプローラーを起動していたらThumbs.dbが開かれる  
 　　というデッドロック状態になるのが原因だと思われる。
