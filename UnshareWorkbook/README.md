# 作成の経緯  
「ブックの共有」機能を使わずに進捗を共有できるExcelが欲しくて作成した。  
Excelで「ブックの共有」をして進捗管理表を使っていると、以下のようなシチュエーションでいちいち①「ブックの共有」を解除 → ②修正や変更 → ③再度「ブックの共有」をする必要があった。  
  
■条件付き書式を使いたい  
![image](https://github.com/user-attachments/assets/a8bc8c48-1f87-4311-ae4e-a91dbb110ff3)  
■データの入力規則を使いたい  
![image](https://github.com/user-attachments/assets/9758ebbe-579b-4f19-a2bd-2094d1f1faae)  
■マクロを編集したい  
![image](https://github.com/user-attachments/assets/f81ae1e6-200a-4283-83c5-93ac442d906e)
  
# 使い方  
**1.ネットワーク上の共有フォルダ内に「shareに置くデータベース.xlsx」を用意する**  
**2.進捗管理を使う人全員にそれぞれ「localに置く進捗管理.xlsm」を用意する**  
![image](https://github.com/user-attachments/assets/48d6d1f2-c5b5-4bca-932b-31869c6f6395)  
  
**3.「localに置く進捗管理.xlsm」を開き、「入力マニュアル」シートのB2セルにパスを入力する**  
![image](https://github.com/user-attachments/assets/4c5a1364-eb89-4c4a-bfdf-44b05f1d338d)
  
**4.試しに「localに置く進捗管理.xlsm」の「作業進捗」シートを編集し、上書き保存する**  
![image](https://github.com/user-attachments/assets/261cf943-6063-49ba-92d7-948081c0008c)  
  
**5.「shareに置くデータベース.xlsx」に内容が反映されていれば準備完了**  
![image](https://github.com/user-attachments/assets/2a99e6d5-f69b-4e16-b122-a31717ccf44b)
