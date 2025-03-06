# 作成の経緯
HTMLを触ってみたい & GitHub Pages を使ったWebサイトの公開を試してみたかったため作成した。  
#URL : https://motoyama-t.github.io/works/  
  
# GitHub Pagesで任意の階層をドキュメントルートにしたい
【引用元】https://qiita.com/debiru/items/b5b8fcfd9dabb9acc3af  
` https://motoyama-t.github.io/works/ `はデフォルトだと` works `階層の直下が指定されてる。  
今回は` work `の下の` SampleHTML `階層の中身を指定したい。  
  
①「Setting」>「Pages」>「GitHub Actions」>「Static HTML」を選択  
②`path: '.'`を`path: 'SampleHTML'`に変更  
③` .github/workflows `リポジトリが生成され、` https://motoyama-t.github.io/works/ `で問題なくサイトが表示できれば完了
