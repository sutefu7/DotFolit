# DotFolit
DotFolit は、「VB.NET で作成された Windows フォーム アプリケーション向けのソースコード」を理解するための補助ツールです。

![DotFolit 使用例](https://raw.githubusercontent.com/sutefu7/DotFolit/master/Docs/Images/01_ClassMemberTree.png "クラスメンバーツリーペイン")

![DotFolit 使用例](https://raw.githubusercontent.com/sutefu7/DotFolit/master/Docs/Images/02_InheritsTree.png "継承関係図ペイン")

![DotFolit 使用例](https://raw.githubusercontent.com/sutefu7/DotFolit/master/Docs/Images/03_MethodFlowchart.png "メソッド内のフローチャート図ペイン")

![DotFolit 使用例](https://raw.githubusercontent.com/sutefu7/DotFolit/master/Docs/Images/04_MethodCallCanvas.png "メソッドの追跡画面")

# 機能

- キャレット位置を元にした、該当クラスのメンバー自動表示
- キャレット位置を元にした、該当クラスの継承関係図の自動表示
- キャレット位置を元にした、該当メソッド内の簡易フローチャート図の自動表示
- メソッドを横並びで表示しながら、呼び出し先の処理を追うことができる横断追跡

# ダウンロード

2018/04/20 時点のビルド分
[DotFolit_20180420.zip](https://github.com/sutefu7/DotFolit/files/1932428/DotFolit_20180420.zip "DotFolit_20180420.zip")


# 開発環境＆動作環境

本プログラムは、以下の環境で作成＆動作確認をおこなっています。

| 項目 | 値                                                               |
| ----- |:---------------------------------------------------- |
| OS   | Windows 8.1 Pro (64 bit)                              |
| IDE  | Visual Studio Community 2015                     |
| 言語 | VB.NET                                                       |
| 種類 | WPF アプリケーション (.NET Framework 4.6) |

# 注意事項

- 開発者向けの用途ですので、Visual Studio (Roslyn 含む) がインストールされていない PC での使用は想定していません。
- Visual Studio 標準コントロールでしか確認していないので、
外部ベンダー提供のコントロールを使ったアプリケーションでもうまく動作するのかは未確認です。
- Visual Studio で作成したソリューション一式＋一回以上ビルドしている状態（ビルドエラーが無い状態で、アセンブリファイルが各プロジェクト下に生成されていること）のみを解析対象としています。
読み取り専用で扱っていますが、生成されたアセンブリファイルを読み込みますので、頻繁にビルドする場合、コピーしたソリューション一式を読み込ませた方がいいかもしれません。
- ソースコードの対象言語は、VB.NET です。

# 利用ライブラリ

各ライブラリは、ライブラリ作成者様のライセンスに帰属します。
詳しくは、ライブラリ作成者様のホームページ、または GitHub をご参照ください。

- AvalonEdit

   [https://github.com/icsharpcode/AvalonEdit](https://github.com/icsharpcode/AvalonEdit)  
   MIT License  

- AvalonDock (Extended.Wpf.Toolkit)

   [https://github.com/xceedsoftware/wpftoolkit](https://github.com/xceedsoftware/wpftoolkit)  
   Microsoft Public License (Ms-PL)  
   著作権: Xceed Software Inc.  

- gong-wpf-dragdrop

   [https://github.com/punker76/gong-wpf-dragdrop](https://github.com/punker76/gong-wpf-dragdrop)  
   BSD 3-Clause  

- Livet

   [https://github.com/ugaya40/Livet](https://github.com/ugaya40/Livet)  
   zlib/libpngライセンス  

- Roslyn API 系

   MICROSOFT SOFTWARE LICENSE TERMS  
   [https://www.microsoft.com/net/dotnet_library_license.htm](https://www.microsoft.com/net/dotnet_library_license.htm)  
   以下は、NuGet に表示される代表名称であり、実際には依存するその他 dll もたくさん付いてくる（最終的には 300 MB くらい？）  
   Microsoft.Build  
   Microsoft.Build.Framework  
   Microsoft.Build.Tasks.Core  
   Microsoft.CodeAnalysis.VisualBasic  
   Microsoft.CodeAnalysis.Analyzers  
   Microsoft.CodeAnalysis.VisualBasic.Workspaces  
   もしかしたら、他にもあったかもしれない...  

- ReadJEnc

   [https://github.com/hnx8/ReadJEnc](https://github.com/hnx8/ReadJEnc)  
   MIT License  

- VS2013 Image Library

   [https://webserver2.tecgraf.puc-rio.br/iup/en/imglib/Visual%20Studio%202013%20Image%20Library%20EULA.docx](https://webserver2.tecgraf.puc-rio.br/iup/en/imglib/Visual%20Studio%202013%20Image%20Library%20EULA.docx)  
   Visual Studio 2013 Image Library EULA  

# 謝辞

無償公開しているライブラリもそうですが、他にもソースコード上に記載しているリンク先の解説サイト様方のおかげで、最低限の形まで実装することができました。
ありがとうございました。

# アプリ名の由来

以下の通りです。  
  
   Do not forget your humility  
   あなたの謙虚さを忘れないでください  
   ↓  
   混ぜたり、順序入れ替えたり  
   ↓  
   dotfolit  
   どっと　ふぉりっと  
     
   DotFolit  

