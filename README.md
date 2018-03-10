# DotFolit
「Windows フォームアプリケーション」のメンテナンスを担当される方へ

![DotFolit 使用例](https://raw.githubusercontent.com/sutefu7/DotFolit/master/Docs/Images/01_InheritsTree.png "DotFolit 使用例")

このアプリ名は DotFolit（ドット　フォリット）と言います。「Windows フォームアプリケーション」のソースコードを理解する必要がある方へ、その手助けをさせていただくための補助ツールです。
使い方は簡単で、sln ファイルをドラッグアンドドロップで放り込むだけです。ちょっとお試しで使ってみませんか？

# はじめに

ソースコードは、テキストファイルです。文字列だらけで構成されています。クラスの定義も、実行処理の命令も、全部含めてプログラミングしていきます。
これらの文字列を読んだりデバッグしたりして、頭の中で相関関係図として理解していくことと思います。

とあるプログラムについて、よく熟知しているベテランプログラマーの頭の中では、各クラスの関係や処理内容など、ソースコード全体について、すでに動作を理解済みのため、
質問されてもすぐに回答できるし、バグ修正や機能強化などもそれなりの工数でやってのけてしまいます。

ところが、初めてプログラムを渡されたプログラマーの頭の中では、これらの地図（知識や経験など）がまだ無いため、最初は手間取ってしまうことは当たり前のことです。
それなのに、質問の回答をせかされたり、分かっている前提でスケジュールを組まされたりで、スタートが、一番精神的に大変な気がしています。

プログラムを最短で理解するにはどうすればいいか、どうすればベテランプログラマーの頭の中と同等になれるのか、
考えた結果、「頑張って自力で理解する」という昔の時代の考え方（苦労することが美徳）では、今の時代では通用しないと感じました。
それではどうすればいいのか、プログラマーなんだからそういうプログラムを作って「プログラムに教えてもらう」という、これからの時代の解決方法をゴールとして、このプログラムは生まれました。

# 機能紹介

## クラスメンバーの表示

![DotFolit 使用例](https://raw.githubusercontent.com/sutefu7/DotFolit/master/Docs/Images/01_InheritsTree.png "DotFolit 使用例")

ソースコードには、１．クラス構造（設計図の作成）と、２．実行処理（実行時の動作命令）がごっちゃで記載されていますが、この２つは分けて理解するべきだと考えました。
まずは、クラスメンバー｛そのクラスとフィールド、プロパティ、メソッド（名のみ、処理内容は無視）｝を知る事、見慣れた後で、次に実行処理（各メソッドの処理内容）を知ることだと考えています。

そこで、常にクラスメンバーが見れるように、ソースコードの隣にツリー表示する機能を設けています。
ツリー形式による表示方法は、デバッグ時に見れる、ローカルウィンドウやウォッチウィンドウと同じ形式ですので、相互のオブジェクトの関係が、視覚的に分かりやすく理解できるかと思います。
また、ソースコードは、コントロールキー＋マウスホイールで、拡大・縮小表示を切り替えられます。

## 継承関係図

クラスは、定義中のクラスの他に、継承元クラスと連携して作成することが多いです。
とあるメンバーがそのクラス内に定義されていなくても、継承元クラス側で定義している場合があり、始めて見る方にとっては混乱の元ではないかと思います。

そこで、継承関係をイメージ図として表示する機能を付けました。オレンジ色が継承元クラス、青色が現在見ているクラスになります。
それぞれ、矢印線でつながれており、継承の流れや各メンバーが分かれて表示されるので、家系図を横断して見ることで、理解につながるのではと考えています。
ソースコードと同じく、コントロールキー＋マウスホイールで、拡大・縮小表示を切り替えられます。

## メソッドのフローチャート

![DotFolit 使用例](https://raw.githubusercontent.com/sutefu7/DotFolit/master/Docs/Images/02_MethodFlowchart.png "DotFolit 使用例")

フローチャートと言っても簡易的なものです。どちらかというとアウトラインに近いかもしれませんが、ループ命令や条件分岐命令のみを抽出・表示することで、処理内容の複雑度を表現しています。
継承関係図と同じく、コントロールキー＋マウスホイールで、拡大・縮小表示を切り替えられます。

## メソッドの横断追跡

![DotFolit 使用例](https://raw.githubusercontent.com/sutefu7/DotFolit/master/Docs/Images/03_MethodCallCanvas.png "DotFolit 使用例")

ソリューションエクスプローラーのソースノードを右クリック→「メソッドの追跡 ... 」より、呼び出しメソッドの追跡をすることができます。
これにより、現在表示中のメソッドをそのまま表示しつつ、横隣に、呼び出し先メソッドの処理内容を表示することができます。

上記の説明画像は、説明の便宜上、メソッドを範囲として表示していますが、実際にはソースコード単位で表示されますので、対象メソッド以外のメンバーも見ることができます。
この機能に関しては、まだ実験段階で、そこそこの頻度で例外エラーが発生することを確認していますので、機能的にはまだベータ版扱いです。

分かる方にはすぐにばれたと思いますが、Debugger Canvas を意識して作成しました（これをパクリとも言うが、機能は劣化版です。本家を使ったことはありませんが）。

# 補足

あれこれ言いましたが、全て Visual Studio に含まれている機能だったりします(^_^.)。
Visual Studio に負けない点としては「見せ方」かなと思っています。
現段階では、クラスメンバーの取得や図形作成処理に時間がかかってしまっていますので（特に画面クラス）、高速化対応が今後の課題となるかと思います。


# 配布

容量が大きいため、exe 形式では配布しません。Visual Studio 2015 以上の環境でビルドして、exe ファイルを作成して使ってください。

# 開発環境

本プログラムは、以下の環境で作成しています。

| 項目 | 値                                                               |
| ----- |:---------------------------------------------------- |
| OS   | Windows 8.1 Pro (64 bit)                              |
| IDE  | Visual Studio Community 2015                     |
| 言語 | VB.NET                                                       |
| 種類 | WPF アプリケーション (.NET Framework 4.6) |

Visual Studio 標準コントロールでしか確認していないので、
外部ベンダー提供のコントロールを使ったアプリケーションでもうまく動作するのかは未確認です。
また、.NET Framework 3.5 以下の Windows フォームアプリケーションを解析できるかどうかも未確認です。
(例えば、Windows 7 & Visual Studio 2005(.NET Framework 2.0) で作成したものを解析できるかどうか)

# 動作環境

開発環境でしか動かしていないため未確認ですが、Roslyn がインストールされているパソコンでしか動作しません。
また、Visual Studio で作成したソリューション一式＋一回以上ビルドしている状態（ビルドエラーが無い状態で、アセンブリファイルが各プロジェクト下に生成されていること）のみを解析対象としています。
読み取り専用で扱っていますが、生成されたアセンブリファイルを読み込みますので、頻繁にビルドする場合、コピーしたソリューション一式を読み込ませた方がいいかもしれません。

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


