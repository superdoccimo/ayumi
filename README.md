# 歩み値保存スクリプト

今回のプロジェクトでは、Pythonを活用してExcelデータの自動処理を実現しました。
このスクリプトは、特定のExcelファイルからデータを読み込み、重複を避けつつ新しい情報を追加するための自動化されたプロセスを提供します。
保存先は別のエクセルとデータベース（MySQL、MariaDBなど）の2つのバージョンを用意しました。
この技術を利用して、株価の歩み値をリアルタイムに保存します。
質問内容を工夫する必要がありますが、生成AI（ChatGPT、Claude3など）で作成できました。

## 前提条件

- Microsoft Excelがインストールされていること。
- pythonがインストールされていること。
- 楽天証券のマーケットスピード2にログインしていること。

### 使い方

このスクリプトを自分の環境で実行するには、少しの調整が必要です。以下は、スクリプトをカスタマイズして、ニーズに合わせるための簡単なステップです。

ファイルパスの設定:
スクリプト内で、データを読み込むExcelファイルやデータを保存するJSONファイルのパスは、あなたの環境に合わせて指定する必要があります。例えば、file_path や json_file_path などの変数を見つけ、それらをあなたのファイルシステムに存在する正確なパスに書き換えてください。
実行環境の準備:
このスクリプトはPythonで書かれており、実行にはPythonのインストールが必要です。また、pandas や openpyxl などのライブラリも使用していますので、これらがインストールされていない場合は、インストールが必要になります。（例:pip install openpyxl）
スクリプトの実行:
環境が準備できたら、ターミナルやコマンドプロンプトからスクリプトを実行してください。スクリプトがデータの読み込み、処理、保存を自動的に行います。
結果の確認:
スクリプトの実行が完了したら、指定した保存先のファイルを開いて、結果を確認してください。新しい情報が正しく追加されているかを確認し、必要に応じて調整を行います。
このスクリプトはエクセルが上書き保存されたタイミングで発動します。それゆえ、VBAで定期的（例えば5秒間隔）に上書き保存するコードを記述します。つまり、一定間隔で上書き保存するマクロを実行すれば歩み値を保存し続けてくれます。エクセルファイルには、必要なVBAマクロが既に組み込まれています。

## 参考資料

- **ブログ記事**: [初回投稿：2024年4月4日](https://minokamo.tokyo/2024/04/04/7075/)
  - このブログ記事でも、その活用方法について詳しく解説しています。

- **YouTube動画**: [初回投稿：2024年4月4日](https://youtu.be/owSIolD2nu8)
  - YouTubeでの解説動画では、実際にこのスクリプトを使用して歩み値を取得する様子を示しています。

## 貢献方法

このプロジェクトへの貢献や改善提案がある場合は、[Issueを作成](https://github.com/superdoccimo/ayumi/issues)するか、[Pull Request](https://github.com/superdoccimo/ayumi/pulls)を送ってください。
