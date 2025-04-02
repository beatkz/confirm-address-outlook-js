# Microsoft Learn リンク

**Outlook Classicのサポート期限は2029年頃です。**

- [Visual Studio Code で Office アドインを開発する 2025/2/26追加](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/develop-add-ins-vscode)

- [Outlook アドインの送信時機能(en-US) 2024/5/20追加](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-on-send-addins?tabs=windows)
- [イベントベースのアクティベーション用にOutlookアドインを構成する(en-US) 2025/2/26追加](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/autolaunch)
- [Outlook アドインで Smart Alerts を使用して OnMessageSend および OnAppointmentSend イベントを処理する(en-US) 2025/2/26追加](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events?tabs=windows)
- [Officeアドインチュートリアル：メッセージを送信する前に添付ファイルを自動的にチェックする(en-US) 2025/2/26追加](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough?tabs=jsonmanifest)
- [Microsoft 365 の統合マニフェストを使用する Office アドインをサイドロードする(en-US) 2025/3/28追加](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest)

## VSCodeのターミナルを用いたnpmコマンドの使い方

- 機能ソースはapp/package.json内のscriptsセクションにあります。

|npm run ～|機能|
|:--:|:--:|
|build|ビルドを実行します。(本番環境用)|
|build:dev|デバッグ用ビルドを実行します。|
|dev-server|？|
|lint|コードの文法チェックをします。(*1)|
|lint:fix|コードの文法チェックをし、自動的に修正を試みます。(*1)|
|prettier|コードの整形をします。|
|signin|M365が有効なMSアカウントにサインインします。|
|signout|MSアカウントからサインアウトします。|
|start|Officeアドインのデバッグを開始します。(*2)|
|stop|Officeアドインのデバッグを安全に停止させます。|
|validate|Officeアドインのmanifest.jsonの文法チェックをします。(*1)|
|watch|？|

npm update: 依存関係のあるnodeモジュールの更新を行います。

- *1：Unixスタイルでの表示。エラーがある行のみを表示します。
- *2：既に起動しているOutlookがある場合は閉じてから実行します。

## Officeアドインのサイドロード確認

- Outlook Classicの場合、以下の順序で辿り、"Confirm-Address for Outlook.js"があるかを確認する。
  ホーム→アプリ→すべてのアプリ→アプリの追加→アプリを管理
