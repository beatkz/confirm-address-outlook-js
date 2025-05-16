# Microsoft Learn リンク

**Outlook Classicのサポート期限は2029年頃です。**

- [Visual Studio Code で Office アドインを開発する 2025/2/26追加](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/develop-add-ins-vscode)

- [Outlook アドインの送信時機能(en-US) 2024/5/20追加](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-on-send-addins?tabs=windows)
- [イベントベースのアクティベーション用にOutlookアドインを構成する(en-US) 2025/2/26追加](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/autolaunch)
- [Outlook アドインで Smart Alerts を使用して OnMessageSend および OnAppointmentSend イベントを処理する(en-US) 2025/2/26追加](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events?tabs=windows)
- [Officeアドインチュートリアル：メッセージを送信する前に添付ファイルを自動的にチェックする(en-US) 2025/2/26追加](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough?tabs=jsonmanifest)
- [Microsoft 365 の統合マニフェストを使用する Office アドインをサイドロードする(en-US) 2025/3/28追加](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest)
- [Microsoft 365 の統合アプリ マニフェストを使用した Office アドイン(en)](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/unified-manifest-overview)
- [Microsoft 365 アプリ マニフェスト スキーマ リファレンス(v1.21)(en)](https://learn.microsoft.com/en-us/microsoft-365/extensibility/schema/?view=m365-app-1.21)
- [OfficeアドインのUXデザインパターン 2025/4/16追加](https://learn.microsoft.com/en-us/office/dev/add-ins/design/ux-design-pattern-templates)

## Grok(X AI社のLLM AI)スレッド

- [Outlookアドインスレッド(v3)](https://grok.com/chat/d6b0bd02-c6df-4ad8-9ab7-ace0b1f7c172)
- [Outlookアドインスレッド(v2)](https://grok.com/chat/b79c016d-375b-4c64-9aef-f34d0d12d6f2)
- [Outlookアドインスレッド(v1)](https://grok.com/chat/0e2d4fde-2e06-4f10-a343-ebcfe2f22b92)

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
