# slack-export2excel

## What is this ?

Slack からエクスポートしたログに対して `Slack Dicovery API` を用いて以下の操作を行います。

* 調査対象ユーザの既に退出済みのものを含めたすべての所属チャットチャンネル(パブリック/プライベート/DM/グループDM)の一覧を取得し、csv で出力する
* 上記で取得したリストをもとに、Slackのエクスポートデータから対象のログ(json)を抽出する
* 取得したjsonファイルの内容をパースし、エクセルへ出力する。調査対象ユーザのメッセージには背景色を設定する。

なお、このスクリプトは `Slack Enterprise Grid` プランで利用できるOrGレベルでのエクスポートを前提としています。  
ref: https://slack.com/intl/ja-jp/help/articles/201658943  

# requirements

## Slack API Token

* `disocovery:read` の scope を付与した API Token

## Python Environment

* Python : 3.6 ↑
* openpyxl : latest

# How To Use

## 事前準備

### エクスポートデータの用意

Slack のエクスポートデータを取得、 zip を解凍し、このスクリプトのディレクトリにある export/ 配下においてください。  
※ディレクトリ名にスペースが入っているとよくないことが起こりそう（未確認）なので、スペースを除去しておくことをお勧めします  
例） `export/Sep_28_2018-Jun_1_2021`    

### config.ini の用意

* config.ini.sample をコピーし、 config.ini として保存してください
* config.ini  に必要なパラメータを設定してください
```
[slack]
token=xoxp-1111111111111-2222222222222-3333333333333-aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa ★API token 
org_id=<Your OrG ID> 
org_name=<Your OrG Name>
org_domain=<Your OrG Domain>
export_dir=export/<export data dir>
```

## 実行

```
$ ./enterprisegrid.py
ID: <調査対象ユーザのユーザID>
```

`./main.py` を実行すると、ユーザIDの入力を待ちます。ここに調査対象としたい Slack の ユーザID を入力してください。  
存在しないユーザIDを入力した場合は、エラーを出して終了します。  
以降、STEP1 ~ STEP5 までで、状況を画面に出力しながら実行します。  

## 出力されるログ xlsx について

元のログ json ファイルをパースして、以下の情報をエクセルに時系列で出力します。  
また、当該ユーザの発言の場合は、当該行の背景色を黄色にします。  

## 出力エクセルフォーマット

* 1つのチャットチャンネルのディレクトリに一つのエクセルファイルが作成されます
* 1つのエクセルワークブックに一つのチャットの履歴
* 1つのエクセルシートに、1日分のチャットの履歴（シート名が日付、ただしこの日付は **UTC** ）
* 表の1行に1つの Slack への投稿が表示されます。以下の並びで表示されます

`datetime(JST), type, subtype, user_id, user_name, thread, text, files`  
| header | 内容 |
| - | - |
| datetime | 当該発言時刻 (JST) |
| type | Slackログのメタデータ、詳細は Slack に確認中 <br> ※`message` 以外存在しない |
| subtype | Slackログのメタデータ <br> ref: https://api.slack.com/events/message |
| user_id | 発言したユーザのユーザID |
| user_name | 発言したユーザの表示名 |
| thread | スレッドでの発言の場合、どのスレッドかを示す数字列(float) <br> ※厳密には *thread_ts* という名前の unixtime |
| text | チャット発言内容 |
| files | ファイルが添付されている場合、そのファイル名 | 


## 注意事項など(確認中の内容含む)

* 時刻表示について  
  元の jsonログファイル がUTCで日付別のファイルとなっており、このファイル単位でエクセルへの挿入を行うため、シートの日付区切りは **UTC** になっています。
  一方チャットログとしては **JST** で表示されています。ご注意ください。※実際に内容見るときはJSTのほうがわかりやすいでしょう。  

* `subtype` のいくつかについて

  * `message_changed` について
    Slack ログの都合上、 `message_changed` がついている行が **変更前の元のメッセージ** を出力しており、  
    それよりも前にあるもともとの時刻のメッセージは **変更後のメッセージ** が出力されています。  
    だってログがそうなってるんだもん。。  

  * `message_changed` になっているが、内容に変更がないものがある
    Slack に仕様確認中。実際にテキスト内容に変更はなく、そのほかの何らかのメタデータの変更のようです。  

  * `message_delete` について
    **メッセージを削除した** というログはあるものの、「どのメッセージを削除したか」が当該ログからわからないケースがあるようです。
    削除されたメッセージは、OrG/ワークスペースのメッセージ保存ポリシーの設定で **保存** になっていれば、ログには残されています。

* エクセルで表示できない文字列について  
    openpyxl で扱えない文字があった場合、当該メッセージを削除して削除されたことを示すメッセージを挿入しています。  
    `"Sanitized openpyxl.utils.exceptions.IllegalCharacterError"` という文字列になっています。  
    ※Slack のコードスニペットを使ってバイナリ文字列をPOSTしていた場合などに発生します

* チャンネル一覧の csv にはあるが、エクスポートデータにはないDMがある  
これは「DMを開始したが、実際に発言はしなかった」場合に発生します。実際にはDMは行われていないため、気にしなくて大丈夫です。

* チャンネル一覧の csv にはあるが、チャンネルが存在しないものがある  
    Slackに確認中。チャンネル一覧 csv 上では、チャンネル名が `channel_not_found` と表示されています。

* チャンネル一覧 csv について、文字コードをBOM付きUTF-8 `(utf_8_sig)` で出力します。これは、日本語環境を前提としており、 csv を開く際にBOM付きでないと文字化けするためです。


# Todo

* READMEを英語で書く
* 一部仕様確認中の項目の完成
* EGプラン以外の対応（ディレクトリ構成が異なる）
