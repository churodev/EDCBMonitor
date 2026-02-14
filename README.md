# EDCBMonitor

EDCB (EpgDataCap_Bon) の予約状況を確認・モニタリングするためのツールです。

* **Latest Release:** [https://github.com/churodev/EDCBMonitor/releases](https://github.com/churodev/EDCBMonitor/releases)

## ■ 概要

本ソフトはパイプ通信とファイル監視を組み合わせたEDCB専用ツールです。
パイプ通信により「予約の有効/無効」や「削除」を操作して本体へ反映できるため
EpgTimerを開く回数を削減できます。

また、OSのファイル変更通知をフックしてReserve.txtの更新を監視して
変更があった場合にパイプ通信により予約情報を取得更新する設計です。
そのため殆どの時間が待機時間となりPCへの負荷を最小限に抑えた設計となっています。

## ■ 動作環境

* **OS**: Windows 10 / 11 (64bit)
* **必須環境**: EDCB (EpgDataCap_Bon) が導入済みであること

### 配布ファイルについて

本ソフトは2種類のパッケージを配布しています。環境に合わせて選択してください。

1. **ランタイム同梱 complete版 (Self-contained)**
* `.NET 8` ランタイムを内包しています。
* 別途ランタイムをインストールする必要がなく解凍してそのまま実行可能です。


2. **軽量版 (Framework-dependent)**
* 実行には **[.NET Desktop Runtime 8.0 (x64)](https://dotnet.microsoft.com/en-us/download/dotnet/8.0)** のインストールが必要です。
* すでに環境にランタイムが入っている場合はこちらの方がファイルサイズが小さく軽量です。



## ■ 使い方

### 1. インストール

ダウンロードしたzipファイルを解凍し、中にある `EDCBMonitor.exe` を任意のフォルダに配置してください。

### 2. 初回起動と設定

`EDCBMonitor.exe` を実行します。
初回設定を行うと同ディレクトリ内に設定ファイル `EM_Config.xml` が自動生成されます。
（※設定ファイルは必ずexeと同じ場所に置いてください。）

### 3. EDCBとの連携

設定画面にてご使用のEDCBフォルダ内にある `Reserve.txt` の場所（パス）を指定してください。
OSのファイル変更通知をフックしEDCB側の予約変動を瞬時に反映するようになります。

### 4. TVTestとの連携

設定画面にてご使用のTVTestの「TVTest.exe」がある場所（パス）を指定してください。
録画中の予約を右クリックして追っかけ再生を選択するとTVTestが起動して再生します。
追っかけ再生にはTVTestとTvtPlayの導入が必要です。
*  ** [https://github.com/DBCTRADO/TVTest](https://github.com/DBCTRADO/TVTest)
*  ** [https://github.com/xtne6f/TvtPlay](https://github.com/xtne6f/TvtPlay)

## ■ 自動起動の設定

PC起動時に本ソフトを自動的に起動したい場合はWindowsの「スタートアップ」フォルダに `EDCBMonitor.exe` のショートカットを登録してください。

**スタートアップフォルダの開き方:**

1. キーボードの `Windowsキー` + `R` を押す
2. 「ファイル名を指定して実行」ダイアログに `shell:startup` と入力してOKを押す

## ■ アップデート時の注意

バージョン更新を行う際は設定の整合性を保つため念のため既存の `EM_Config.xml` を削除してから、新しいバージョンのexeで再設定を行うことを推奨します。

## ■ 免責事項

本ソフトウェアを使用したことによって生じたすべての障害・損害・不具合等に関して作者は一切の責任を負いません。自己責任でご利用ください。
