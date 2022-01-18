# OpenPOS-CSSO

これらはOpenPOS サービスオブジェクトを C# で作成するためのスケルトンコードです。
OpenPOS(UnifiedPOS) ver1.16の45デバイスすべてを用意しています。

開発途中のため、十分な検証はされておらず、機能・構成などは大きく変わる可能性があります。


## 開発/実行環境

このプログラムとそれを元にしたサービスオブジェクトの開発および実行には以下が必要です。

- Visual Studio 2022またはVisual Studio Community 2022 version 17.0.5
- .NET Framework 4.8
- OPOS-CCO ver1.16 の以下のDLLおよびそのPIA(Primary Interop Assembly)が登録されていること
   - Opos_Interfaces.dll
   - Opos_Internals.dll
   - Opos_Constants.dll


## 概要と既知の課題

OpenPOSの各デバイスクラスのサービスオブジェクトをC#で作成・レジストリ登録・動作させるために必要な共通的なソースコードのスケルトンコードを提供しています。

この枠組みを使って実際のデバイスのための処理を組み込むことで、プロトタイプを素早く作ることが出来ますし、製品に応用することも簡単になることを目指しています。

- メソッドのエントリポイントと共通的な処理は既に組み込み済みです。
- 対象PlatformにはAnyCPUを使わず、必ずx86/x64を明示的に指定してください。
- バーコードスキャナデバイスのソースコードは例として実際に動作するものを組み込んでいますが、デバイス依存の部分は分離独立させるつもりです。
- 最初は全てのデバイスが排他制御型の処理を組み込んでいます。CashDrawer/Keylock等の排他制御ではないデバイスの処理は今後組み込む予定です。
- 他にも機械的にコピー＆ペーストして各デバイスのソースコードを作成しているので仕様に合っていない部分があるかもしれません。
- 不具合を見つけたら情報共有をお願いします。


## カスタマイズ

- カスタマイズする場合、以下の内容は必ず変更してください。
   - プロジェクト名・ファイル名(CSxxxxSO)(xxxxはOpenPOSデバイスクラス名称:以下同じ)
   - namespace(OpenPOS.CSSO)
   - Guid
   - ProgId(OPOS.xxxx.OpenPOS.CSSO.CSxxxxSO.1)
   - AssemblyInfoの各情報(OpenPOS.CSSO.CSxxxxSOやその他情報)

## ライセンス

[zlib License](./LICENSE) の下でライセンスされています。
