# エクセル VBA を VSCode で

## 環境

- Windows 11
- PowerShell 7.4.0
- Gnu Make 3.81
- Git 2.43.0.windows.1

## 事前設定

1. 「リボンのユーザー設定」で、メインタブに「開発」リボンを追加

1. 「トラスト センター」で、VBA プロジェクト オブジェクト モデルへのアクセスを信頼する

## Excel ファイルの準備

以下のディレクトリ構成となるように、workbook.xlsm を設置する。

```
.
│  .gitignore
│  Makefile
│  README.md
│  Run-Macro.ps1
│  vbac.wsf
│
└─bin
      workbook.xlsm
```

worksheet.xlsm には "Hello, World!" を MsgBox に表示するマクロが記載されている。


## Excel から VBA を抽出・統合するスクリプト

![Ariawase](https://github.com/vbaidiot/ariawase)

```
Invoke-WebRequest https://raw.githubusercontent.com/vbaidiot/ariawase/master/vbac.wsf -OutFile vbac.wsf
```

## CLI ワークフロー

Excel から VBA を抽出する (初回に一回実行する)。

```
make decombine
```

Excel に local の VBA を統合する (随時実行する)。

```
make # make combine
```

Excel Macro を実行 (エクセルのパス と マクロ名は Makefile で定義)

```
make run
```

また、エクセルのパスとマクロ名を個別に指定して実行することも可能

```
.\Run-Macro.ps1 .\bin\workbook.xlsm {macro_name}
```


## Git

普通に Git で管理可能。
