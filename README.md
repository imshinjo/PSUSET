（最初だけ）作業フォルダ内でクローンしてビルドする。
```
> git clone https://github.com/imshinjo/PSUSET.git
> cd .\PSUSET\
PSUSET> docker build -t psuset .
```

（ここから月次）「機能×カラー別集計レポート」は不具合があるので手入れする必要がある。
手動で開いて，「機器別サマリー」シートをコピーして，新規作成したExcel(.xlsx)に貼り付け，保存する。
新規作成したExcelのファイル名を「機能×カラー別集計レポート」としておく。元のファイル名をコピペすればよい。
シート名は編集しなくてよい。元ファイル(.xls)は削除しておく。

`statistics_report`には統計記入ファイルを
`number_report`にはプリンタサーバの利用集計ファイルを置く。
```
PSUSET> tree /f
C:.
│  Dockerfile
│  main.py
│
├─number_report
│      最新の機器カウンターレポート(JA)_YYYYMMDD_xxxxxx.csv
│      機能×カラー別集計レポート 月毎(JA)_YYYYMMDD_xxxxxx.xlsx # 元ファイルから[機器別サマリー]シートを移植した新規ファイル
│
└─statistics_report
        Ricohスキャナ統計.xlsx
        ロビープリンタ印刷統計.xlsx
        教員カラープリンタ印刷統計.xlsx
        教室等モノクロプリンタ印刷統計.xlsx
		
```

プログラムを実行する。下記の通り`PSUSET`ディレクトリを，コンテナ内の`/mnt`ディレクトリにバインドマウントする必要がある。
```
> docker run --rm -v .:/mnt psuset
```

以上