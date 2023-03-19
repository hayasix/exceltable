======
README
======

| 2016-08-12 林秀樹
| 2023-03-19 林秀樹


このアーカイブについて
======================

このアーカイブには、Microsoft Excel のワークシート上に記述された
表の内容を読み取るための Python パッケージを収録しています。

このアーカイブに含まれるプログラム、データ、本書を含む文書その他の
ファイルは、特に記載がない限り、林秀樹が著作権を有し、所有権を
留保しています。

    Copyright (C) 2013 HAYASI Hideki.  All rights reserved.

Python は Python Software Foundation の登録商標または商標であり、
その知的財産権は同団体が保有し管理しています。
詳しくは http://www.python.org/about/legal/ をご覧ください。


動作条件
========

- Python 3.7+


ライセンス
==========

Zope Public License (ZPL) Version 2.1 を採用しています。


収録物
======

:README.rst:

    本書

:LICENSE:

    Zope Public License (ZPL) Version 2.1

:exceltable/:

    パッケージ本体

:setup.py:

    インストールスクリプト


使い方
======

``exceltable.reader`` は次のクラスを提供します。

:Reader(source, sheet, password=None, start_row=0, stop_row='', start_col='$A', stop_col='', header_rows=1, empty=None, repeat=False, trim=True):

    指定のファイル・シート・範囲に記述された表の内容を読み取り、
    各レコードの内容を順次返すジェネレーターを生成します。
    表の 1 行目がフィールド名、2 行目以降が各フィールドの値となります。
    ジェネレーターは各レコードを、各フィールド名を属性として持つ
    CSVRecord(namedtuple) オブジェクトとして返します。

:DictReader(source, sheet, password=None, start_row=0, stop_row='', start_col='$A', stop_col='', header_rows=1, empty=None, repeat=False, trim=True):

    指定のファイル・シート・範囲に記述された表の内容を読み取り、
    各レコードの内容を順次返すジェネレーターを生成します。
    表の 1 行目がフィールド名、2 行目以降が各フィールドの値となります。
    ジェネレーターは各レコードを、各フィールド名とその値からなる
    OrderedDict オブジェクトとして返します。


コマンドラインツール
========================

インストールスクリプトを実行すると::

    python setup.py install

コマンドラインツール ``exceltable`` が利用可能になります。このコマンド
により、指定のファイル・シート・範囲に記述された表の内容を読み取り、
全レコードを標準出力へ CSV 形式で書き出せます。::

    C:> python exceltable.py --header-rows=2
    HHHH,HHHHHH,HHHHHH,HHHHHH,HHHHHH,HHHHHH
    xxxx,xxxxxx,xxxxxx,xxxxxx,xxxxxx,xxxxxx
    xxxx,xxxxxx,xxxxxx,xxxxxx,xxxxxx,xxxxxx
    xxxx,xxxxxx,xxxxxx,xxxxxx,xxxxxx,xxxxxx
