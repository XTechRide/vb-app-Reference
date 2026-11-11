-- =====================================================================
--  Excel取込 DDL(サンプル)
--  前提:
--    ・Excel 1ファイル = 既存の本番ヘッダ1件。取込はそのヘッダ配下にレコードを追加する
--    ・Excel 47項目のうち KEY1〜KEY6 がキー項目、COL07〜COL47 がその他項目
--    ・3項目キー = KEY1, KEY2, KEY3(KEY3 のみ NULL になり得る)
--    ・6項目キー = KEY1〜KEY6
--    ・本番レコードは論理削除(DEL_FLG = '1')
--    ・一意制約は使わず、重複チェックと排他はアプリ層で行う
--  実際の列名・型に置き換えてください。
-- =====================================================================

CREATE SEQUENCE SEQ_IMPORT_BATCH;
CREATE SEQUENCE SEQ_IMPORT_ERR;

-- ---------------------------------------------------------------------
-- 本番ヘッダ(既存テーブルに VERSION / DEL_FLG がなければ追加)
--   VERSION:楽観的排他用。配下のレコードを登録・更新・削除する処理は
--            コミット直前に「VERSION = 読んだ値」を条件に +1 する
-- ---------------------------------------------------------------------
CREATE TABLE PROD_HEADER (
    HEADER_ID    NUMBER(18)     NOT NULL,
    -- ... ヘッダ項目 ...
    VERSION      NUMBER(18)     DEFAULT 0 NOT NULL,
    DEL_FLG      CHAR(1)        DEFAULT '0' NOT NULL,
    CREATED_BY   VARCHAR2(30)   NOT NULL,
    CREATED_AT   TIMESTAMP      DEFAULT SYSTIMESTAMP NOT NULL,
    UPDATED_BY   VARCHAR2(30),
    UPDATED_AT   TIMESTAMP,
    CONSTRAINT PK_PROD_HEADER PRIMARY KEY (HEADER_ID)
);
-- 既存テーブルへ追加する場合:
-- ALTER TABLE PROD_HEADER ADD (VERSION NUMBER(18) DEFAULT 0 NOT NULL);

-- ---------------------------------------------------------------------
-- 本番レコード
-- ---------------------------------------------------------------------
CREATE TABLE PROD_DETAIL (
    HEADER_ID    NUMBER(18)     NOT NULL,
    LINE_NO      NUMBER(6)      NOT NULL,
    KEY1         VARCHAR2(50)   NOT NULL,
    KEY2         VARCHAR2(50)   NOT NULL,
    KEY3         VARCHAR2(50),
    KEY4         VARCHAR2(50)   NOT NULL,
    KEY5         VARCHAR2(50)   NOT NULL,
    KEY6         VARCHAR2(50)   NOT NULL,
    COL07        VARCHAR2(200),
    -- ... COL08 〜 COL46 ...
    COL47        VARCHAR2(200),
    DEL_FLG      CHAR(1)        DEFAULT '0' NOT NULL,
    BATCH_ID     NUMBER(18),                -- どの取込で登録されたか(追跡用)
    CREATED_BY   VARCHAR2(30)   NOT NULL,
    CREATED_AT   TIMESTAMP      DEFAULT SYSTIMESTAMP NOT NULL,
    CONSTRAINT PK_PROD_DETAIL PRIMARY KEY (HEADER_ID, LINE_NO),
    CONSTRAINT FK_PROD_DETAIL_HEADER FOREIGN KEY (HEADER_ID) REFERENCES PROD_HEADER (HEADER_ID)
);
-- 重複チェック SQL の検索性能用(非ユニーク)
CREATE INDEX IX_PROD_DETAIL_KEY ON PROD_DETAIL (KEY1, KEY2, KEY3);

-- ---------------------------------------------------------------------
-- ワーク:BATCH_ID で取込ごとに区切る
--   STATUS:'HOLD' = 論理削除済みデータとキーが一致したため本番登録を保留中
--   (正常終了・エラー終了したバッチの行は残らない)
-- ---------------------------------------------------------------------
CREATE TABLE WK_IMPORT (
    BATCH_ID     NUMBER(18)     NOT NULL,
    ROW_NO       NUMBER(6)      NOT NULL,   -- Excel上の行番号
    HEADER_ID    NUMBER(18)     NOT NULL,   -- 取込先ヘッダ
    STATUS       VARCHAR2(10)   DEFAULT 'NEW' NOT NULL,
    KEY1         VARCHAR2(50)   NOT NULL,
    KEY2         VARCHAR2(50)   NOT NULL,
    KEY3         VARCHAR2(50),
    KEY4         VARCHAR2(50)   NOT NULL,
    KEY5         VARCHAR2(50)   NOT NULL,
    KEY6         VARCHAR2(50)   NOT NULL,
    COL07        VARCHAR2(200),
    -- ... COL08 〜 COL46 ...
    COL47        VARCHAR2(200),
    CREATED_BY   VARCHAR2(30)   NOT NULL,
    CREATED_AT   TIMESTAMP      DEFAULT SYSTIMESTAMP NOT NULL,
    CONSTRAINT PK_WK_IMPORT PRIMARY KEY (BATCH_ID, ROW_NO)
);
CREATE INDEX IX_WK_IMPORT_KEY ON WK_IMPORT (KEY1, KEY2, KEY3);

-- ---------------------------------------------------------------------
-- エラー履歴(保留の理由もここに残す)
-- ---------------------------------------------------------------------
CREATE TABLE IMPORT_ERROR_HISTORY (
    ERR_ID       NUMBER(18)     NOT NULL,
    BATCH_ID     NUMBER(18)     NOT NULL,
    HEADER_ID    NUMBER(18),
    FILE_NAME    VARCHAR2(260)  NOT NULL,
    ROW_NO       NUMBER(6),                 -- ファイル全体のエラーは NULL
    ERR_TYPE     VARCHAR2(30)   NOT NULL,   -- DUP3 / DUP6 / HOLD / HEADER / CONFLICT / SYSTEM
    MESSAGE      VARCHAR2(1000) NOT NULL,
    CREATED_BY   VARCHAR2(30)   NOT NULL,
    CREATED_AT   TIMESTAMP      DEFAULT SYSTIMESTAMP NOT NULL,
    CONSTRAINT PK_IMPORT_ERROR_HISTORY PRIMARY KEY (ERR_ID)
);
CREATE INDEX IX_IMPORT_ERR_BATCH ON IMPORT_ERROR_HISTORY (BATCH_ID);
