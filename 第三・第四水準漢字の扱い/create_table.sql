/* 文字テーブル */
CREATE TABLE tbl_moji (
    moji NVARCHAR(2) NOT NULL -- 文字
  , jis VARCHAR(4) NOT NULL -- JISコード
  , kuten VARCHAR(6) -- 句点コード
  , high_level SMALLINT NOT NULL -- 高水準フラグ
  , PRIMARY KEY (kuten)
);
