--概要　　　：指定された年の休日をカレンダーテーブルに追加します
--引数　　　：[@yyyy]…対象となる年
--戻り値　　：正常終了なら0、そうでなければ-1
--結果セット：例外が発生した場合、エラー情報
IF (EXISTS(SELECT * FROM sys.objects WHERE (type = 'P') AND (name = 'sp_holiday_setdata')))
BEGIN
  DROP PROCEDURE sp_holiday_setdata;
END
GO

CREATE PROCEDURE sp_holiday_setdata
  @yyyy INT
AS
BEGIN
  SET NOCOUNT ON;

/*
******************************
テンポラリテーブルにデータをセット
******************************
*/
  SET NOCOUNT ON;

  --プロシージャの戻り値を返す変数を定義します
  DECLARE @RTCD INT;
  SET @RTCD = 0;

  --対象年の文字列型を変数に格納します
  DECLARE @str_year VARCHAR(4);
  SET @str_year = CONVERT(VARCHAR, @yyyy);

  --日付データの作業用変数です
  DECLARE @date_work DATETIME;
  SET @date_work = CONVERT(DATETIME, @str_year + '-01-01');

  --土日を休日として登録します
  WHILE (0 = 0)
  BEGIN
    --年が変わるまで1日から繰り返し土日かどうかを判断し、土日であればカレンダーテーブルに登録します
    IF (@yyyy < YEAR(@date_work))
    BEGIN
      BREAK;
    END

    IF ((DATEPART(weekday, @date_work) = 1) OR (DATEPART(weekday, @date_work) = 7))
    BEGIN
      BEGIN TRY
        INSERT INTO カレンダー (対象日付, 対象区分) VALUES (@date_work, 1);
      END TRY
      BEGIN CATCH
        EXECUTE sp_returnerror 'sp_holiday_setdata:土日の追加に失敗しました。';
        RETURN -1;
      END CATCH
    END

    SET @date_work = DATEADD(d, 1, @date_work);
  END

  --年始1日
  SET @date_work = CONVERT(DATETIME, @str_year + '-01-01');
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --年始2日
  SET @date_work = CONVERT(DATETIME, @str_year + '-01-02');
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --年始3日
  SET @date_work = CONVERT(DATETIME, @str_year + '-01-03');
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --成人の日（1月の第2週月曜日）
  SET @date_work = dbo.fn_getdate_dayofweek(@yyyy, 1, 2, 2);
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --建国記念の日
  SET @date_work = CONVERT(DATETIME, @str_year + '-02-11');
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --春分の日
  SET @date_work = CONVERT(DATETIME, @str_year + '-03-' + CONVERT(VARCHAR, (dbo.fn_getdate_syunbun(@yyyy))));
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --昭和の日
  SET @date_work = CONVERT(DATETIME, @str_year + '-04-29');
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --憲法記念日
  SET @date_work = CONVERT(DATETIME, @str_year + '-05-03');
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --みどりの日
  SET @date_work = CONVERT(DATETIME, @str_year + '-05-04');
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --こどもの日
  SET @date_work = CONVERT(DATETIME, @str_year + '-05-05');
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --ハッピーマンデー
  If ((DATEPART(weekday, (CONVERT(DATETIME, @str_year + '-05-03'))) = 2) OR (DATEPART(weekday, (CONVERT(DATETIME, @str_year + '-05-04'))) = 2) OR (DATEPART(weekday, (CONVERT(DATETIME, @str_year + '-05-05'))) = 2))
  BEGIN
    SET @date_work = CONVERT(DATETIME, @str_year + '-05-06');
    EXECUTE @RTCD = sp_holiday_insert @date_work;
    IF (@RTCD = -1)
    BEGIN
      RETURN -1;
    END
  END

  --海の日（7月の第3週月曜日）
  SET @date_work = dbo.fn_getdate_dayofweek(@yyyy, 7, 3, 2);
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --山の日
  SET @date_work = CONVERT(DATETIME, @str_year + '-08-11');
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --敬老の日（9月の第3週月曜日）
  SET @date_work = dbo.fn_getdate_dayofweek(@yyyy, 9, 3, 2);
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --国民の休日（敬老の日と秋分の日が1日だけ空いていれば）
  IF (DAY(dbo.fn_getdate_dayofweek(@yyyy, 9, 3, 2)) = dbo.fn_getdate_syuubun(@yyyy))
  BEGIN
    SET @date_work = dbo.fn_getdate_dayofweek(@yyyy, 9, 3, 1);
    EXECUTE @RTCD = sp_holiday_insert @date_work;
    IF (@RTCD = -1)
    BEGIN
      RETURN -1;
    END
  END

  --秋分の日
  SET @date_work = CONVERT(DATETIME, @str_year + '-09-' + CONVERT(VARCHAR, (dbo.fn_getdate_syuubun(@yyyy))));
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --体育の日（10月の第2週月曜日）
  SET @date_work = dbo.fn_getdate_dayofweek(@yyyy, 10, 2, 2);
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --文化の日
  SET @date_work = CONVERT(DATETIME, @str_year + '-11-03');
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --敬老感謝の日
  SET @date_work = CONVERT(DATETIME, @str_year + '-11-23');
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --天皇誕生日
  SET @date_work = CONVERT(DATETIME, @str_year + '-12-23');
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --年末29日
  SET @date_work = CONVERT(DATETIME, @str_year + '-12-29');
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --年末30日
  SET @date_work = CONVERT(DATETIME, @str_year + '-12-30');
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  --年末31日
  SET @date_work = CONVERT(DATETIME, @str_year + '-12-31');
  EXECUTE @RTCD = sp_holiday_insert @date_work;
  IF (@RTCD = -1)
  BEGIN
    RETURN -1;
  END

  RETURN 0;
END
GO
