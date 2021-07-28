--概要　　　：引数に指定されたテンポラリテーブルに休日をセットします（振替休日を考慮します）
--引数　　　：[@日付]…テンポラリテーブルに追加する休日
--戻り値　　：正常終了なら0、そうでなければ-1
--結果セット：例外が発生した場合、エラー情報
IF (EXISTS(SELECT * FROM sys.objects WHERE (type = 'P') AND (name = 'sp_holiday_insert')))
BEGIN
  DROP PROCEDURE sp_holiday_insert;
END
GO

CREATE PROCEDURE sp_holiday_insert
  @日付 DATETIME
AS
BEGIN
  SET NOCOUNT ON;

  --引数に指定されている休日をテンポラリテーブルに追加します
  DECLARE @sql_ins VARCHAR(8000);
  SET @sql_ins = '';
  SET @sql_ins = @sql_ins + ' IF (NOT EXISTS(SELECT * FROM 休日 WHERE 日付 = ''' + CONVERT(VARCHAR, @日付, 120) + '''))';
  SET @sql_ins = @sql_ins + ' INSERT INTO';
  SET @sql_ins = @sql_ins + '   休日';
  SET @sql_ins = @sql_ins + ' (';
  SET @sql_ins = @sql_ins + '   日付';
  SET @sql_ins = @sql_ins + ' )';
  SET @sql_ins = @sql_ins + ' VALUES';
  SET @sql_ins = @sql_ins + ' (';
  SET @sql_ins = @sql_ins + '   ''' + CONVERT(VARCHAR, @日付, 120) + '''';
  SET @sql_ins = @sql_ins + ' )';

  BEGIN TRY
    EXECUTE (@sql_ins);
  END TRY
  BEGIN CATCH
    RETURN -1;
  END CATCH

  --引数に指定されている日付が日曜日であれば、翌日を振替休日としてテンポラリテーブルに追加します
  DECLARE @weekday INT;
  SET @weekday = DATEPART(weekday, @日付);
  IF (@weekday = 1)
  BEGIN
    DECLARE @month INT;
    SET @month = MONTH(@日付);

    DECLARE @day INT;
    SET @day = DAY(@日付);

    --ただし、三が日には振替休日がありません
    IF NOT ((@month = 1) AND (@day = 3))
    BEGIN
      SET @日付 = DATEADD(day, 1, @日付);

      DECLARE @sql_ins2 VARCHAR(8000);
      SET @sql_ins2 = '';
      SET @sql_ins2 = @sql_ins2 + ' IF (NOT EXISTS(SELECT * FROM 休日 WHERE 日付 = ''' + CONVERT(VARCHAR, @日付, 120) + '''))';
      SET @sql_ins2 = @sql_ins2 + ' INSERT INTO';
      SET @sql_ins2 = @sql_ins2 + '   休日';
      SET @sql_ins2 = @sql_ins2 + ' (';
      SET @sql_ins2 = @sql_ins2 + '   日付';
      SET @sql_ins2 = @sql_ins2 + ' )';
      SET @sql_ins2 = @sql_ins2 + ' VALUES';
      SET @sql_ins2 = @sql_ins2 + ' (';
      SET @sql_ins2 = @sql_ins2 + '   ''' + CONVERT(VARCHAR, @日付, 120) + '''';
      SET @sql_ins2 = @sql_ins2 + ' )';

      BEGIN TRY
        EXECUTE (@sql_ins2);
      END TRY
      BEGIN CATCH
        RETURN -1;
      END CATCH
    END
  END

  RETURN 0;
END
GO
