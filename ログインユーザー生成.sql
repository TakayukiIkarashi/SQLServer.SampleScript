-- [master]データベースに接続します
USE master;

-- SQL Serverに接続するためのログインアカウント[hoge_login]を生成します
CREATE LOGIN hoge_login WITH PASSWORD = '[ここにパスワードを指定]';

-- [hoge]データべースに接続します
USE hoge;

-- SQL Serverにユーザー[hoge_user]を生成し、先ほど生成したログインアカウント[hoge_login]に紐づけます
CREATE USER hoge_user FROM LOGIN hoge_login;

-- 生成したユーザー[hoge_user]に、[hoge]のデータベース所有者権限を与えます
ALTER ROLE db_owner ADD MEMBER hoge_user;
