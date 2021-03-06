[void][System.Reflection.Assembly]::LoadWithPartialName("System.Data")

# DSNを選択
$connectionsString = "DSN=CData SharePoint Source"
$odbcCon = New-Object System.Data.Odbc.OdbcConnection($connectionsString)
$odbcCon.Open();

$odbcCmd = New-Object System.Data.Odbc.OdbcCommand
$odbcCmd.Connection = $odbcCon

# コマンド実行  Cache Data 作成
$odbcCmd.CommandText = "CACHE SELECT * FROM Tasks"
$odbcCmd.ExecuteNonQuery() | Out-Null

# コマンド実行　Cacheテーブルからデータを取得
$odbcCmd.CommandText = "SELECT * FROM Tasks#Cache "
$odbcReader = $odbcCmd.ExecuteReader() 

while ($odbcReader.Read()) {
    $odbcReader["Id"].ToString()
}

$odbcReader.Read()

# コマンドオブジェクト破棄
$odbcCmd.Dispose()

# DB切断
$odbcCon.Close()
$odbcCon.Dispose()

