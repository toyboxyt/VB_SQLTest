Imports System.Data.OleDb

Module Module1

    Sub Main()
        '▼データ取得
        Dim Cn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Database\Animals.mdb")
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable

        SQLCm.CommandText = "SELECT 説明 FROM T_目マスタ"
        Adapter.Fill(Table)

        ''▼値の表示
        'DataGridView1.DataSource = Table

        '▼後処理
        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()
    End Sub

End Module
