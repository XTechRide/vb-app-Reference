Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Data
Imports System.Linq
Imports Oracle.ManagedDataAccess.Client   ' NuGet: Oracle.ManagedDataAccess (ODP.NET Managed)

''' <summary>
''' Excel取込(1ファイル = 既存ヘッダ1件)
'''   ワーク登録 → 3項目チェック → 6項目チェック → 論理削除データとの一致チェック
'''     ├ 重複あり         : ロールバック、エラー履歴に登録して終了
'''     ├ 論理削除と一致   : 本番には登録せずワークに STATUS='HOLD' で残す(保留)
'''     └ 問題なし         : 本番レコードに追加 → ワーク削除
'''
''' 楽観的排他(DB の一意制約は使わない):
'''   ・開始時に取込先ヘッダの VERSION をロックせずに読む
'''   ・コミット直前に UPDATE PROD_HEADER SET VERSION = VERSION + 1 WHERE VERSION = 読んだ値
'''     → 0 件なら、取込中に他ユーザーが同じヘッダを変更した → 競合としてロールバック
'''   ・同じヘッダへ同時に取り込むと、後の人は先の人のコミットを待ったうえで 0 件になる
'''
''' 前提:同じヘッダ配下のレコードを変更する他の画面・処理も、ヘッダの VERSION を +1 すること
''' </summary>
Public Class ExcelImportService

    ' ---- 列定義(47項目)。実際の列名・型に合わせて書き換える ----
    Private Class ColumnDef
        Public ReadOnly Name As String
        Public ReadOnly DbType As OracleDbType
        Public Sub New(name As String, dbType As OracleDbType)
            Me.Name = name : Me.DbType = dbType
        End Sub
    End Class

    Private Shared ReadOnly Columns As List(Of ColumnDef) = BuildColumns()

    Private Shared Function BuildColumns() As List(Of ColumnDef)
        Dim list As New List(Of ColumnDef)
        For i As Integer = 1 To 6
            list.Add(New ColumnDef("KEY" & i, OracleDbType.Varchar2))
        Next
        For i As Integer = 7 To 47
            list.Add(New ColumnDef("COL" & i.ToString("00"), OracleDbType.Varchar2))
        Next
        Return list
    End Function

    ''' <summary>Excelの1データ行目の行番号(見出し1行なら2)</summary>
    Private Const FirstDataRowNo As Integer = 2

    Private ReadOnly _connectionString As String

    Public Sub New(connectionString As String)
        _connectionString = connectionString
    End Sub

    ' =================================================================
    '  公開メソッド
    ' =================================================================

    ''' <param name="source">Excelから読み込んだ DataTable(列名 = Columns の Name)</param>
    ''' <param name="headerId">取込先の既存ヘッダ</param>
    Public Function Import(source As DataTable, headerId As Long, fileName As String, userId As String) As ImportResult
        Using conn As New OracleConnection(_connectionString)
            conn.Open()
            Dim batchId As Long = NextVal(conn, "SEQ_IMPORT_BATCH")
            Dim errors As New List(Of ImportError)

            ' 0. 取込先ヘッダの VERSION を読む(重複チェックより前・ロックしない)
            Dim startVersion As Long? = ReadHeaderVersion(conn, headerId)
            If Not startVersion.HasValue Then
                errors.Add(New ImportError(Nothing, "HEADER", $"取込先ヘッダ(ID={headerId})が存在しないか、削除されています。"))
                WriteHistory(conn, batchId, headerId, fileName, userId, errors)
                Return ImportResult.Failed(batchId, errors)
            End If

            Dim tx As OracleTransaction = conn.BeginTransaction(IsolationLevel.ReadCommitted)
            Try
                ' 1. ワーク登録
                InsertWork(conn, batchId, headerId, source, userId)

                ' 2. 3項目キーチェック(KEY3 が NULL の行はスルー、論理削除済みは対象外)
                errors.AddRange(QueryMessages(conn, batchId, Sql3KeyCheck, "DUP3"))

                ' 3. 6項目キーチェック(3項目でエラーがあれば実施しない)
                If errors.Count = 0 Then
                    errors.AddRange(QueryMessages(conn, batchId, Sql6KeyCheck, "DUP6"))
                End If

                If errors.Count > 0 Then
                    tx.Rollback()   ' ワーク分も消える
                    WriteHistory(conn, batchId, headerId, fileName, userId, errors)
                    Return ImportResult.Failed(batchId, errors)
                End If

                ' 4. 論理削除済みの本番レコードとキーが一致する行があれば、ワークで止める
                Dim holds As List(Of ImportError) = QueryMessages(conn, batchId, SqlDeletedMatch, "HOLD")
                If holds.Count > 0 Then
                    SetWorkStatus(conn, batchId, "HOLD")
                    If Not TryIncrementHeaderVersion(conn, headerId, startVersion.Value, userId) Then
                        Return Conflict(conn, tx, batchId, headerId, fileName, userId)
                    End If
                    tx.Commit()     ' ワークだけ確定(本番には登録しない)
                    WriteHistory(conn, batchId, headerId, fileName, userId, holds)
                    Return ImportResult.HeldInWork(batchId, holds)
                End If

                ' 5. 本番レコードに追加 → ワーク削除
                InsertDetail(conn, batchId, headerId, userId)
                DeleteWork(conn, batchId)

                ' 6. 楽観的排他チェック(コミット直前)
                If Not TryIncrementHeaderVersion(conn, headerId, startVersion.Value, userId) Then
                    Return Conflict(conn, tx, batchId, headerId, fileName, userId)
                End If

                tx.Commit()
                Return ImportResult.Succeeded(batchId)

            Catch ex As Exception
                tx.Rollback()
                errors.Add(New ImportError(Nothing, "SYSTEM", Truncate(ex.Message, 1000)))
                Try
                    WriteHistory(conn, batchId, headerId, fileName, userId, errors)
                Catch
                    ' エラー履歴の書き込み失敗で元の例外を握りつぶさない
                End Try
                Throw
            Finally
                tx.Dispose()
            End Try
        End Using
    End Function

    Private Function Conflict(conn As OracleConnection, tx As OracleTransaction, batchId As Long, headerId As Long,
                              fileName As String, userId As String) As ImportResult
        tx.Rollback()
        Dim errors As New List(Of ImportError) From {
            New ImportError(Nothing, "CONFLICT", "取込中に他のユーザーが同じヘッダのデータを更新しました。もう一度取込んでください。")
        }
        WriteHistory(conn, batchId, headerId, fileName, userId, errors)
        Return ImportResult.Failed(batchId, errors)
    End Function

    ' =================================================================
    '  チェック SQL
    '   ・本番は DEL_FLG = '0'(有効データ)だけを重複の対象にする
    '   ・ワークは自分以外のバッチ(保留中の HOLD 行を含む)と照合する
    ' =================================================================

    ' 3項目:KEY3 IS NOT NULL の行だけが対象
    Private Const Sql3KeyCheck As String =
"SELECT MIN(ROW_NO) AS ROW_NO,
        'Excel内で3項目キーが重複しています(行: ' || LISTAGG(ROW_NO, ',') WITHIN GROUP (ORDER BY ROW_NO) || ')' AS MSG
   FROM WK_IMPORT
  WHERE BATCH_ID = :BATCH_ID AND KEY3 IS NOT NULL
  GROUP BY KEY1, KEY2, KEY3
 HAVING COUNT(*) > 1
UNION ALL
SELECT w.ROW_NO, 'ワーク(他の取込)に同じ3項目キーが存在します'
  FROM WK_IMPORT w
 WHERE w.BATCH_ID = :BATCH_ID AND w.KEY3 IS NOT NULL
   AND EXISTS (SELECT 1 FROM WK_IMPORT o
                WHERE o.BATCH_ID <> :BATCH_ID
                  AND o.KEY1 = w.KEY1 AND o.KEY2 = w.KEY2 AND o.KEY3 = w.KEY3)
UNION ALL
SELECT w.ROW_NO, '本番に同じ3項目キーが存在します'
  FROM WK_IMPORT w
 WHERE w.BATCH_ID = :BATCH_ID AND w.KEY3 IS NOT NULL
   AND EXISTS (SELECT 1 FROM PROD_DETAIL p
                WHERE p.DEL_FLG = '0'
                  AND p.KEY1 = w.KEY1 AND p.KEY2 = w.KEY2 AND p.KEY3 = w.KEY3)
ORDER BY 1"

    ' 6項目:KEY3 は NULL 同士も一致とみなす(DECODE は NULL=NULL を真と扱う)
    Private Const Sql6KeyCheck As String =
"SELECT MIN(ROW_NO) AS ROW_NO,
        'Excel内で6項目キーが重複しています(行: ' || LISTAGG(ROW_NO, ',') WITHIN GROUP (ORDER BY ROW_NO) || ')' AS MSG
   FROM WK_IMPORT
  WHERE BATCH_ID = :BATCH_ID
  GROUP BY KEY1, KEY2, KEY3, KEY4, KEY5, KEY6
 HAVING COUNT(*) > 1
UNION ALL
SELECT w.ROW_NO, 'ワーク(他の取込)に同じ6項目キーが存在します'
  FROM WK_IMPORT w
 WHERE w.BATCH_ID = :BATCH_ID
   AND EXISTS (SELECT 1 FROM WK_IMPORT o
                WHERE o.BATCH_ID <> :BATCH_ID
                  AND o.KEY1 = w.KEY1 AND o.KEY2 = w.KEY2 AND DECODE(o.KEY3, w.KEY3, 1, 0) = 1
                  AND o.KEY4 = w.KEY4 AND o.KEY5 = w.KEY5 AND o.KEY6 = w.KEY6)
UNION ALL
SELECT w.ROW_NO, '本番に同じ6項目キーが存在します'
  FROM WK_IMPORT w
 WHERE w.BATCH_ID = :BATCH_ID
   AND EXISTS (SELECT 1 FROM PROD_DETAIL p
                WHERE p.DEL_FLG = '0'
                  AND p.KEY1 = w.KEY1 AND p.KEY2 = w.KEY2 AND DECODE(p.KEY3, w.KEY3, 1, 0) = 1
                  AND p.KEY4 = w.KEY4 AND p.KEY5 = w.KEY5 AND p.KEY6 = w.KEY6)
ORDER BY 1"

    ' 論理削除済み(DEL_FLG = '1')の本番レコードとキーが一致する行 → 保留対象
    Private Const SqlDeletedMatch As String =
"SELECT w.ROW_NO, '本番の論理削除済みデータと3項目キーが一致したため、ワークで保留しました'
  FROM WK_IMPORT w
 WHERE w.BATCH_ID = :BATCH_ID AND w.KEY3 IS NOT NULL
   AND EXISTS (SELECT 1 FROM PROD_DETAIL p
                WHERE p.DEL_FLG = '1'
                  AND p.KEY1 = w.KEY1 AND p.KEY2 = w.KEY2 AND p.KEY3 = w.KEY3)
UNION ALL
SELECT w.ROW_NO, '本番の論理削除済みデータと6項目キーが一致したため、ワークで保留しました'
  FROM WK_IMPORT w
 WHERE w.BATCH_ID = :BATCH_ID
   AND EXISTS (SELECT 1 FROM PROD_DETAIL p
                WHERE p.DEL_FLG = '1'
                  AND p.KEY1 = w.KEY1 AND p.KEY2 = w.KEY2 AND DECODE(p.KEY3, w.KEY3, 1, 0) = 1
                  AND p.KEY4 = w.KEY4 AND p.KEY5 = w.KEY5 AND p.KEY6 = w.KEY6)
ORDER BY 1"

    ' =================================================================
    '  DB 操作
    ' =================================================================

    ''' <summary>有効なヘッダの VERSION。存在しない/論理削除済みなら Nothing。</summary>
    Private Shared Function ReadHeaderVersion(conn As OracleConnection, headerId As Long) As Long?
        Using cmd As New OracleCommand("SELECT VERSION FROM PROD_HEADER WHERE HEADER_ID = :HEADER_ID AND DEL_FLG = '0'", conn)
            cmd.BindByName = True
            cmd.Parameters.Add("HEADER_ID", OracleDbType.Int64).Value = headerId
            Dim v As Object = cmd.ExecuteScalar()
            If v Is Nothing OrElse v Is DBNull.Value Then Return Nothing
            Return Convert.ToInt64(v)
        End Using
    End Function

    ''' <summary>読んだ時点から VERSION が変わっていなければ +1 して True。変わっていれば False。</summary>
    Private Shared Function TryIncrementHeaderVersion(conn As OracleConnection, headerId As Long, expected As Long, userId As String) As Boolean
        Const sql As String =
            "UPDATE PROD_HEADER SET VERSION = VERSION + 1, UPDATED_BY = :USER_ID, UPDATED_AT = SYSTIMESTAMP " &
            " WHERE HEADER_ID = :HEADER_ID AND VERSION = :EXPECTED AND DEL_FLG = '0'"
        Using cmd As New OracleCommand(sql, conn)
            cmd.BindByName = True
            cmd.Parameters.Add("USER_ID", OracleDbType.Varchar2).Value = userId
            cmd.Parameters.Add("HEADER_ID", OracleDbType.Int64).Value = headerId
            cmd.Parameters.Add("EXPECTED", OracleDbType.Int64).Value = expected
            Return cmd.ExecuteNonQuery() = 1
        End Using
    End Function

    ''' <summary>配列バインドで1回のラウンドトリップで全行 INSERT</summary>
    Private Sub InsertWork(conn As OracleConnection, batchId As Long, headerId As Long, source As DataTable, userId As String)
        Dim n As Integer = source.Rows.Count
        If n = 0 Then Throw New InvalidOperationException("取込対象の行がありません。")

        Dim names As IEnumerable(Of String) = Columns.Select(Function(c) c.Name)
        Dim sql As String =
            "INSERT INTO WK_IMPORT (BATCH_ID, ROW_NO, HEADER_ID, " & String.Join(", ", names) & ", CREATED_BY) " &
            "VALUES (:BATCH_ID, :ROW_NO, :HEADER_ID, " & String.Join(", ", names.Select(Function(x) ":" & x)) & ", :CREATED_BY)"

        Using cmd As New OracleCommand(sql, conn)
            cmd.BindByName = True
            cmd.ArrayBindCount = n

            Dim batchIds(n - 1) As Object, rowNos(n - 1) As Object, headerIds(n - 1) As Object, users(n - 1) As Object
            For i As Integer = 0 To n - 1
                batchIds(i) = batchId
                rowNos(i) = i + FirstDataRowNo
                headerIds(i) = headerId
                users(i) = userId
            Next
            cmd.Parameters.Add("BATCH_ID", OracleDbType.Int64).Value = batchIds
            cmd.Parameters.Add("ROW_NO", OracleDbType.Int32).Value = rowNos
            cmd.Parameters.Add("HEADER_ID", OracleDbType.Int64).Value = headerIds
            cmd.Parameters.Add("CREATED_BY", OracleDbType.Varchar2).Value = users

            For Each col As ColumnDef In Columns
                Dim values(n - 1) As Object
                For i As Integer = 0 To n - 1
                    values(i) = ToDbValue(source.Rows(i)(col.Name))
                Next
                cmd.Parameters.Add(col.Name, col.DbType).Value = values
            Next

            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Function QueryMessages(conn As OracleConnection, batchId As Long, sql As String, errType As String) As List(Of ImportError)
        Dim list As New List(Of ImportError)
        Using cmd As New OracleCommand(sql, conn)
            cmd.BindByName = True
            cmd.Parameters.Add("BATCH_ID", OracleDbType.Int64).Value = batchId
            Using rd As OracleDataReader = cmd.ExecuteReader()
                While rd.Read()
                    list.Add(New ImportError(Convert.ToInt32(rd(0)), errType, rd.GetString(1)))
                End While
            End Using
        End Using
        Return list
    End Function

    Private Sub SetWorkStatus(conn As OracleConnection, batchId As Long, status As String)
        Using cmd As New OracleCommand("UPDATE WK_IMPORT SET STATUS = :STATUS WHERE BATCH_ID = :BATCH_ID", conn)
            cmd.BindByName = True
            cmd.Parameters.Add("STATUS", OracleDbType.Varchar2).Value = status
            cmd.Parameters.Add("BATCH_ID", OracleDbType.Int64).Value = batchId
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ''' <summary>
    ''' ワーク → 既存ヘッダ配下に追加。LINE_NO は既存の最大値の続きから振る。
    ''' (同じヘッダへの同時取込はヘッダ VERSION で弾くので、MAX+1 が衝突しても確定はしない)
    ''' </summary>
    Private Sub InsertDetail(conn As OracleConnection, batchId As Long, headerId As Long, userId As String)
        Dim cols As String = String.Join(", ", Columns.Select(Function(c) c.Name))
        Dim sql As String =
            "INSERT INTO PROD_DETAIL (HEADER_ID, LINE_NO, " & cols & ", DEL_FLG, BATCH_ID, CREATED_BY) " &
            "SELECT :HEADER_ID, " &
            "       (SELECT NVL(MAX(LINE_NO), 0) FROM PROD_DETAIL WHERE HEADER_ID = :HEADER_ID) + ROW_NUMBER() OVER (ORDER BY ROW_NO), " &
            "       " & cols & ", '0', :BATCH_ID, :CREATED_BY " &
            "  FROM WK_IMPORT WHERE BATCH_ID = :BATCH_ID"
        Using cmd As New OracleCommand(sql, conn)
            cmd.BindByName = True
            cmd.Parameters.Add("HEADER_ID", OracleDbType.Int64).Value = headerId
            cmd.Parameters.Add("BATCH_ID", OracleDbType.Int64).Value = batchId
            cmd.Parameters.Add("CREATED_BY", OracleDbType.Varchar2).Value = userId
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub DeleteWork(conn As OracleConnection, batchId As Long)
        Using cmd As New OracleCommand("DELETE FROM WK_IMPORT WHERE BATCH_ID = :BATCH_ID", conn)
            cmd.BindByName = True
            cmd.Parameters.Add("BATCH_ID", OracleDbType.Int64).Value = batchId
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ''' <summary>本処理のトランザクション終了後に、別トランザクションで履歴を確定させる</summary>
    Private Sub WriteHistory(conn As OracleConnection, batchId As Long, headerId As Long, fileName As String,
                             userId As String, errors As List(Of ImportError))
        If errors.Count = 0 Then Return
        Dim n As Integer = errors.Count
        Const sql As String =
            "INSERT INTO IMPORT_ERROR_HISTORY (ERR_ID, BATCH_ID, HEADER_ID, FILE_NAME, ROW_NO, ERR_TYPE, MESSAGE, CREATED_BY) " &
            "VALUES (SEQ_IMPORT_ERR.NEXTVAL, :BATCH_ID, :HEADER_ID, :FILE_NAME, :ROW_NO, :ERR_TYPE, :MESSAGE, :CREATED_BY)"

        Using tx As OracleTransaction = conn.BeginTransaction()
            Using cmd As New OracleCommand(sql, conn)
                cmd.BindByName = True
                cmd.ArrayBindCount = n
                Dim b(n - 1) As Object, h(n - 1) As Object, f(n - 1) As Object, r(n - 1) As Object
                Dim t(n - 1) As Object, m(n - 1) As Object, u(n - 1) As Object
                For i As Integer = 0 To n - 1
                    b(i) = batchId : h(i) = headerId : f(i) = fileName : u(i) = userId
                    r(i) = If(errors(i).RowNo.HasValue, CObj(errors(i).RowNo.Value), DBNull.Value)
                    t(i) = errors(i).ErrType
                    m(i) = Truncate(errors(i).Message, 1000)
                Next
                cmd.Parameters.Add("BATCH_ID", OracleDbType.Int64).Value = b
                cmd.Parameters.Add("HEADER_ID", OracleDbType.Int64).Value = h
                cmd.Parameters.Add("FILE_NAME", OracleDbType.Varchar2).Value = f
                cmd.Parameters.Add("ROW_NO", OracleDbType.Int32).Value = r
                cmd.Parameters.Add("ERR_TYPE", OracleDbType.Varchar2).Value = t
                cmd.Parameters.Add("MESSAGE", OracleDbType.Varchar2).Value = m
                cmd.Parameters.Add("CREATED_BY", OracleDbType.Varchar2).Value = u
                cmd.ExecuteNonQuery()
            End Using
            tx.Commit()
        End Using
    End Sub

    ' =================================================================
    '  ユーティリティ
    ' =================================================================

    Private Shared Function NextVal(conn As OracleConnection, seqName As String) As Long
        Using cmd As New OracleCommand("SELECT " & seqName & ".NEXTVAL FROM DUAL", conn)
            Return Convert.ToInt64(cmd.ExecuteScalar())
        End Using
    End Function

    ''' <summary>空文字・空白のみは NULL として扱う(Oracle は '' = NULL なので揃えておく)</summary>
    Private Shared Function ToDbValue(v As Object) As Object
        If v Is Nothing OrElse v Is DBNull.Value Then Return DBNull.Value
        Dim s As String = v.ToString().Trim()
        Return If(s.Length = 0, DBNull.Value, CObj(s))
    End Function

    Private Shared Function Truncate(s As String, max As Integer) As String
        If s Is Nothing Then Return ""
        Return If(s.Length <= max, s, s.Substring(0, max))
    End Function

End Class

' =====================================================================
'  結果クラス
' =====================================================================
Public Enum ImportStatus
    Succeeded       ' 本番登録完了
    HeldInWork      ' 論理削除データと一致 → ワークで保留
    Failed          ' 重複・競合・ヘッダ不正など
End Enum

Public Class ImportError
    Public ReadOnly Property RowNo As Integer?
    Public ReadOnly Property ErrType As String
    Public ReadOnly Property Message As String
    Public Sub New(rowNo As Integer?, errType As String, message As String)
        Me.RowNo = rowNo : Me.ErrType = errType : Me.Message = message
    End Sub
End Class

Public Class ImportResult
    Public ReadOnly Property Status As ImportStatus
    Public ReadOnly Property BatchId As Long
    Public ReadOnly Property Messages As IReadOnlyList(Of ImportError)

    Private Sub New(status As ImportStatus, batchId As Long, messages As IReadOnlyList(Of ImportError))
        Me.Status = status : Me.BatchId = batchId : Me.Messages = messages
    End Sub

    Public Shared Function Succeeded(batchId As Long) As ImportResult
        Return New ImportResult(ImportStatus.Succeeded, batchId, New List(Of ImportError)())
    End Function

    Public Shared Function HeldInWork(batchId As Long, messages As List(Of ImportError)) As ImportResult
        Return New ImportResult(ImportStatus.HeldInWork, batchId, messages)
    End Function

    Public Shared Function Failed(batchId As Long, errors As List(Of ImportError)) As ImportResult
        Return New ImportResult(ImportStatus.Failed, batchId, errors)
    End Function
End Class