Imports System.IO
Imports System.Xml.Serialization

''' <summary>
''' My.Settings の代わりに使う自作設定クラス。
''' %APPDATA%\(会社名)\(アプリ名)\settings.xml (ローミングフォルダ) に保存する。
''' 使い方: AppSettings.Default.LastCsvFolder = "..." / AppSettings.Default.Save()
''' </summary>
Public Class AppSettings

    ' ===== 保存先 =====
    ' ※ 会社名・アプリ名は自分のものに変更する。バージョン番号を含まない固定パスなので、
    '    アプリを更新しても設定が引き継がれる(My.Settings の Upgrade() は不要)。
    Private Const CompanyFolder As String = "MyCompany"
    Private Const AppFolder As String = "MyApp"

    Public Shared ReadOnly Property FilePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                     CompanyFolder, AppFolder, "settings.xml")

    ' ===== 設定項目(ここに追加していく。初期値 = 既定値) =====
    ' XmlSerializer の対象は Public で読み書きできるプロパティのみ。
    Public Property LastCsvFolder As String = ""
    Public Property GridFontSize As Single = 9.0F

    ' ウィンドウ位置(Rectangle は余分な要素まで出力されるので int で持つ)
    Public Property WindowX As Integer = -1
    Public Property WindowY As Integer = -1
    Public Property WindowWidth As Integer = 0
    Public Property WindowHeight As Integer = 0
    Public Property WindowMaximized As Boolean = False

    ' リストは既定値を空にしておくこと
    ' (XmlSerializer は既存のリストに「追加」するため、既定要素があると重複する)
    Public Property RecentFiles As New List(Of String)

    ' ===== シングルトン =====
    Private Shared ReadOnly _lock As New Object()
    Private Shared _default As AppSettings

    ''' <summary>My.Settings.Default 相当。初回アクセス時にファイルから読み込む。</summary>
    Public Shared ReadOnly Property [Default] As AppSettings
        Get
            SyncLock _lock
                If _default Is Nothing Then _default = Load()
                Return _default
            End SyncLock
        End Get
    End Property

    ''' <summary>既定値に戻す(保存はしない。必要なら続けて Save を呼ぶ)。</summary>
    Public Shared Sub Reset()
        SyncLock _lock
            _default = New AppSettings()
        End SyncLock
    End Sub

    ''' <summary>ファイルを読み直す(設定ファイルを外部で編集した場合など)。</summary>
    Public Shared Sub Reload()
        SyncLock _lock
            _default = Load()
        End SyncLock
    End Sub

    ' ===== 読み込み =====
    Private Shared Function Load() As AppSettings
        If Not File.Exists(FilePath) Then Return New AppSettings()

        Try
            Using fs As New FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                Dim ser As New XmlSerializer(GetType(AppSettings))
                Return DirectCast(ser.Deserialize(fs), AppSettings)
            End Using
        Catch ex As Exception
            ' 壊れたファイルは退避して既定値で起動する(起動不能を防ぐ)
            Try
                File.Copy(FilePath, FilePath & ".broken_" & DateTime.Now.ToString("yyyyMMddHHmmss"), True)
            Catch
                ' 退避に失敗しても起動は続ける
            End Try
            Return New AppSettings()
        End Try
    End Function

    ' ===== 保存 =====
    ''' <summary>一時ファイルに書いてから置き換えるので、書き込み中に落ちても元ファイルは壊れない。</summary>
    Public Sub Save()
        SyncLock _lock
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath))

            Dim tmp As String = FilePath & ".tmp"
            Using fs As New FileStream(tmp, FileMode.Create, FileAccess.Write, FileShare.None)
                Dim ser As New XmlSerializer(GetType(AppSettings))
                ser.Serialize(fs, Me)
            End Using

            If File.Exists(FilePath) Then
                File.Replace(tmp, FilePath, Nothing)
            Else
                File.Move(tmp, FilePath)
            End If
        End SyncLock
    End Sub

End Class


' ======================================================================
' 使用例(フォーム側)
' ======================================================================
'Public Class MainForm
'
'    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
'        Dim s = AppSettings.Default
'
'        ' 保存位置が今の画面内にあるときだけ復元(モニター構成が変わった PC 対策)
'        If s.WindowWidth > 0 AndAlso s.WindowHeight > 0 Then
'            Dim r As New Rectangle(s.WindowX, s.WindowY, s.WindowWidth, s.WindowHeight)
'            If Screen.AllScreens.Any(Function(sc) sc.WorkingArea.IntersectsWith(r)) Then
'                Me.StartPosition = FormStartPosition.Manual
'                Me.Bounds = r
'            End If
'        End If
'        If s.WindowMaximized Then Me.WindowState = FormWindowState.Maximized
'    End Sub
'
'    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
'        Dim s = AppSettings.Default
'
'        ' 最大化・最小化中は RestoreBounds(通常時の位置)を保存する
'        Dim b As Rectangle = If(Me.WindowState = FormWindowState.Normal, Me.Bounds, Me.RestoreBounds)
'        s.WindowX = b.X : s.WindowY = b.Y
'        s.WindowWidth = b.Width : s.WindowHeight = b.Height
'        s.WindowMaximized = (Me.WindowState = FormWindowState.Maximized)
'
'        Try
'            s.Save()
'        Catch ex As Exception
'            ' 保存失敗で終了を止めない
'        End Try
'    End Sub
'
'End Class
