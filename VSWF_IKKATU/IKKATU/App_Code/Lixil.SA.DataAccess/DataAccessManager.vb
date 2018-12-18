Option Explicit On
Option Strict On

'---10---+---20---+---30---+---40---+---50---+---60---+---70---+---80---+---90---+--100---+

''' <summary>
''' DBアクセスに関する管理クラス
''' </summary>
''' <remarks>
'''  初回アクセス時にのみ起動し、DB接続文字列を設定する。
''' </remarks>
Public NotInheritable Class DataAccessManager
    ''' <summary>
    ''' DB接続文字列の設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared connStr As String = _
        System.Configuration.ConfigurationManager. _
        ConnectionStrings("connectionString").ConnectionString

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub New()

    End Sub


    ''' <summary>
    ''' DB接続文字列の取得
    ''' </summary>
    ''' <value>DB接続文字列</value>
    ''' <returns>DB接続文字列</returns>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property Connection() As String
        Get
            Return connStr
        End Get
    End Property

End Class
