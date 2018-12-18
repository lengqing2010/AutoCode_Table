Option Explicit Off

Imports Microsoft.VisualBasic
Imports System.IO

Partial Class _Default
    Inherits System.Web.UI.Page




'******************************************************************'
'*【モジュール名】
'　    CHacTenCom.cls

'*【概要】
'　    発注中止日一括登録or店別情報一括変更登録

'*【設定】
'　　　@TransactionModeProperty
'           RequiresTransaction
'　　　@InstancingProperty
'           MultiUse
'　　　@ReferenceSettings
'           COM+ Services Type Library
'           Microsoft ActiveX Data Object 2.6 Library
'           COMDBUtil(DBAccess共通コンポーネント)
'           Microsoft Scripting Runtime

'******************************************************************'

'CSVファイル列番号 (この並びがCSVの並びになります。)
Private Enum enmHacColNo    '発注中止日用CSVデータ項目
    jigCD = 0               '該当事業コード
    SyuryouDate             '発注終了日
    janCD                   'JANコード
    Syobun                  '処分方法
    ChushiKbn               '発注中止理由区分
    DaiJanCd                '代替商品JANコード
End Enum

Private Enum enmTenColNo    '店別用CSVデータ項目
    TenCd = 0               '該当店コード
    HenkouDate              '変更日
    janCD                   'JANコード
    Genka                   '原価
    Baika                   '売価
    BaikaRiyuu              '売価変更理由
    Torihikisaki            '取引先コード
    Hacchuutani             '発注単位
    CenterKbn               'センター区分
End Enum
'2009.06.06 Add by Supattra T.(Tostem Thai)Start.
Private Enum enmModColNo
    TenCd = 0               '該当店コード
    HenkouDate              '変更日
    janCD                   'JANコード
    Genka                   '原価
    Baika                   '売価
    BaikaRiyuu              '売価変更理由
    Torihikisaki            '取引先コード
    Hacchuutani             '発注単位
    CenterKbn               'センター区分
    centerCD                'センターコード
    MoGenka                 '戻し原価
    MoBaika                 '戻し売価
    HattyuuTani             '戻し発注単位
    PopSize                 'POPサイズ
End Enum
'2009.06.06 Add by Supattra T.(Tostem Thai)End.

'********** 2005/08/31 高附 追加 **********↓
'CSVファイル列番号 (この並びがCSVの並びになります。)
Private Enum enmSyoColNo    '商品基本用CSVデータ項目
    janCD = 0               'JANコード
    SyouhinmeiKana          'カナ商品名(TA伝票用)
    SyouhinmeiKanji         '漢字商品名(SA/ﾌﾟﾗｲｽｶｰﾄﾞ用)
    ReceiptmeiKana          'レシート名(ｶﾅ)
    ReceiptmeiKanji         'レシート名(漢字)
    SyouhihinmeiHansoku     '商品名(販促用)
    
    MakerMei                'メーカー名
    KikakuSize              '規格/サイズ
    HinbanKataban           '品番/型番
    IroGara                 '色/柄
    Taste                   'テースト
    Seisantimei             '生産地名
    Tokutyou                '特徴
    SyouhinCopy             '商品コピー
    TuujyouKakaku           '通常価格
    MakerKakaku             'メーカー希望小売価格
    POPLayout1              'POPレイアウト１
    POPLayout2              'POPレイアウト２
    PopSize                 'POPサイズ
    
    AtukaiPtn               '扱いパターン
    HyoujyunBaika           '標準売価
    SikiriGenka             '仕切原価
    VhJyuyouCd              'VH需要コード
    RdJyuyouCd              'RD需要コード
    VcsJyuyouCd             'VCS需要コード
    SvhJyuyouCd             'SVH需要コード
End Enum

Private Enum enmBunColNo    '分類用CSVデータ項目
    HenkouDate = 0          '変更日
    janCD                   'JANコード
    Bunrui1                 '第1分類コード
    Bunrui2                 '第2分類コード
    Bunrui3                 '第3分類コード
End Enum
'********** 2005/08/31 高附 追加 **********↑

'CSVファイル フィールド名
Private Const FLD_HACCHUU = "該当事業コード,発注終了日,JANコード,処分方法,発注中止理由区分,代替商品コード"
'*M 2014/01/22; 李松涛(大連) ;　P-46854; 税制改正に伴う対応(商品選定WF・新商品情報);　START
Private Const FLD_TENBETU = "該当店コード,変更日,JANコード,原価,売価,売価変更理由,取引先,発注単位,センター区分,センターコード,戻し原価,戻し売価,戻し発注単位,POPサイズ"
'*M 2014/01/22; 李松涛(大連) ;　P-46854; 税制改正に伴う対応(商品選定WF・新商品情報);　END
Private Const FLD_SYOUHIN = "JANコード,カナ商品名(TA伝票用),漢字商品名(SA/ﾌﾟﾗｲｽｶｰﾄﾞ用),レシート名(ｶﾅ),レシート名(漢字),商品名(販促用),メーカー名,規格/サイズ,品番/型番,色/柄,テースト,生産地名,特徴,商品コピー,通常価格,ﾒｰｶｰ希望小売価格,POPレイアウト１,POPレイアウト２,POPサイズ,扱いパターン,標準売価,仕切原価,VH需要コード,R&D需要コード,VCS需要コード,SVH需要コード"
Private Const FLD_BUNRUI = "変更日,JANコード,第1分類コード,第2分類コード,第3分類コード"
'2009.06.05 Add by Supattra T.(Tostem Thai)Start.
'Private Const FLD_MODOSI = "該当店コード,変更日,JANコード,原価,売価,売価変更理由,取引先,発注単位,センター区分,戻し原価,戻し売価,戻し発注単位,POPサイズ"
'*M 2014/01/22; 李松涛(大連) ;　P-46854; 税制改正に伴う対応(商品選定WF・新商品情報);　START
'Private Const FLD_MODOSI = "該当店コード,変更日,JANコード,原価,売価,売価変更理由,取引先,発注単位,センター区分,センターコード,戻し原価,戻し売価,戻し発注単位,POPサイズ"
Private Const FLD_MODOSI = "該当店コード,変更日,JANコード,原価,売価(税別),売価変更理由,取引先,発注単位,センター区分,センターコード,戻し原価,戻し売価,戻し発注単位,POPサイズ"
'*M 2014/01/22; 李松涛(大連) ;　P-46854; 税制改正に伴う対応(商品選定WF・新商品情報);　END
'2009.06.05 Add by Supattra T.(Tostem Thai)End.

'店・事業コード（全店,全事業,本部）
Private Const ALL_TEN_CD = "9999"                 '全店
Private Const ALL_JIG_CD = "99"                   '全事業
Private Const HON_TEN_CD = "0000"                 '本部

'呼び出し側へ戻す為の戻り値
Private Const RET_FLD_HACH As Long = 1                  '発注中止日
Private Const RET_FLD_TENB As Long = 2                  '店別情報
Private Const RET_FLD_SYO As Long = 3                   '商品基本情報
Private Const RET_FLD_BUN As Long = 4                   '分類情報
Private Const RET_NO_DATA As Long = 0                   'データ無し
'2009.06.05 Add by Supattra T.(Tostem Thai)Start.
Private Const RET_FLD_MODO As Long = 5
'2009.06.05 Add by Supattra T.(Tostem Thai)End.

'事業別需要コード
Private Const JYU_VH As String = "01"                   'VH需要コード
Private Const JYU_RD As String = "02"                   'RD需要コード
Private Const JYU_VCS As String = "03"                  'VCS需要コード
Private Const JYU_SVH As String = "06"                  'SVH需要コード


'KET区分
Private Const KEYCD_BAIKARIYUU = "025"            '売価変更理由

'センターコード
'Private Const CENTERCD_KODAMA = "0151"            '(児玉センター)
Private Const CENTERCD_KODAMA = "0810"            '(大生郷備蓄センター)

'商品選定WFのテーブル名
Private Const TBL_BAI_HEN As String = "t_syousen_baika_henkou"      '商品選定DB(売価変更情報)
Private Const TBL_SII_HEN As String = "t_syousen_siire_henkou"      '商品選定DB(仕入変更情報)
Private Const TBL_CEN_HEN As String = "t_syousen_center_cd_henkou"  '商品選定DB(センター変更情報)
'********** 2008/05/27 ラン峰 追加 **********↓
Private Const TBL_JYO_TEN As String = "t_syousen_jyogai_tenpo"      '商品選定WF(変更除外店舗)
'********** 2008/05/27 ラン峰 追加 **********↑

' 各フィールドの桁数
Private Const LENGTH_JANCODE1 As Long = 8                   'JANコード桁数１
Private Const LENGTH_JANCODE2 As Long = 13                  'JANコード桁数２
Private Const LENGTH_TENCODE As Long = 4                    '店コードの桁数
Private Const LENGTH_CENTERCODE As Long = 4                 'センターコードの桁数
Private Const LENGTH_BAIKA As Long = 7                      '売価の桁数
Private Const LENGTH_HATTYU_TANI As Long = 5                '発注単位桁数
Private Const LENGTH_GENKA_SEISU As Long = 7                '原価(整数部)桁数
Private Const LENGTH_GENKA_SYOUSUU As Long = 2              '原価(小数部)桁数
Private Const LENGTH_HENKOU_DATE As Long = 8                '変更日の桁数
Private Const LENGTH_BUNRUI As Long = 6                     '分類コードの桁数
Private Const LENGTH_JYUYOU As Long = 6                     '需要コードの桁数
Private Const LENGTH_SYOKANA As Long = 40                   'カナ商品名(TA伝票用)の桁数
Private Const LENGTH_SYOKANJI As Long = 30                  '漢字商品名(SA/ﾌﾟﾗｲｽｶｰﾄﾞ用)の桁数
Private Const LENGTH_RESKANA As Long = 14                   'レシート名(ｶﾅ)の桁数
Private Const LENGTH_RESKANJI As Long = 14                  'レシート名(漢字)の桁数
Private Const LENGTH_SYOHANSOKU As Long = 20                '商品名(販促用)の桁数
Private Const LENGTH_ATUPTN As Long = 3                     '扱いパターンの桁数

'エラーの原因を示す定数
Private Const PARAM_EMPTY = 10              ' パラメータが入力されていない
Private Const PARAM_NOTDISIT = 20           ' パラメータが非数値である
Private Const PARAM_NOTFINDMASTER = 30      ' マスタに存在しないデータ
Private Const PARAM_WRONGCHECKDIGIT = 40    ' チェックディジット不正
Private Const PARAM_WRONGDATE = 50          ' 日付不正
Private Const PARAM_SHORTDISIT = 60         ' 桁数不足
Private Const PARAM_NOTEXISTDB = 70         ' データベース(マスタ)に存在しない
Private Const PARAM_SAMECODEDB = 80         ' すでにデータベース上に登録されている
Private Const PARAM_TORIHIKISTOP = 90       ' 取引先との取引は停止されている
Private Const PARAM_REPDATEBAIKA = 100      ' 統合DB(売価変更情報)に既に同日変更日のデータが存在する
Private Const PARAM_REPDATESIIRE = 110      ' 統合DB(仕入変更情報)に既に同日変更日のデータが存在する
Private Const PARAM_REPDATETENPO = 120      ' 統合DB(店舗センターコード変更情報)に既に同日変更日のデータが存在する
'********** 2005/04/25 浅野 追加 **********↓
Private Const PARAM_NOTDEFINED = 130        ' その値はパラメータの定義から外れている
Private Const PARAM_EARLYCANCELDAY = 140    ' 発注終了日が早すぎる（開始日より前)
'********** 2005/04/25 浅野 追加 **********↑
'********** 2005/05/07 北澤 **********↓
Private Const PARAM_ZERO = 150              ' 金額類が正の値で無い
Private Const PARAM_MISENTAKU = 160         ' 発注中止理由が00(未選択)
'********** 2005/05/07 北澤 **********↑
'********** 2005/06/16 北澤 **********↓
Private Const PARAM_BAIKARIYUU_HISSU = 170  ' 店コード=0000の場合、売価が入力されていたら、売価変更理由は入力必須
Private Const PARAM_BAIKARIYUU_FUKA = 180   ' 売価が未入力の場合、又は店コード<>0000の場合、売価変更理由は入力不可
Private Const PARAM_NONARROW = 190          ' パラメーターが非半角文字である
Private Const PARAM_BUNRUIDB = 200          ' 分類情報が存在しない
Private Const PARAM_URIBADB = 210           ' 売場体系情報が存在しない
Private Const PARAM_ZENKAKU = 220           ' 全角文字列以外で入力されている
Private Const PARAM_SAMEKEYDB = 230         ' 企画外の内部コードと重複している
Private Const PARAM_HANEISUU = 240          ' 半角英数字以外の文字が入力されている。
Private Const PARAM_KINSIMOJI = 250         ' 半角の禁止文字が入力されている。
'2014/05/02 add wadak7
Private Const PARAM_KINSIMOJI_RECEIPT = 255 ' 禁止文字が入力されている(レシート名(漢字))。
'********** 2005/06/16 北澤 **********↑
Private Const PARAM_GROUP = 260             ' グループ品が入力されている。
Private Const PARAM_KINSI = 270             ' 入力できない。
'********** 2012/08/03 楊双（大連）入力チェックの追加 **********↓
Private Const PARAM_HONBU_NAI = 280         ' センター区分が2：PDCの場合、0000：本部のデータは必須入力です
Private Const PARAM_HATTYUU_TANI = 290      ' センター区分が2：PDCの発注単位は全て同じ値ではない
Private Const PARAM_GENKA = 300             ' センター区分が2：PDCの原価は全て同じ値ではない
Private Const PARAM_GENKA_HATTYUUTANI = 310 ' 更新店舗以外にPDC区分になっている店舗があります。
'********** 2012/08/03 楊双（大連）入力チェックの追加 **********↑
'エラーを表現するモジュール変数
    Private Structure typ_Error_Structure
        Dim intErrorLine As Integer
        Dim intErrorField As Integer
        Dim intErrorType As Integer
        Dim strErrorParameter As String
    End Structure

' エラー構造体配列、発生順にエラー情報が格納される
Private typErrorArray() As typ_Error_Structure
'********** 2012/08/03 楊双（大連）入力チェックの追加 **********↓
Private typErrorArrayPDC() As typ_Error_Structure
'********** 2012/08/03 楊双（大連）入力チェックの追加 **********↑

' 発生したエラーの数
Private intErrorCount As Integer
'********** 2012/08/03 楊双（大連）入力チェックの追加 **********↓
Private intErrorCountPDC As Integer
'********** 2012/08/03 楊双（大連）入力チェックの追加 **********↑

'モジュール変数
Private pstrTantousyaId As String   '担当者ID
Private pstrFilePath    As String   'ファイルパス
Private pstrErrMsg As String        'エラーメッセージ
Private pintSyouhizei As Integer    '消費税率(%)

'*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************
Private plngSaibanNo2 As Long       '採番№
Private plngDataCnt As Long         'データ数
Private pstrFileSyurui As String    'ファイル種類
'*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************

Private pstrKikakuSyurui As String      '企画種類
Private pstrKikakuNenrenban As String   '企画年連番
Private pstrKikakuEdaban As String      '企画枝版

Private pflgBaikaToukbn As Boolean   '売価変更情報登録区分フラグ（1:追加の先付情報あり、0:なし）
Private pflgSiireToukbn As Boolean   '仕入変更情報登録区分フラグ（1:追加の先付情報あり、0:なし）
Private pflgCenterToukbn As Boolean  'センターコード変更情報登録区分フラグ（1:追加の先付情報あり、0:なし）
'********** 2005/04/25 黒沢 追加 **********↓
Private pflgJigJyoToukbn As Boolean  '事業別情報登録区分フラグ（1:追加の先付情報あり、0:なし）
'********** 2005/04/25 黒沢 追加 **********↑

'********** 2005/05/11 土田 追加 **********↓
'DB参照エラーチェック用の切断レコードセット

'店別、発注中止両方使用
Private pChkSameJANCodeRS As Object     '重複JANコードチェック用切断レコードセット
'********** 2012/08/03 楊双（大連）入力チェックの追加 **********↓
Private pChkSameGenkaHattyuuTaniRS As Object  '発注単位、原価が一致チェック用切断レコードセット
Private pChkHonbuGenkaHattyuuTaniRS As Object  '発注単位、原価が一致チェック用切断レコードセット（本部）
'********** 2012/08/03 楊双（大連）入力チェックの追加 **********↑
Private pChkvk1JANCodeRS As Object      '統合マスタ登録済みJANコードチェック用切断レコードセット

'分類情報使用
Private pChkBunruiCodeRS As Object      '分類マスタKEYチェック用切断レコードセット
Private pCsvBunrui1CodeList As String   'CSVファイル内第1分類コードのリスト(第1分類コードの重複無し)

'需要情報使用
Private pChkJyuyouCodeRS As Object      '売場体系表示用マスタ需要コードチェック用切断レコードセット
Private pCsvJyuyouCodeList As String    'CSVファイル内需要コードのリスト(第1分類コードの重複無し)

'扱いパターン情報使用
Private pChkAtukaiPtnCdRS As Object      '扱いパターンコードコードチェック用切断レコードセット

'利益管理情報使用
'Private pChktsrkJANCodeRS As Object
'WF管理情報使用
'Private pChktswkJANCodeRS As Object

'********** 2005/05/18 土田 発注中止でも使用する為、変数名変更 **********↓
Private pCsvJanCodeList As String      'CSVファイル内Janコードのリスト(Janコードの重複無し)
'********** 2005/05/18 土田 変数名変更 **********↑
'2009.06.08 Add by Supattra T.(Tostem Thai)Start.
Private pCsvHenkouDate As String
Private pCsvhenkouDateKeizoku As String
Private pTTenCode As String
Private pTJanCode As String
'2009.06.08 Add by Supattra T.(Tostem Thai)End.

'店別用
Private pTChkJigyouRS As Object         '該当事業コードチェック用切断レコードセット
'2009.07.10 Add by Supattra T.(Tostem Thai)Start.
Private pTChkkikakuRS As Object
Private pTChkCenterCdRS As Object       'センターコードチェック用切断レコードセット

'2009.07.10 Add by Supattra T.(Tostem Thai)End
Private pTChkBaikaHenkouRS As Object    '売価変更理由チェック用切断レコードセット
Private pTChkTorihikisakiRS As Object   '取引先チェック用切断レコードセット
Private pTChkCenterRS As Object         'センター区分チェック用切断レコードセット

Private pTTenCodeList As String         '展開店コードのリスト(全事業、全店）

'********** 2005/05/19 土田 追加 **********↓
Private pTKeyCDList As String      '統合マスタ登録済みJANコード紐付け主キーのリスト(主キー重複無し)
'********** 2005/05/19 土田 追加 **********↑

'********** 2005/05/18 土田 追加 **********↓
Private pstrMinDate As String      'CSVファイル内最小日付
Private pstrMaxDate As String      'CSVファイル内最大日付
'********** 2005/05/18 土田 追加 **********↑

'発注中止用
Private pHChkJigyouRS As Object         '該当事業コードチェック用切断レコードセット
Private pHChkHacchuRS As Object         '発注中止理由区分チェック用切断レコードセット

'********** 2005/05/18 土田 追加 **********↓
Private pHChkHattyuuKaisiDateRS As Object   '発注開始日チェック用切断レコードセット
Private pHChkJHTYKaisiDateRS As Object      '発注開始日を事業別情報テーブルから取得用切断レコードセット

Private pHKeyCDList As String      '統合マスタ登録済みJANコード紐付け主キーのリスト(主キー重複無し)
'********** 2005/05/18 土田 追加 **********↑

'********** 2005/05/11 土田 追加 **********↑

'********** 2005/05/13 田中 追加 **********↓
'DB参照エラーチェック用の切断レコードセット（追加）
Private pTChkMstBaikaRS As Object       '統合マスタ・売価変更情報の変更日チェック用切断レコードセット
Private pTChkMstSiireRS As Object       '統合マスタ・仕入変更情報の変更日チェック用切断レコードセット
Private pTChkMstCenterRS As Object       'センターコードチェック用切断レコードセット
'********** 2005/05/13 田中 追加 **********↑

    '********** 2005/06/15 北澤 **********↓
    Dim arrTourokuKbn2(0, 0) As String      '登録区分=2の店コードの保管用配列
    Private Const conColBaika As Integer = 0 '上記配列の2次元目の列番号(売価変更情報)
    Private Const conColSiire As Integer = 1 '上記配列の2次元目の列番号(仕入変更情報)
    Private Const conColCenter As Integer = 2 '上記配列の2次元目の列番号(センター変更情報)
'********** 2005/06/15 北澤 **********↓


'定数
Private Const conLngStartRow As Integer = 2  'CSVファイルの1行目のデータ位置

'CSVデータ格納配列 (件数)(項目番号)
Private paryCsvData() As Object

'********** 2018/12/17 李松涛 **********↓
    Private pStrRenban As String


    'Property Let letFilePath(strPath As String)                     'ファイルパス
    '   pstrFilePath = strPath
    '    End Property
    Public Property letFilePath() As String
        Get
            Return pstrFilePath
        End Get
        Set(ByVal value As String)
            pstrFilePath = value
        End Set
    End Property


    'Property Let letKikakuSyurui(strKikakuSyurui As String)         '企画種類
    '   pstrKikakuSyurui = strKikakuSyurui
    '    End Property


    Public Property letKikakuSyurui() As String
        Get
            Return pstrKikakuSyurui
        End Get
        Set(ByVal value As String)
            pstrKikakuSyurui = value
        End Set
    End Property


    'Property Let letKikakuNenrenban(strKikakuNenrenban As String)   '企画年連番 (YYYY + 連番)
    '   pstrKikakuNenrenban = strKikakuNenrenban
    '    End Property

    Public Property letKikakuNenrenban() As String
        Get
            Return pstrKikakuNenrenban
        End Get
        Set(ByVal value As String)
            pstrKikakuNenrenban = value
        End Set
    End Property


    'Property Let letKikakuEdaban(strKikakuEdaban As String)         '企画枝版
    '   pstrKikakuEdaban = strKikakuEdaban
    '    End Property

    Public Property letKikakuEdaban() As String
        Get
            Return pstrKikakuEdaban
        End Get
        Set(ByVal value As String)
            pstrKikakuEdaban = value
        End Set
    End Property


    'Property Let letTantousyaId(strTantousyaId As String)           '担当者ID
    '   pstrTantousyaId = strTantousyaId
    '    End Property

    Public Property letTantousyaId() As String
        Get
            Return pstrTantousyaId
        End Get
        Set(ByVal value As String)
            pstrTantousyaId = value
        End Set
    End Property
    'Property Let letSaibanNo2(lngSaibanNo2 As Long)                 'ログ記載用_採番№  2018/11/19 山田 予期せぬエラー対応
    '   plngSaibanNo2 = lngSaibanNo2
    'End Property
    Public Property letSaibanNo2() As String
        Get
            Return plngSaibanNo2
        End Get
        Set(ByVal value As String)
            plngSaibanNo2 = value
        End Set
    End Property
'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    FncCSVIkkatsuTouroku

'*【処理概要】
'*?    CSVファイルのデータを商品選定WFのDBに取り込む

'*【パラメータ】
'*P    strErrMsg;        string; エラーメッセージ;

'*【戻り値】
'*R    int; RET_FLD_HACH:   発注中止日用データ;
'*R         RET_FLD_TENB:   店別情報用データ;
'*R         RET_NO_DATA:    データ無し;
'*R         上記以外:        エラー;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/18; 黒沢(ユニシステム株式会社); 新規作成;
'*M    2005/08/31; 高附(IC)                 ; 分類情報一括変更追加;
'*M    2009/06/06; Supattra T.(Tostem Thai) ; Add case CSVファイル;
'*< Comment End >**************************************************'
    Public Function FncCSVIkkatsuTouroku(ByRef strErrMsg _
                                        , ByRef aryCSVData As String()) As Integer
        Dim intRecNum
        Dim intRet

        plngDataCnt = 0                 '2018/11/19 山田 予期せぬエラー対応
        pstrFileSyurui = ""             '2018/11/19 山田 予期せぬエラー対応
        intRecNum = 0
        FncCSVIkkatsuTouroku = -1

        '李 一括登録ログ追加 2018/12/17
        Call SetIkkatuLog("1", "CSVファイルのデータを商品選定WFのDBに取り込む", "FncCSVIkkatsuTouroku", "", "", pstrTantousyaId, "INS")

        '*******************************************************************************************************************************
        'ログ出力（①連番、②処理ファイル、③データ数、④処理メソッド、⑤処理詳細、⑥ユーザーID、⑦データ更新区分）
        Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "FncCSVToArray_処理開始前", pstrTantousyaId, "1")
        '*******************************************************************************************************************************

        'CSVファイルのデータを取得して、配列に入れる。
        'ここで、発注中止日or店別情報に分ける。
        intRet = FncCSVToArray(strErrMsg)

        'CSVデータをASPに返す
        aryCSVData = paryCsvData

        If intRet = RET_FLD_HACH Then       '発注中止日一括登録の場合
            '*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
            pstrFileSyurui = "発注中止日"
            'ログ出力（①連番、②処理ファイル、③データ数、④処理メソッド、⑤処理詳細、⑥ユーザーID、⑦データ更新区分）
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "発注中止日一括登録処理_処理開始前", pstrTantousyaId, "2")
            '*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

            '********** 2005/04/26 黒沢 変更 **********↓
            '発注中止日一括登録処理の追加
            If FncHacTouroku(strErrMsg) = True Then
                FncCSVIkkatsuTouroku = RET_FLD_HACH
            Else
                FncCSVIkkatsuTouroku = -1
            End If
            '********** 2005/04/26 黒沢 変更 **********↑

            '*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
            'ログ出力（①連番、②処理ファイル、③データ数、④処理メソッド、⑤処理詳細、⑥ユーザーID、⑦データ更新区分）
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "発注中止日一括登録処理_処理終了", pstrTantousyaId, "3")
            '*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

        ElseIf intRet = RET_FLD_TENB Then   '店別情報一括変更登録の場合
            '*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
            pstrFileSyurui = "店別情報"
            'ログ出力（①連番、②処理ファイル、③データ数、④処理メソッド、⑤処理詳細、⑥ユーザーID、⑦データ更新区分）
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "店別情報一括登録処理_処理開始前", pstrTantousyaId, "2")
            '*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

            If FncTenTouroku(strErrMsg) = True Then
                FncCSVIkkatsuTouroku = RET_FLD_TENB
            Else
                FncCSVIkkatsuTouroku = -1
            End If

            '*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
            'ログ出力（①連番、②処理ファイル、③データ数、④処理メソッド、⑤処理詳細、⑥ユーザーID、⑦データ更新区分）
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "店別情報一括登録処理_処理終了", pstrTantousyaId, "3")
            '*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

            '2009.06.06 Add by Supattra T.(Tostem Thai)Start.
        ElseIf intRet = RET_FLD_MODO Then   'Add 店別情報一括変更登録票
            '*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
            pstrFileSyurui = "戻し価格"
            'ログ出力（①連番、②処理ファイル、③データ数、④処理メソッド、⑤処理詳細、⑥ユーザーID、⑦データ更新区分）
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "戻し価格一括登録処理_処理開始前", pstrTantousyaId, "2")
            '*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

            If FncModosiKakaku(strErrMsg) = True Then
                FncCSVIkkatsuTouroku = RET_FLD_MODO
                aryCSVData = paryCsvData
            Else
                FncCSVIkkatsuTouroku = -1
            End If

            '*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
            'ログ出力（①連番、②処理ファイル、③データ数、④処理メソッド、⑤処理詳細、⑥ユーザーID、⑦データ更新区分）
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "戻し価格一括登録処理_処理終了", pstrTantousyaId, "3")
            '*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
            '2009.06.06 Add by Supattra T.(Tostem Thai)End.
            '********** 2005/08/31 高附 追加 **********↓
        ElseIf intRet = RET_FLD_SYO Then   '商品基本情報一括変更登録の場合
            '*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
            pstrFileSyurui = "商品基本情報"
            'ログ出力（①連番、②処理ファイル、③データ数、④処理メソッド、⑤処理詳細、⑥ユーザーID、⑦データ更新区分）
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "商品基本情報一括登録処理_処理開始前", pstrTantousyaId, "2")
            '*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

            If FncSyoTouroku(strErrMsg) = True Then
                FncCSVIkkatsuTouroku = RET_FLD_SYO
            Else
                FncCSVIkkatsuTouroku = -1
            End If

            '*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
            'ログ出力（①連番、②処理ファイル、③データ数、④処理メソッド、⑤処理詳細、⑥ユーザーID、⑦データ更新区分）
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "商品基本情報一括登録処理_処理終了", pstrTantousyaId, "3")
            '*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

        ElseIf intRet = RET_FLD_BUN Then   '分類情報一括変更登録の場合
            '*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
            pstrFileSyurui = "分類情報"
            'ログ出力（①連番、②処理ファイル、③データ数、④処理メソッド、⑤処理詳細、⑥ユーザーID、⑦データ更新区分）
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "分類情報一括登録処理_処理開始前", pstrTantousyaId, "2")
            '*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

            If FncBunTouroku(strErrMsg) = True Then
                FncCSVIkkatsuTouroku = RET_FLD_BUN
            Else
                FncCSVIkkatsuTouroku = -1
            End If

            '*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
            'ログ出力（①連番、②処理ファイル、③データ数、④処理メソッド、⑤処理詳細、⑥ユーザーID、⑦データ更新区分）
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "分類情報一括登録処理_処理終了", pstrTantousyaId, "3")
            '*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

            '********** 2005/08/31 高附 追加 **********↑
        ElseIf intRet = RET_NO_DATA Then              'データが１件も無い場合
            '*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
            pstrFileSyurui = "データなし"
            'ログ出力（①連番、②処理ファイル、③データ数、④処理メソッド、⑤処理詳細、⑥ユーザーID、⑦データ更新区分）
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "データが１件もナシ判定", pstrTantousyaId, "3")
            '*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

            FncCSVIkkatsuTouroku = RET_NO_DATA
        ElseIf intRet = -2 Then             '「ファイルを開けません。」エラーの場合
            '*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
            pstrFileSyurui = "OPENエラー"
            'ログ出力（①連番、②処理ファイル、③データ数、④処理メソッド、⑤処理詳細、⑥ユーザーID、⑦データ更新区分）
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "ファイルを開けません判定", pstrTantousyaId, "3")
            '*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

            FncCSVIkkatsuTouroku = -2
        Else
            '*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
            pstrFileSyurui = "その他エラー"
            'ログ出力（①連番、②処理ファイル、③データ数、④処理メソッド、⑤処理詳細、⑥ユーザーID、⑦データ更新区分）
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "その他エラー判定", pstrTantousyaId, "3")
            '*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

            FncCSVIkkatsuTouroku = -1       'それ以外（エラー）の場合
        End If

        Call SetIkkatuLog("1", "", "", "", "OK", "", "UPD")

        '*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
        'ログ出力（①連番、②処理ファイル、③データ数、④処理メソッド、⑤処理詳細、⑥ユーザーID、⑦データ更新区分）
        Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "FncCSVIkkatsuTouroku_処理終了", pstrTantousyaId, "4")
        '*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    End Function

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    GetCsvLineData

'*【処理概要】
'*?    CSVファイルのデータを1次元配列に転換する
'*【パラメータ】
'*P    arrayData;      配列; CSVファイルのデータ;

'*【戻り値】
'*R    なし
'******************************************************************'
'*【変更履歴】
'*M    2008/08/26; 張楠(大連); 　　　　　新規作成;
'*< Comment End >**************************************************'
Public Sub GetCsvLineData(ByVal arrayData As Object, ByRef arrayCsvLineData As Object)
    Dim aryTest() As Object
    Dim strTemp As String
    strTemp = ""
    Dim Count As Integer
    Dim countCols As Integer
    For Count = 0 To UBound(arrayData)
    ReDim Preserve aryTest(Count)
        For countCols = 0 To UBound(arrayData(0))
            If countCols = 0 Then
                strTemp = arrayData(Count)(countCols)
            Else
                strTemp = strTemp & "," & Trim(arrayData(Count)(countCols))
            End If
        Next
        countCols = 0
        aryTest(Count) = strTemp
        strTemp = ""
    Next
    arrayCsvLineData = aryTest
End Sub

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    FncCSVToArray

'*【処理概要】
'*?    CSVファイルのデータを配列に入れる。
'*【パラメータ】
'*P    strErrMsg;      string; エラーメッセージ;

'*【戻り値】
'*R    int; RET_FLD_HACH:   発注中止日用データ;
'*R         RET_FLD_TENB:   店別情報用データ;
'*R         RET_NO_DATA:    データ無し;
'*R         -1:             エラー;
'*R         -2:             エラー(ファイルが開けません。);

'******************************************************************'
'*【変更履歴】
'*M    2005/04/18; 黒沢(ユニシステム株式会社); 新規作成;
'*M    2009/06/06; Supattra T.(Tostem Thai);  Add case CSVファイル ;
'*M    2009/07/31; 和田                     ;  まる得のチェック追加 ;
'*M    2011/02/10; 和田                     ;  センターコード追加 ;
'*< Comment End >**************************************************'
Private Function FncCSVToArray(ByRef strErrMsg) As Integer
    Dim intRecNum, intIndex
    Dim lngCount As Long
    Dim intFile
    Dim varData As Object      'CSVデータ(1行)
    Dim conIntKoumoku           'インデックス番号の最大値（要素数確認用）

    FncCSVToArray = -1
    
    intFile = FreeFile          '使用可能なファイルナンバーを取得（ファイル操作に必要）

    'CSVファイルからデータを取得する。
        On Error Resume Next

    'ファイルを開く
        'Open pstrFilePath For Input Access Read Lock Read As #intFile
        'If Err.Number <> 0 Then
        '    FncCSVToArray = -2
        '    strErrMsg = "ファイルが開けません。(" & Err.Description & ")"
        '    Exit Function
        '    End If



        On Error GoTo 0

        'CSVファイルのデータを配列paryCsvDataに格納
        lngCount = 0
        intRecNum = 0

        For Each varData In File.ReadAllLines(pstrFilePath)
            '配列の拡張
            ReDim Preserve paryCsvData(intRecNum)

            'フィールド部分では、どちらの一括登録かを判別する。（発注中止日or店別情報）
            If lngCount = 0 Then
                If pstrKikakuSyurui = "012" Or pstrKikakuSyurui = "015" Then
                    '2009.06.06 Add by Supattra  T.(Tostem Thai)Start.
                    If varData = FLD_MODOSI Then       '店別情報
                        FncCSVToArray = RET_FLD_MODO
                        conIntKoumoku = enmModColNo.PopSize
                    Else
                        FncCSVToArray = -1
                        strErrMsg = "一括登録用(まる得)のCSVファイルではありません。"
                        Exit Function
                    End If
                Else
                    '2009.06.06 Add by Supattra  T.(Tostem Thai)End.
                    If varData = FLD_HACCHUU Then           '発注中止日
                        FncCSVToArray = RET_FLD_HACH
                        conIntKoumoku = enmHacColNo.DaiJanCd
                    ElseIf varData = FLD_TENBETU Then  '店別情報
                        'FncCSVToArray = RET_FLD_TENB
                        'conIntKoumoku = enmTenColNo.CenterKbn

                        FncCSVToArray = -1
                        strErrMsg = "旧バージョンの登録票です。最新の登録票をダウンロードしてください。"
                        Exit Function

                    ElseIf varData = FLD_SYOUHIN Then       '商品基本情報
                        FncCSVToArray = RET_FLD_SYO
                        conIntKoumoku = enmSyoColNo.SvhJyuyouCd
                    ElseIf varData = FLD_BUNRUI Then        '分類情報
                        FncCSVToArray = RET_FLD_BUN
                        conIntKoumoku = enmBunColNo.Bunrui3
                    ElseIf varData = FLD_MODOSI Then        '戻し原価
                        FncCSVToArray = RET_FLD_MODO
                        conIntKoumoku = enmModColNo.PopSize
                    Else                                    'エラー
                        FncCSVToArray = -1
                        strErrMsg = "一括登録用のCSVファイルではありません。"

                        Exit Function
                    End If
                End If

                'データ部分は、カンマで分割して配列に格納
            ElseIf lngCount >= 1 Then
                paryCsvData(intRecNum) = Split(varData, ",")

                'エラーチェック
                If IsArray(paryCsvData(intRecNum)) = False Then       'ファイル形式の確認
                    FncCSVToArray = -1
                    strErrMsg = "CSVファイルの形式が不正です。"

                    Exit Function
                End If
                If UBound(paryCsvData(intRecNum)) <> conIntKoumoku Then  '要素数の確認
                    FncCSVToArray = -1
                    strErrMsg = "CSVファイルの項目数に過不足があります。"
                    Exit Function
                End If

                '各項目の両端の空白を取り除く
                intIndex = 0
                For intIndex = 0 To conIntKoumoku Step 1
                    paryCsvData(intRecNum)(intIndex) = Trim(paryCsvData(intRecNum)(intIndex))
                Next

                intRecNum = intRecNum + 1
            End If
            varData = ""
            lngCount = lngCount + 1
        Next
        Do While Not EOF(intFile)           'ファイルの終端までループを繰り返します。
        
        Loop


        '*** ↓2018/11/20 山田 予期せぬエラー対応 ************************************************************************
        'データ件数を取得
        plngDataCnt = lngCount
        '*** ↑2018/11/20 山田 予期せぬエラー対応 ************************************************************************

        '該当データが無い場合
        If intRecNum = 0 Then
            FncCSVToArray = RET_NO_DATA
        End If

        Dim blnNotAllBlank As Boolean

        '********** 2005/05/31 北澤 **********↓
        '店別CSVでキー以外が全てブランクだったら該当データ無しメッセージを表示する(事業別は必須項目があるのであり得ない)
        If FncCSVToArray = RET_FLD_TENB Then

            blnNotAllBlank = False
            If intRecNum > 0 Then
                For lngCount = 0 To (intRecNum - 1)
                    For intIndex = 3 To conIntKoumoku
                        If paryCsvData(lngCount)(intIndex) <> "" Then
                            blnNotAllBlank = True
                        End If
                    Next
                Next
                If blnNotAllBlank = False Then
                    FncCSVToArray = RET_NO_DATA
                End If
            End If
        End If
        '********** 2005/05/31 北澤 **********↑

        If FncCSVToArray = RET_FLD_MODO Then
            Dim blnNotAllBlankMarutoku As Boolean
            blnNotAllBlank = False
            If intRecNum > 0 Then
                For lngCount = 0 To (intRecNum - 1)
                    For intIndex = 3 To conIntKoumoku - 4
                        If paryCsvData(lngCount)(intIndex) <> "" Then
                            blnNotAllBlankMarutoku = True
                        End If
                    Next
                Next
                If blnNotAllBlankMarutoku = False Then
                    FncCSVToArray = RET_NO_DATA
                End If
            End If
        End If

        '********** 2005/06/15 北澤 **********↓
        '登録区分=2の店コードの保管用配列を再定義する
        If intRecNum > 0 Then
            ReDim arrTourokuKbn2(intRecNum - 1, 2)
        End If
        '********** 2005/06/15 北澤 **********↑

    End Function

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    FncHacTouroku

'*【処理概要】
'*?    発注中止日のデータを商品選定WFのDBに取り込む

'*【パラメータ】
'*P    strErrMsg;        string; エラーメッセージ;

'*【戻り値】
'*R    Boolean; TRUE:成功、FALSE:失敗;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/25; 黒沢(ユニシステム株式会社); 新規作成;
'*M    2005/05/12; 土田　祐歩(USC); パフォーマンス向上の為､DB参照エラーチェック用切断レコードを用意する処理追加;
'*M    2006/09/13; 和田(IC);        商品選定ＷＦ問合せ撲滅改善;
'*< Comment End >**************************************************'
Private Function FncHacTouroku(ByRef strErrMsg) As Boolean
    Dim intN As Integer
    Dim blnRet As Boolean
    Dim strJigCD As String
    Dim intNowLine As Integer
    Dim strCSVRecord() As String
    Dim flgChkError As Boolean      'チェック時のエラーフラグ（TRUE:エラーあり）
    
    FncHacTouroku = False
    
'********** 2005/05/18 土田 追加 **********↓
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(該当事業コードチェック用切断レコードセット生成)"
'********** 2005/05/18 土田 追加 **********↑

    '李 一括登録ログ追加 2018/12/17
    Call SetIkkatuLog("1-1", "発注中止日のデータを商品選定WFのDBに取り込む", "FncHacTouroku", "", "", pstrTantousyaId, "INS")
   
'*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    'ログ出力
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

'********** 2005/05/12 土田 追加 **********↓
'DB参照エラーチェック用切断レコードセットの生成
    Dim strSQL          As String
    Dim objCDBAccess    As Object
        
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    '一括登録時、毎回一つユーザーのみ実行できる
    Call LockTableIkkatu(objCDBAccess)

'該当事業コードチェック用切断レコードセット生成
    'SQL文作成（条件：設定内容の4桁目が1、KEY区分=W02）
    Call SetIkkatuLog("1-1-1", "該当事業コードチェック用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")
    strSQL = "SELECT key_cd "
    strSQL = strSQL & "FROM m_wf_meisyou WITH(readcommitted) "
    strSQL = strSQL & "WHERE Substring(settei_naiyou,4,1) = '1' " & vbCrLf
    strSQL = strSQL & "AND key_kbn = '" & KEYCD_KIKAKU_DOUNYU_TAISYOU & "' " & vbCrLf
    
        pHChkJigyouRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pHChkJigyouRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ
    
        Call SetIkkatuLog("1-1-1", "", "", "", "OK", "", "UPD")
        Call SetIkkatuLog("1-1-2", "MST_KEYマスタ1のKEYコードを抽出する（内部コードの重複をチェックする為）", "SQL", "", "", pstrTantousyaId, "INS")
        
'重複JANコードチェック用切断レコードセット生成
'********** 2005/05/18 土田 追加 **********↓
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(重複JANコードチェック用切断レコードセット生成)"
'********** 2005/05/18 土田 追加 **********↑

'*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    'ログ出力
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

'********** 2005/09/12 高附 追加 **********↓
'VMST_KEYマスタ1のKEYコードを抽出する（内部コードの重複をチェックする為）
    strSQL = "SELECT twk.jan_cd, mk1.key_cd, twk.kikaku_no" & vbCrLf
    strSQL = strSQL & " FROM t_syousen_wf_kanri as twk WITH(readcommitted)" & vbCrLf
    strSQL = strSQL & "      INNER JOIN v_key_1 as mk1 WITH(readcommitted)" & vbCrLf
    strSQL = strSQL & "              ON(twk.jan_cd = mk1.jan_cd)" & vbCrLf
    strSQL = strSQL & " WHERE" & vbCrLf
    strSQL = strSQL & " NOT(" & vbCrLf
    strSQL = strSQL & "      kikaku_syurui_cd = '" & pstrKikakuSyurui & "' " & vbCrLf
    strSQL = strSQL & "      AND nen_ren_no = '" & pstrKikakuNenrenban & "' " & vbCrLf
    strSQL = strSQL & "      AND eda_no = '" & pstrKikakuEdaban & "'" & vbCrLf
    strSQL = strSQL & "    )" & vbCrLf
    strSQL = strSQL & " AND twk.touroku_kbn = '2'" & vbCrLf
    
    
        pChkSameJANCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pChkSameJANCodeRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    pChkSameJANCodeRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ
'********** 2005/09/12 高附 追加 **********↑
'CSVファイル内Janコードのリスト(Janコードの重複無し)作成
'********** 2005/05/18 土田 追加 **********↓
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(CSVファイル内Janコードのリスト(Janコードの重複無し)作成)"
'*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    'ログ出力
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

    intNowLine = 0
    Do Until UBound(paryCsvData) < intNowLine
        
        strCSVRecord = paryCsvData(intNowLine)
         
        If InStr(pCsvJanCodeList, "'" & strCSVRecord(enmHacColNo.janCD) & "'") = 0 Then
            pCsvJanCodeList = pCsvJanCodeList & "'" & strCSVRecord(enmHacColNo.janCD) & "',"
        End If

'********** 2005/05/19 土田 追加 **********↓
'代替Janコードもリストに加える
        If InStr(pCsvJanCodeList, "'" & strCSVRecord(enmHacColNo.DaiJanCd) & "'") = 0 Then
            pCsvJanCodeList = pCsvJanCodeList & "'" & strCSVRecord(enmHacColNo.DaiJanCd) & "',"
        End If
'********** 2005/05/19 土田 追加 **********↑

        intNowLine = intNowLine + 1
    Loop
    
    If pCsvJanCodeList <> "" Then
        pCsvJanCodeList = Left(pCsvJanCodeList, Len(pCsvJanCodeList) - 1)
    Else
        pCsvJanCodeList = "''"
    End If
    intNowLine = 0
'********** 2005/05/18 土田 追加 **********↑
    Call SetIkkatuLog("1-1-2", "", "", "", "OK", "", "UPD")
    Call SetIkkatuLog("1-1-3", " 統合マスタ登録済みJANコードチェック用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")
    
'統合マスタ登録済みJANコードチェック用切断レコードセット生成
'********** 2005/05/18 土田 変更 **********↓
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(統合マスタ登録済みJANコードチェック用切断レコードセット生成)"
'*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    'ログ出力
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

'    strSQL = "SELECT key_cd , jan_cd FROM v_key_1 WITH(readcommitted)"
    strSQL = "SELECT jan_cd, key_cd FROM v_key_1 WITH(readcommitted)"
    strSQL = strSQL & " WHERE jan_cd IN(" & pCsvJanCodeList & ") "
    

        pChkvk1JANCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)

'    pChkvk1JANCodeRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ
    pChkvk1JANCodeRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
'********** 2005/05/18 土田 変更 **********↑

'********** 2005/05/18 土田 追加 **********↓
'統合マスタ登録済みJANコード紐付け主キーのリスト(主キー重複無し)作成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(CSVファイル内JANコード紐付け主キーのリスト(主キー重複無し)作成)"
'*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    'ログ出力
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    
    With pChkvk1JANCodeRS
        Do Until .EOF
            If InStr(pHKeyCDList, .Fields("key_cd")) = 0 Then
                pHKeyCDList = pHKeyCDList & "'" & .Fields("key_cd") & "',"
            End If
            .MoveNext
        Loop
    
        If pHKeyCDList <> "" Then
            pHKeyCDList = Left(pHKeyCDList, Len(pHKeyCDList) - 1)
        Else
            pHKeyCDList = "''"
        End If
    End With
        Call SetIkkatuLog("1-1-3", "", "", "", "OK", "", "UPD")
'********** 2005/05/18 土田 追加 **********↑

    
'********** 2005/05/18 土田 追加 **********↓
    Call SetIkkatuLog("1-1-4", "発注開始日チェック用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")

'発注開始日チェック用切断レコードセット生成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(発注開始日チェック用切断レコードセット生成)"
'*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    'ログ出力
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

    strSQL = "SELECT jan_cd, jigyou_cd, hattyuu_kaisi_date"
    strSQL = strSQL & " FROM t_syousen_jigyoubetu_settei WITH(readcommitted)"
    strSQL = strSQL & " WHERE kikaku_syurui_cd = '" & pstrKikakuSyurui & "' "
    strSQL = strSQL & " AND nen_ren_no = '" & pstrKikakuNenrenban & "' "
    strSQL = strSQL & " AND eda_no = '" & pstrKikakuEdaban & "'"
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "
    

        pHChkHattyuuKaisiDateRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)

    pHChkHattyuuKaisiDateRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    
'(商品選定WF(事業別設定情報項目)にデータが無い場合)発注開始日を事業別情報テーブルから取得用切断レコードセット生成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(発注開始日を事業別情報テーブルから取得用切断レコードセット生成)"
'*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    'ログ出力
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    Call SetIkkatuLog("1-1-4", "", "", "", "OK", "", "UPD")
   
    Call SetIkkatuLog("1-1-5", "(商品選定WF(事業別設定情報項目)にデータが無い場合)発注開始日を事業別情報テーブルから取得用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")

    strSQL = "SELECT key_cd, jigyou_cd, hattyuu_kaisi_date"
    strSQL = strSQL & " FROM v_jigyou_betu WITH(readcommitted)"
    strSQL = strSQL & " WHERE key_cd IN(" & pHKeyCDList & ") "
    
        pHChkJHTYKaisiDateRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pHChkJHTYKaisiDateRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ

    Call SetIkkatuLog("1-1-5", "", "", "", "OK", "", "UPD")
'********** 2005/05/18 土田 追加 **********↑
    
'発注中止理由区分チェック用切断レコードセット生成
    Call SetIkkatuLog("1-1-6", "発注中止理由区分チェック用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")
'********** 2005/05/18 土田 追加 **********↓
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(発注中止理由区分チェック用切断レコードセット生成)"
'********** 2005/05/18 土田 追加 **********↑
'*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    'ログ出力
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    
    strSQL = "SELECT key_cd "
    strSQL = strSQL & " FROM m_wf_meisyou WITH(readcommitted)"
    strSQL = strSQL & " WHERE key_kbn = '" & KEYCD_HATTYUU_TYUUSI_RIYU_KBN & "'"

        pHChkHacchuRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pHChkHacchuRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ

    Call SetIkkatuLog("1-1-6", "", "", "", "OK", "", "UPD")
'********** 2005/05/12 土田 追加 **********↑

    Call SetIkkatuLog("1-1-7", "全ての配列データのチェックを行う。", "SQL", "", "", pstrTantousyaId, "INS")
    '全ての配列データのチェックを行う。
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(全ての配列データのチェックを行う。)"
'*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    'ログ出力
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    
    flgChkError = False
    intNowLine = 0
    strErrMsg = ""
    Do Until UBound(paryCsvData) < intNowLine
        
        strCSVRecord = paryCsvData(intNowLine)
        
        '(1)DBに依存しないチェック：FncCheckParameter_CorrectForm_Hac
        '(2)DBを参照してチェック:FncCheckParameter_DataBase_Hac
        If FncCheckParameter_CorrectForm_Hac(strCSVRecord, intNowLine + 1) = False Or _
            FncCheckParameter_DataBase_Hac(strCSVRecord, intNowLine + 1) = False Then
            flgChkError = True
        End If
        
        intNowLine = intNowLine + 1
    Loop
    Call SetIkkatuLog("1-1-7", "", "", "", "OK", "", "UPD")
    
    Call SetIkkatuLog("1-1-8", "B参照エラーチェック用切断レコードセットの開放", "SQL", "", "", pstrTantousyaId, "INS")
'********** 2005/05/12 土田 追加 **********↓
'DB参照エラーチェック用切断レコードセットの開放
'********** 2005/05/18 土田 追加 **********↓
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(DB参照エラーチェック用切断レコードセットの開放)"
'********** 2005/05/18 土田 追加 **********↑
'*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    'ログ出力
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
        
    pHChkJigyouRS.Close
        pHChkJigyouRS = Nothing
        
    pChkSameJANCodeRS.Close
        pChkSameJANCodeRS = Nothing
    
    pChkvk1JANCodeRS.Close
        pChkvk1JANCodeRS = Nothing
    
    pHChkHacchuRS.Close
        pHChkHacchuRS = Nothing
        
    'DBクローズ
    objCDBAccess = Nothing
    'GetObjectContext.SetComplete
        Call SetIkkatuLog("1-1-8", "", "", "", "OK", "", "UPD")
        
'********** 2005/05/12 土田 追加 **********↑
    
    'データに１つでも問題がある場合は、ここで終了！！！
    If flgChkError = True Then
        'エラーがあればエラーメッセージを取得する。
        strErrMsg = FncGetErrorMsg_Hac()
        
'********** 2005/05/19 土田 追加 **********↓
        If Not pHChkHattyuuKaisiDateRS Is Nothing Then
                pHChkHattyuuKaisiDateRS = Nothing
        End If
        If Not pHChkJHTYKaisiDateRS Is Nothing Then
                pHChkJHTYKaisiDateRS = Nothing
        End If
 '********** 2005/05/19 土田 追加 **********↑
        
        Exit Function
    End If
    
'********** 2005/05/19 土田 追加 **********↓
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(データ更新処理開始)"
'********** 2005/05/19 土田 追加 **********↑
'*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    'ログ出力
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************

    Call SetIkkatuLog("1-2", "データ更新処理開始", "SQL", "", "", pstrTantousyaId, "INS")
    
    intN = 0
    '店別情報用フラグも一応クリアする。（ここでは使用しない）
    pflgBaikaToukbn = False
    pflgSiireToukbn = False
    pflgCenterToukbn = False
    
    '配列データをDBに取り込む
    Do Until UBound(paryCsvData) < intN         '配列の終端までループを繰り返します。
        '商品選定DB(WF管理項目)の先付変更の登録区分更新確認用フラグ
        '事業別情報登録区分は常に更新する為、フラグを立てる。
        pflgJigJyoToukbn = True
    
        '商品選定DB(事業別設定情報項目)の追加or更新
        strJigCD = paryCsvData(intN)(enmHacColNo.jigCD)      '該当事業コードを取得
        If strJigCD = ALL_JIG_CD Then
            blnRet = FncHacTourokuAllJig(strErrMsg, intN)
        Else
            blnRet = FncHacTourokuJig(strErrMsg, intN, strJigCD)
        End If

        If blnRet = False Then      'エラー処理
'********** 2005/05/19 土田 修正 **********↓
'            Exit Function           'エラー処理
            GoTo DefectErrorHandler:
'********** 2005/05/19 土田 修正 **********↑
        End If

'********** 2005/05/07 北澤 **********↓
'ある１事業のレコードを作成する際、他の全ての事業のレコードも作成
        If strJigCD <> ALL_JIG_CD Then
            blnRet = Insert_t_syousen_jigyoubetu_settei(strErrMsg, intN _
                   , pstrKikakuSyurui, pstrKikakuNenrenban, pstrKikakuEdaban _
                   , paryCsvData(intN)(enmTenColNo.janCD) _
                   , strJigCD _
                   , pstrTantousyaId)
            If blnRet = False Then
'********** 2005/05/19 土田 修正 **********↓
'              Exit Function           'エラー処理
              GoTo DefectErrorHandler:
'********** 2005/05/19 土田 修正 **********↑
            End If
        End If
'********** 2005/05/07 北澤 **********↑
        
'********** 2005/05/02 北澤 **********↓
'VSWF表示仕様テーブルにレコードが存在しない場合はVMST一般基本テーブルから情報を取得しレコードを追加する
'(照会などで商品名が表示されない為)
        blnRet = Insert_t_syousen_hyouji_siyou(strErrMsg, intN _
                , pstrKikakuSyurui, pstrKikakuNenrenban, pstrKikakuEdaban _
                , paryCsvData(intN)(enmTenColNo.janCD) _
                , pstrTantousyaId)
        If blnRet = False Then
'********** 2005/05/19 土田 修正 **********↓
'            Exit Function           'エラー処理
            GoTo DefectErrorHandler:
'********** 2005/05/19 土田 修正 **********↑
        End If
'********** 2005/05/02 北澤 **********↑
        
        '商品選定DB(WF管理項目)の更新
        If FncSetWFKanri(strErrMsg, intN, RET_FLD_HACH) = False Then
'********** 2005/05/19 土田 修正 **********↓
'            Exit Function           'エラー処理
            GoTo DefectErrorHandler:
'********** 2005/05/19 土田 修正 **********↑
        End If
        
        intN = intN + 1
    Loop
    
    FncHacTouroku = True

'********** 2005/05/19 土田 追加 **********↓
    If Not pHChkHattyuuKaisiDateRS Is Nothing Then
            pHChkHattyuuKaisiDateRS = Nothing
    End If
    If Not pHChkJHTYKaisiDateRS Is Nothing Then
            pHChkJHTYKaisiDateRS = Nothing
    End If
 '********** 2005/05/19 土田 追加 **********↑

'********** 2005/05/18 土田 追加 **********↓
'*** ↓2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
    'ログ出力
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", "FncHacTouroku_処理終了", pstrTantousyaId, "3")
'*** ↑2018/11/19 山田 予期せぬエラー対応 ****************************************************************************************
        Call SetIkkatuLog("1-2", "", "", "", "OK", "", "UPD")
Exit Function
DefectErrorHandler:
    
    If Not pHChkJigyouRS Is Nothing Then
            pHChkJigyouRS = Nothing
    End If
    If Not pChkSameJANCodeRS Is Nothing Then
            pChkSameJANCodeRS = Nothing
    End If
    If Not pChkvk1JANCodeRS Is Nothing Then
            pChkvk1JANCodeRS = Nothing
    End If
    If Not pHChkHattyuuKaisiDateRS Is Nothing Then
            pHChkHattyuuKaisiDateRS = Nothing
    End If
    If Not pHChkJHTYKaisiDateRS Is Nothing Then
            pHChkJHTYKaisiDateRS = Nothing
    End If
    If Not pHChkHacchuRS Is Nothing Then
            pHChkHacchuRS = Nothing
    End If
    
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    Call SetIkkatuLog("1-1", "", "", Left(Err.Description, 400), "NG", "", "UPD")
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description
    
    
'********** 2005/05/18 土田 追加 **********↑

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckParameter_CorrectForm_Hac
'*【処理概要】
'*?    CSVファイルのパラメータに対して、DBに依存しない形式チェックを行います。
'*?    （発注中止日用チェック）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckParameter_CorrectForm; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/25; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncCheckParameter_CorrectForm_Hac(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheckParameter_CorrectForm_Hac = False
    
    Dim blnIsOKRequired
    Dim blnIsOKNumeric
    Dim blnIsOKLimit
    Dim blnIsOKHacEndDay
    Dim blnIsOKCheckDigit
    Dim blnIsOKCheckDigit_DaiJAN
    Dim blnIsOKSyobun
'********** 2005/05/07 北澤 **********↓
    Dim blnIsOKNotZero
'********** 2005/05/07 北澤 **********↑
    
    'a.必須項目（該当事業コード、発注終了日、JANコード）が入っているか確認する。
        blnIsOKRequired = FncCheck_RequiredParameter_Hac(arryCSVRecord, intNowLine)
    
    'b.全ての項目において、半角数字であるか確認する。
        blnIsOKNumeric = FncCheck_IsParameterNumeric_Hac(arryCSVRecord, intNowLine)
    
    'c.桁数が制限を超えていないか確認する。
        blnIsOKLimit = FncCheck_Limit_Hac(arryCSVRecord, intNowLine)

    'd.発注終了日の日付が有効であるか確認する。
    blnIsOKHacEndDay = FncCheck_HacEndDay_Hac(arryCSVRecord(enmHacColNo.SyuryouDate), intNowLine)
        
    'e.JANコードに正しいチェックデジットの値がセットされているかを調べる
    blnIsOKCheckDigit = FncCheck_CheckDigit(arryCSVRecord(enmHacColNo.janCD), intNowLine, 2)
    
    '代替JANコードに正しいチェックディジットの値がセットされているかを調べる
    blnIsOKCheckDigit_DaiJAN = True
    If arryCSVRecord(enmHacColNo.DaiJanCd) <> "" Then
        blnIsOKCheckDigit_DaiJAN = FncCheck_CheckDigit(arryCSVRecord(enmHacColNo.DaiJanCd), intNowLine, 3)
    End If
    
    'f.処分方法が"0"（未設定）、"1"（返品）、"2"（一時停止）、"3"（取扱終了（店振り対象））の
    'いずれかであるか確認する。
    blnIsOKSyobun = True
    If arryCSVRecord(enmHacColNo.Syobun) <> "" Then
        blnIsOKSyobun = FncCheck_CheckSyobun(arryCSVRecord(enmHacColNo.Syobun), intNowLine)
    End If
    
'********** 2005/05/07 北澤 **********↓
'金額類が正の値かをチェックする
        blnIsOKNotZero = FncCheck_NotZero(arryCSVRecord, intNowLine, "HAC")
'********** 2005/05/07 北澤 **********↑
    
'********** 2005/05/07 北澤 **********↓
'    If blnIsOKRequired And blnIsOKNumeric And blnIsOKLimit And _
'       blnIsOKHacEndDay And blnIsOKCheckDigit And blnIsOKCheckDigit_DaiJAN And _
'       blnIsOKSyobun Then
    If blnIsOKRequired And blnIsOKNumeric And blnIsOKLimit And _
       blnIsOKHacEndDay And blnIsOKCheckDigit And blnIsOKCheckDigit_DaiJAN And _
       blnIsOKSyobun And blnIsOKNotZero Then
'********** 2005/05/07 北澤 **********↑
        FncCheckParameter_CorrectForm_Hac = True
    End If
    
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_RequiredParameter_Hac
'*【処理概要】
'*?    CSVデータの１行に対し、必須項目が空白でないかどうかチェックします。
'*?    （発注中止日用チェック）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_RequiredParameter; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/25; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'

Private Function FncCheck_RequiredParameter_Hac(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_RequiredParameter_Hac = True
    
    If arryCSVRecord(enmHacColNo.jigCD) = "" Then       '該当事業コード
            FncSetError(intNowLine, enmTenColNo.TenCd, PARAM_EMPTY)
        FncCheck_RequiredParameter_Hac = False
    End If
    If arryCSVRecord(enmHacColNo.SyuryouDate) = "" Then  '発注終了日
            FncSetError(intNowLine, enmTenColNo.HenkouDate, PARAM_EMPTY)
        FncCheck_RequiredParameter_Hac = False
    End If
    If arryCSVRecord(enmHacColNo.janCD) = "" Then       'JANコード
            FncSetError(intNowLine, enmTenColNo.janCD, PARAM_EMPTY)
        FncCheck_RequiredParameter_Hac = False
    End If

'********** 2005/05/07 北澤 **********↓
'発注中止理由は必須
    If arryCSVRecord(enmHacColNo.ChushiKbn) = "" Then
            FncSetError(intNowLine, enmHacColNo.ChushiKbn, PARAM_EMPTY)
        FncCheck_RequiredParameter_Hac = False
    End If
'発注中止理由は00(未選択)は不可
    If arryCSVRecord(enmHacColNo.ChushiKbn) = "00" Then
            FncSetError(intNowLine, enmHacColNo.ChushiKbn, PARAM_MISENTAKU)
        FncCheck_RequiredParameter_Hac = False
    End If
'********** 2005/05/07 北澤 **********↑

'********** 2005/05/09 北澤 **********↓
''********** 2005/05/07 北澤 **********↓
''処分方法を一時的に入力必須にする(処分方法を基幹マスタから取得するロジックが完成したら削除する)
'    If arryCSVRecord(enmHacColNo.Syobun) = "" Then
'        FncSetError (intNowLine, enmHacColNo.Syobun, PARAM_EMPTY
'        FncCheck_RequiredParameter_Hac = False
'    End If
''代替商品コードを一時的に入力必須にする(代替商品コードを基幹マスタから取得するロジックが完成したら削除する)
'    If arryCSVRecord(enmHacColNo.DaiJanCd) = "" Then
'        FncSetError (intNowLine, enmHacColNo.DaiJanCd, PARAM_EMPTY
'        FncCheck_RequiredParameter_Hac = False
'    End If
''********** 2005/05/07 北澤 **********↑
'********** 2005/05/09 北澤 **********↑

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_IsParameterNumeric_Hac
'*【処理概要】
'*?    CSVデータの１行に対し、全フィールドが半角数値かどうかチェックします。
'*?    （発注中止日用チェック）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_IsParameterNumeric_Hac; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/25; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterNumeric_Hac(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_IsParameterNumeric_Hac = True

    If FncChkHankakuNumber(arryCSVRecord(enmHacColNo.jigCD)) = False Then       '該当事業コード
        FncSetError (intNowLine, enmTenColNo.TenCd, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Hac = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmHacColNo.SyuryouDate)) = False Then '発注終了日
        FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Hac = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmHacColNo.janCD)) = False Then       'JANコード
        FncSetError (intNowLine, enmTenColNo.janCD, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Hac = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmHacColNo.Syobun)) = False Then      ' 処分方法
        FncSetError (intNowLine, enmHacColNo.Syobun, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Hac = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmHacColNo.ChushiKbn)) = False Then   ' 発注中止理由区分
    
        FncSetError (intNowLine, enmHacColNo.ChushiKbn, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Hac = False
        
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmHacColNo.DaiJanCd)) = False Then    ' 代替JANコード
    
        FncSetError (intNowLine, enmHacColNo.DaiJanCd, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Hac = False
    End If

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_Limit_Hac
'*【処理概要】
'*?    CSVデータの１行に対し、桁数が制限を越えていないかを確認します。
'*?    （発注中止日用チェック）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_Limit_Hac; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/25; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
Function FncCheck_Limit_Hac(arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_Limit_Hac = True
    
    '(b)JANコード
    If FncCheck_JANCode(arryCSVRecord(enmHacColNo.janCD)) = False Then
        FncSetError (intNowLine, enmHacColNo.janCD, PARAM_SHORTDISIT)
        FncCheck_Limit_Hac = False
    End If

    '(c)代替JANコード
    If arryCSVRecord(enmHacColNo.DaiJanCd) <> "" Then
    
        If FncCheck_JANCode(arryCSVRecord(enmHacColNo.DaiJanCd)) = False Then
            FncSetError (intNowLine, enmHacColNo.DaiJanCd, PARAM_SHORTDISIT)
            FncCheck_Limit_Hac = False
        End If
    End If

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_HacEndDay_Hac
'*【処理概要】
'*?    発注中止日の日付が有効であるかを確認します。
'*?    （発注中止日用チェック）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strHacEndDay ; String ; 発注終了日をあらわす文字列
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_CheckSyobun; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/25; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
Function FncCheck_HacEndDay_Hac(strHacEndDay As String, intNowLine As Integer)

    FncCheck_HacEndDay_Hac = True
    ' 桁数が８桁である事と、変更日の日付が有効であるかを確認する。
    If Len(strHacEndDay) <> 8 Or FncCheckDate(strHacEndDay) = False Then
        FncSetError (intNowLine, enmHacColNo.SyuryouDate, PARAM_WRONGDATE)
        FncCheck_HacEndDay_Hac = False
    End If

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_CheckSyobun
'*【処理概要】
'*?    処分方法が規定されている値に含まれることを確認します。
'*?    （発注中止日用チェック）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strSyobun ; String ; 処分方法をあらわす文字列
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_CheckSyobun; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/25; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'

Function FncCheck_CheckSyobun(strSyobun As String, intNowLine As Integer) As Boolean
    
    ' 処分方法が0,1,2,3のいずれかであることを確かめる
    
    Select Case strSyobun
        Case "0":   FncCheck_CheckSyobun = True
        Case "1":   FncCheck_CheckSyobun = True
        Case "2":   FncCheck_CheckSyobun = True
        Case "3":   FncCheck_CheckSyobun = True
        Case Else
            
                FncSetError(intNowLine, enmHacColNo.Syobun, PARAM_NOTDEFINED)
            FncCheck_CheckSyobun = False
        
    End Select
    
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckParameter_DataBase_Hac
'*【処理概要】
'*?    CSVファイルのパラメータに対して、DBにアクセスしてチェックを行います。
'*?    (発注中止版)
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVRecord() ; String CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckParameter_DataBase_Hac; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/25; 浅野　正徳(USC); 新規作成;
'*M    2005/09/12; 高附　周造(IC);  変更 内部コードの重複チェック
'*< Comment End >**************************************************'
Private Function FncCheckParameter_DataBase_Hac(ByRef arryCSVRecord() As String, intNowLine As Integer)

    Dim blnIsOKJigyouCode
    Dim blnIsOKSyuryouDate
    Dim blnIsOKJANCode
    Dim blnIsOKNaibuCode
    Dim blnIsOKChushiKbn
    Dim blnIsOKDaiJANCode
    
    Dim strKeyCode As String
    
    'a.該当事業コードがWF名称マスタテーブルに登録されているか確認する。
    blnIsOKJigyouCode = FncCheckDB_JigyouCode(arryCSVRecord(enmHacColNo.jigCD), intNowLine)

    'b1.JANコードが有効か確認する。
    blnIsOKJANCode = FncCheckDB_JANCode(arryCSVRecord(enmHacColNo.janCD), strKeyCode, intNowLine, RET_FLD_HACH)
    
    'b2.KEYコードが有効か確認する。
    blnIsOKNaibuCode = FncCheckDB_NaibuCode(strKeyCode, intNowLine, RET_FLD_HACH)
    
    'c.発注終了日が発注開始日より後であるか確認する。
    '(JANコードをキーとして検索するため、JANコードの有効性を先に確認する：有効ならば検索を行う）
    blnIsOKSyuryouDate = True
    If blnIsOKJANCode = True Then
        blnIsOKSyuryouDate = FncCheckDB_SyuryouDate(arryCSVRecord, strKeyCode, intNowLine)
    End If
    
    'd.発注中止理由区分が有効か確認する。
    If arryCSVRecord(enmHacColNo.ChushiKbn) <> "" Then
        blnIsOKChushiKbn = FncCheckDB_ChushiKbn(arryCSVRecord(enmHacColNo.ChushiKbn), intNowLine)
    Else
        blnIsOKChushiKbn = True
    End If
    
    'e.代替JANコードが統合マスタに登録されているか確認する。
    If arryCSVRecord(enmHacColNo.DaiJanCd) <> "" Then
        blnIsOKDaiJANCode = FncCheckDB_DaiJANCode(arryCSVRecord(enmHacColNo.DaiJanCd), intNowLine)
    Else
        blnIsOKDaiJANCode = True
    End If
    
    If blnIsOKJigyouCode And blnIsOKJANCode And blnIsOKSyuryouDate And blnIsOKChushiKbn _
        And blnIsOKDaiJANCode And blnIsOKNaibuCode Then
        FncCheckParameter_DataBase_Hac = True
    End If
    
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!     FncCheckDB_JigyouCode
'*【処理概要】
'*?     事業コードがWF名称マスタに登録されていることを確認する。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P     strJigyouCode ; String ; 事業コードを表す文字列
'*P     intNowLine ; Integer ; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R     FncCheckDB_JigyouCode ; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/25; 浅野　正徳(USC); 新規作成;
'*M    2005/05/12; 土田　祐歩(USC); パフォーマンス向上の為､この関数内での切断レコード取得をやめ大域変数の切断レコードを利用する様修正;
'*< Comment End >**************************************************'
'
Private Function FncCheckDB_JigyouCode(strJigyouCode As String, intNowLine As Integer)

'********** 2005/05/12土田 削除 **********↓
'この関数内での切断レコードセット取得処理削除
'    Dim strSql            As String
'
'    '　展開している事業コードは、WF名称マスタから取得する。
'    '   　DB：      WF名称マスタ（m_wf_meisyou）
'    '   　キー：        KEY区分（key_kbn）                          「W02」
'    '           KEYコード（key_cd）                         CSVファイルから取得
'    '           設定内容（settei_naiyou）                           4桁目に"1"が立っている
'    '　全事業（該当事業コード欄に「99」指定）の場合は、KEYコードを指定せずに検索
'    '　する事により、展開している事業コード全てを取得できる。
'    '　検索に１件もヒットしない場合は、エラーとする。
'
'    strSql = "SELECT key_cd "
'    strSql = strSql & "FROM m_wf_meisyou WITH(readcommitted) "
'    strSql = strSql & "WHERE Substring(settei_naiyou,4,1) = '1' " & vbCrLf
'    strSql = strSql & "AND key_kbn = '" & KEYCD_KIKAKU_DOUNYU_TAISYOU & "' " & vbCrLf
'
'    '全事業でない場合は、事業コードを指定する。
'    If strJigyouCode <> ALL_JIG_CD Then
'        strSql = strSql & "AND key_cd = '" & strJigyouCode & "'"      '各事業
'    End If
'
'    Dim IsExistJigyouCode
'    IsExistJigyouCode = FncCheckDB_IsRecordExists("VSWF", strSql)
'
'    If IsExistJigyouCode = False Then
'        FncSetError (intNowLine, enmHacColNo.JigCd, PARAM_NOTFINDMASTER
'    End If
'
'    FncCheckDB_JigyouCode = IsExistJigyouCode
'********** 2005/05/12土田 削除 **********↑
    
'********** 2005/05/12土田 追加 **********↓
'大域変数の切断レコードセットを使用しエラーチェック
    With pHChkJigyouRS
'********** 2005/05/17 土田 変更 **********↓
'********** 2005/05/19 土田 変更 **********↓
'       If Not .EOF Then
       .Filter = adFilterNone
       If Not .BOF Then
'********** 2005/05/19 土田 変更 **********↑
            .MoveFirst
            '全事業でない場合は、事業コードを指定する。
            If strJigyouCode <> ALL_JIG_CD Then
                    .Find("key_cd = '" & strJigyouCode & "'")
            End If
        End If
'********** 2005/05/17 土田 変更 **********↑
    End With
    
    If pHChkJigyouRS.EOF Then
            FncSetError(intNowLine, enmHacColNo.jigCD, PARAM_NOTFINDMASTER)
    End If

    FncCheckDB_JigyouCode = Not pHChkJigyouRS.EOF
'********** 2005/05/12土田 追加 **********↑
    
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!     FncCheckDB_SyuryouDate
'*【処理概要】
'*?     発注終了日が発注開始日よりも後の日付であるかどうかを確認します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P     arryCSVRecord() ; String CSVデータを１行分格納する配列へのポインタ
'*P     strKeyCode ; String 統合マスタから取得した、JANコードに紐付けられたKeyCode
'*P     intNowLine ; Integer CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R     FncCheckDB_SyuryouDate ; Boolean チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/25; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'

Private Function FncCheckDB_SyuryouDate(ByRef arryCSVRecord() As String, strKeyCode As String, intNowLine As Integer)
    
    Dim strKaisiDate As String
    Dim IsOKGetKaisiDate As Boolean
    
    FncCheckDB_SyuryouDate = False
    
    ' 発注開始日を取得する。
    
'********** 2005/05/18 土田 変更 **********↓
'    If arryCSVRecord(enmHacColNo.jigCD) = ALL_JIG_CD Then
        IsOKGetKaisiDate = FncCheckDB_GetKaisiDateAllJig(arryCSVRecord, strKaisiDate, strKeyCode)
'    Else
'        IsOKGetKaisiDate = FncCheckDB_GetKaisiDate(arryCSVRecord, strKaisiDate, strKeyCode)
'    End If
'********** 2005/05/18 土田 追加 **********↑
    
    ' 発注開始日＜発注終了日かどうかチェックする。
    If IsOKGetKaisiDate = True Then
        
        ' 発注開始日＞＝発注終了日の場合はエラー
        If StrComp(strKaisiDate, arryCSVRecord(enmHacColNo.SyuryouDate)) >= 0 Then
            FncCheckDB_SyuryouDate = False
                FncSetError(intNowLine, enmHacColNo.SyuryouDate, PARAM_EARLYCANCELDAY)
            Exit Function
        End If
        
        ' すべての発注開始日は発注終了日より前の日付
        FncCheckDB_SyuryouDate = True
    
    Else
        ' 発注開始日が見つからない場合
            FncSetError(intNowLine, enmHacColNo.SyuryouDate, PARAM_NOTEXISTDB)
        FncCheckDB_SyuryouDate = False
    
    End If
        
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!     FncCheckDB_GetKaisiDateAllJig
'*【処理概要】
'*?     商品選定WF(事業別設定情報項目)から発注開始日を取得します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P     arryCSVRecord() ; String CSVデータを１行分格納する配列へのポインタ
'*P     strKaisiDate ; String 発注開始日を格納する文字列
'*P     strKeyCode ; String ; 統合マスタから取得した、JANコードに紐付けられたKeyCode
'*【戻り値】
'*    < Name Type Comment >
'*R     FncCheckDB_GetKaisiDateAllJig ; Boolean 開始日取得成功:True,失敗:False
'*R     strKaisiDate ; 発注開始日()を格納
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/27; 黒沢(USC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheckDB_GetKaisiDateAllJig(ByRef arryCSVRecord() As String, ByRef strKaisiDate As String, strKeyCode As String) As Boolean

'********** 2005/05/18 土田 削除 **********↓
'    Dim objCDBAccess As Object
'    Dim objRSJig As Object
'    Dim objRS As Object
'
'    Dim strSQL As String
'    Dim strJigCd As String
'    Dim strTmpKaisiDate As String       '各事業毎の発注開始日
'    Dim dtKaisiDate As Date
'
'    Dim strMethodCallInfo As String
'    on error GoTo DefectErrorHandler
'    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncCheckDB_GetKaisiDateAllJig()"
'
'    'CDBAccessインスタンス生成
'    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
'
'    '展開している事業コードを取得する。
'    strSQL = "SELECT key_cd "
'    strSQL = strSQL & "FROM m_wf_meisyou WITH(readcommitted) "
'    strSQL = strSQL & "WHERE Substring(settei_naiyou,4,1) = '1' " & vbCrLf
'    strSQL = strSQL & "AND key_kbn = '" & KEYCD_KIKAKU_DOUNYU_TAISYOU & "' " & vbCrLf
'    'SQL実行
'        objRSJig = objCDBAccess.GetDCRecordset("VSWF", strSQL)
'
'    If objRSJig.EOF = True Then
'        '展開している事業コードが無い場合は、ここで終了。
'        FncCheckDB_GetKaisiDateAllJig = False
'        GoTo Proc_End:
'    Else
'        Do Until objRSJig.EOF
'            strJigCd = CStr(objRSJig.Fields("key_cd"))
'
'            '各事業毎に発注開始日を取得する。
'            '　発注開始日と終了日が、発注開始日＜発注終了日となっているか確認する。
'            '   　DB：      商品選定DB（事業別設定情報項目）（t_syousen_jigyoubetu_settei）
'            '   　キー：        企画種類ｺｰﾄﾞ(kikaku_syurui_cd)                          前画面から取得
'            '           年連番（nen_ren_no）                            前画面から取得
'            '           枝番（eda_no）                          前画面から取得
'            '           JANコード（jan_cd）                         CSVファイルから取得
'            '           事業コード（jigyou_cd）                         CSVファイルから取得
'
'            strSQL = "SELECT hattyuu_kaisi_date"
'            strSQL = strSQL & " FROM t_syousen_jigyoubetu_settei WITH(readcommitted)"
'            strSQL = strSQL & " WHERE kikaku_syurui_cd = '" & pstrKikakuSyurui & "' "
'            strSQL = strSQL & " AND nen_ren_no = '" & pstrKikakuNenrenban & "' "
'            strSQL = strSQL & " AND eda_no = '" & pstrKikakuEdaban & "'"
'            strSQL = strSQL & " AND jan_cd = '" & arryCSVRecord(enmHacColNo.JanCd) & "'"
'            strSQL = strSQL & " AND jigyou_cd = '" & strJigCd & "'"
'
'            'SQL実行
'                objRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
'
'            If objRS.EOF = True Then
'                '商品選定WF(事業別設定情報項目)にデータが無い場合は、統合マスタ(事業別情報)から取得する。
'                If FncCheckDB_GetKaisiDateMas(strKaisiDate, strKeyCode, strJigCd) = False Then
'                    FncCheckDB_GetKaisiDateAllJig = False
'                    GoTo Proc_End:
'                End If
'            Else    '発注開始日データがNULLor空白である場合は、統合マスタ(事業別情報)から取得する。
'                If IsNull(Trim(objRS.Fields("hattyuu_kaisi_date"))) = True Then
'                    If FncCheckDB_GetKaisiDateMas(strKaisiDate, strKeyCode, strJigCd) = False Then
'                        FncCheckDB_GetKaisiDateAllJig = False
'                        GoTo Proc_End:
'                    End If
'                ElseIf Trim(objRS.Fields("hattyuu_kaisi_date")) = "" Then
'                    If FncCheckDB_GetKaisiDateMas(strKaisiDate, strKeyCode, strJigCd) = False Then
'                        FncCheckDB_GetKaisiDateAllJig = False
'                        GoTo Proc_End:
'                    End If
'                Else
'                    ' 先付変更DBテーブルから発注開始日を取得
'                    ' 先付変更DBの日付書式が異なるため、ここで８桁に調整
'                    dtKaisiDate = CDate(Trim(objRS.Fields("hattyuu_kaisi_date")))
'                    strKaisiDate = Format(dtKaisiDate, "yyyymmdd")
'                End If
'            End If
'
'            '最も新しい発注開始日のみを変数に格納する。
'            If StrComp(strTmpKaisiDate, strKaisiDate) > 0 Then
'                strKaisiDate = strTmpKaisiDate
'            End If
'
'            objRSJig.MoveNext
'        Loop
'
'    End If
'
'    FncCheckDB_GetKaisiDateAllJig = True
'
'    'DBクローズ
''    objRS.Close
'        objRSJig = Nothing
'        objRS = Nothing
'    objCDBAccess = Nothing
'
'    'GetObjectContext.SetComplete
'********** 2005/05/18 土田 削除 **********↑

'********** 2005/05/18 土田 追加 **********↓

    Dim strTmpKaisiDate As String       '各事業毎の発注開始日
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    
    'チェック対象事業コードの取得
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.FncCheckDB_GetKaisiDateAllJig(チェック対象事業コードの取得)"
    
    'このレコードの事業コード
    Dim jigCD As String
    jigCD = arryCSVRecord(enmHacColNo.jigCD)
    
    'チェック対象事業コードの取得に該当事業コードチェック用RSを再利用する
    With pHChkJigyouRS

        '***  チェック対象事業コードの絞込み ***
        
        'フィルターを解除しておく
        .Filter = adFilterNone
        
        If jigCD = ALL_JIG_CD Then
            '全事業の場合フィルタリング無し
        Else
            '各事業の場合
            '該当事業コードチェック用RSを事業コードでフィルタリングする
            .Filter = "key_cd = '" & jigCD & "'"
        End If
        
        If .EOF Then
            '展開している事業コードが無い場合は、ここで終了。
            FncCheckDB_GetKaisiDateAllJig = False
            Exit Function
        Else
            '***  チェック対象事業コードのループ ***
            Do Until .EOF
                Dim strJigCD As String
                strJigCD = CStr(.Fields("key_cd"))
                
                With pHChkHattyuuKaisiDateRS
                    '発注開始日を取得する。
                    .Filter = "jan_cd = '" & arryCSVRecord(enmHacColNo.janCD) & "' AND jigyou_cd = '" & strJigCD & "'"
                    
                    '商品選定WF(事業別設定情報項目)にデータが無い場合、又は
                    '発注開始日データがNULL、又は空白である場合、統合マスタ(事業別情報)から取得する。
                    'strMethodCallInfo = App.Title & ".VswfIkatuTou.FncCheckDB_GetKaisiDateAllJig(チェック対象事業コードの取得:統合マスタ(事業別情報)から取得する。)"
                    
                    If .EOF Then
                        If FncCheckDB_GetKaisiDateMas(strKaisiDate, strKeyCode, strJigCD) = False Then
                            FncCheckDB_GetKaisiDateAllJig = False
                            Exit Function
                        End If
                    Else
                    
                        If IsNull(Trim(.Fields("hattyuu_kaisi_date"))) Or _
                            Trim(.Fields("hattyuu_kaisi_date")) = "" Then
                                
                            If FncCheckDB_GetKaisiDateMas(strKaisiDate, strKeyCode, strJigCD) = False Then
                                FncCheckDB_GetKaisiDateAllJig = False
                                Exit Function
                            End If
                        Else
                            ' 先付変更DBテーブルから発注開始日を取得
                            ' 先付変更DBの日付書式が異なるため、ここで８桁に調整
'                            strKaisiDate = Format(CDate(Trim(.Fields("hattyuu_kaisi_date"))), "yyyymmdd")
                            strTmpKaisiDate = Format(CDate(Trim(.Fields("hattyuu_kaisi_date"))), "yyyymmdd")
                        End If
                    End If
                    
                    '最も新しい発注開始日のみを変数に格納する。
                    If StrComp(strTmpKaisiDate, strKaisiDate) > 0 Then
                        strKaisiDate = strTmpKaisiDate
                    End If
                End With
                .MoveNext
            Loop
        End If
    End With
    FncCheckDB_GetKaisiDateAllJig = True
'********** 2005/05/18 土田 追加 **********↑

Exit Function

'********** 2005/05/18 土田 削除 **********↓
'Proc_End:
'    If Not objRSJig Is Nothing Then
'            objRSJig = Nothing
'    End If
'    If Not objRS Is Nothing Then
'            objRS = Nothing
'    End If
'    If Not objCDBAccess Is Nothing Then
'        objCDBAccess = Nothing
'    End If
'    Exit Function
'********** 2005/05/18 土田 削除 **********↑
DefectErrorHandler:
'********** 2005/05/18 土田 削除 **********↓
'    If Not objRSJig Is Nothing Then
'            objRSJig = Nothing
'    End If
'    If Not objRS Is Nothing Then
'            objRS = Nothing
'    End If
'    If Not objCDBAccess Is Nothing Then
'        objCDBAccess = Nothing
'    End If
'    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
'********** 2005/05/18 土田 削除 **********↑
'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!     FncCheckDB_GetKaisiDate
'*【処理概要】
'*?     商品選定WF(事業別設定情報項目)から発注開始日を取得します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P     arryCSVRecord() ; String CSVデータを１行分格納する配列へのポインタ
'*P     strKaisiDate ; String 発注開始日を格納する文字列
'*P     strKeyCode ; String ; 統合マスタから取得した、JANコードに紐付けられたKeyCode
'*【戻り値】
'*    < Name Type Comment >
'*R     FncCheckDB_GetKaisiDate ; Boolean 開始日取得成功:True,失敗:False
'*R     strKaisiDate ; 発注開始日()を格納
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/25; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheckDB_GetKaisiDate(ByRef arryCSVRecord() As String, ByRef strKaisiDate As String, strKeyCode As String) As Boolean

    Dim objCDBAccess As Object
    Dim objRs As Object
    Dim strSQL As String
    Dim strTmp As String
    Dim dtKaisiDate As Date
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncCheckDB_GetKaisiDate()"

    '　発注開始日と終了日が、発注開始日＜発注終了日となっているか確認する。
    '   　DB：      商品選定DB（事業別設定情報項目）（t_syousen_jigyoubetu_settei）
    '   　キー：        企画種類ｺｰﾄﾞ(kikaku_syurui_cd)                          前画面から取得
    '           年連番（nen_ren_no）                            前画面から取得
    '           枝番（eda_no）                          前画面から取得
    '           JANコード（jan_cd）                         CSVファイルから取得
    '           事業コード（jigyou_cd）                         CSVファイルから取得

    strSQL = "SELECT hattyuu_kaisi_date"
    strSQL = strSQL & " FROM t_syousen_jigyoubetu_settei WITH(readcommitted)"
    strSQL = strSQL & " WHERE kikaku_syurui_cd = '" & pstrKikakuSyurui & "' "
    strSQL = strSQL & " AND nen_ren_no = '" & pstrKikakuNenrenban & "' "
    strSQL = strSQL & " AND eda_no = '" & pstrKikakuEdaban & "'"
    strSQL = strSQL & " AND jan_cd = '" & arryCSVRecord(enmHacColNo.janCD) & "'"
    strSQL = strSQL & " AND jigyou_cd = '" & arryCSVRecord(enmHacColNo.jigCD) & "'"
        
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL実行
        objRs = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    
    If objRs.EOF = True Then
        '商品選定WF(事業別設定情報項目)にデータが無い場合は、統合マスタ(事業別情報)から取得する。
        If FncCheckDB_GetKaisiDateMas(strKaisiDate, strKeyCode, arryCSVRecord(enmHacColNo.jigCD)) = False Then
            FncCheckDB_GetKaisiDate = False
            GoTo Proc_End:
        End If
    Else    '発注開始日データがNULLor空白である場合は、統合マスタ(事業別情報)から取得する。
            If IsNull(Trim(objRs.Fields("hattyuu_kaisi_date"))) = True Then
                If FncCheckDB_GetKaisiDateMas(strKaisiDate, strKeyCode, arryCSVRecord(enmHacColNo.jigCD)) = False Then
                    FncCheckDB_GetKaisiDate = False
                    GoTo Proc_End
                End If
            ElseIf Trim(objRs.Fields("hattyuu_kaisi_date")) = "" Then
                If FncCheckDB_GetKaisiDateMas(strKaisiDate, strKeyCode, arryCSVRecord(enmHacColNo.jigCD)) = False Then
                    FncCheckDB_GetKaisiDate = False
                    GoTo Proc_End
                End If
            Else
                ' 先付変更DBテーブルから発注開始日を取得
                ' 先付変更DBの日付書式が異なるため、ここで８桁に調整
                dtKaisiDate = CDate(Trim(objRs.Fields("hattyuu_kaisi_date")))
                strKaisiDate = Format(dtKaisiDate, "yyyymmdd")
            End If
    End If
    
    FncCheckDB_GetKaisiDate = True
    
    'DBクローズ
'    objRS.Close
        objRs = Nothing
    objCDBAccess = Nothing
    
    'GetObjectContext.SetComplete
Exit Function

Proc_End:
    If Not objRs Is Nothing Then
            objRs = Nothing
    End If
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
    Exit Function

DefectErrorHandler:
    If Not objRs Is Nothing Then
            objRs = Nothing
    End If
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!     FncCheckDB_GetKaisiDateMas
'*【処理概要】
'*?     統合マスタ(事業別情報)から発注開始日を取得します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P     strKaisiDate ; String 発注開始日を格納する文字列
'*P     strKeyCode ; String ; 統合マスタから取得した、JANコードに紐付けられたKeyCode
'*P     strJigCd ; String ; 事業コード
'*【戻り値】
'*    < Name Type Comment >
'*R     FncCheckDB_GetKaisiDateMas ; Boolean 開始日取得成功:True,失敗:False
'*R     strKaisiDate ; 発注開始日()を格納
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/27; 浅野　正徳(USC); 新規作成;
'*M    2005/05/18; 土田　祐歩(USC); パフォーマンス向上の為､この関数内での切断レコード取得をやめ大域変数の切断レコードを利用する様修正;
'*< Comment End >**************************************************'
Private Function FncCheckDB_GetKaisiDateMas(ByRef strKaisiDate As String, _
                                      ByVal strKeyCode As String, _
                                      ByVal strJigCD As String) As Boolean
                    
'********** 2005/05/18 土田 削除 **********↓
'    Dim objCDBAccess As Object
'    Dim objRS As Object
'    Dim strSQL As String
'    Dim strTmp As String
'********** 2005/05/18 土田 削除 **********↑
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncCheckDB_GetKaisiDateMas()"
    
'********** 2005/05/18 土田 削除 **********↓
'    '　先付変更DBに登録されていない場合は、発注開始日を事業別情報テーブルから
'    '　予め取得しておく。
'    '   　DB：      事業別情報テーブル（m_jigyou_betu）
'    '   　キー：        主キー（key_cd）                            1.(2)c.(b) KEY情報1から取得
'    '           事業コード（jigyou_cd）                         CSVファイルから取得
'    strSQL = "SELECT hattyuu_kaisi_date"
'    strSQL = strSQL & " FROM v_jigyou_betu WITH(readcommitted)"
'    strSQL = strSQL & " WHERE key_cd = '" & strKeyCode & "'" & vbCrLf
'    strSQL = strSQL & " AND jigyou_cd = '" & strJigCd & "'"
'
'    'CDBAccessインスタンス生成
'    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
'
'    'SQL実行
'        objRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
'********** 2005/05/18 土田 削除 **********↑
    

'********** 2005/05/18 土田 変更 **********↓
    With pHChkJHTYKaisiDateRS
        .Filter = "key_cd = '" & strKeyCode & "' AND jigyou_cd = '" & strJigCD & "'"
    
    
'        If objRS.EOF = True Then
        If .EOF = True Then
            FncCheckDB_GetKaisiDateMas = False
'            GoTo Proc_End:
            Exit Function
        Else    '発注開始日データがNULLor空白である場合は、エラー
'条件分岐を一つにまとめた
'            If IsNull(Trim(objRS.Fields("hattyuu_kaisi_date"))) = True Then
'                FncCheckDB_GetKaisiDateMas = False
'                GoTo Proc_End:
'            ElseIf Trim(objRS.Fields("hattyuu_kaisi_date")) = "" Then
'                FncCheckDB_GetKaisiDateMas = False
'                GoTo Proc_End:
            If IsNull(Trim(.Fields("hattyuu_kaisi_date"))) Or _
                Trim(.Fields("hattyuu_kaisi_date")) = "" Then
                FncCheckDB_GetKaisiDateMas = False
    '            GoTo Proc_End:
                Exit Function
            Else
                ' 事業別情報テーブルから発注開始日を取得
'                strKaisiDate = Trim(objRS.Fields("hattyuu_kaisi_date"))     'char型の為、そのまま入れる。
                strKaisiDate = Trim(.Fields("hattyuu_kaisi_date"))     'char型の為、そのまま入れる。
            End If
        End If
    End With
'********** 2005/05/18 土田 変更 **********↑
    
    FncCheckDB_GetKaisiDateMas = True
    

'********** 2005/05/18 土田 削除 **********↓
'    'DBクローズ
'    objRS.Close
'        objRS = Nothing
'    objCDBAccess = Nothing
'
'    'GetObjectContext.SetComplete
'********** 2005/05/18 土田 削除 **********↑
Exit Function

'********** 2005/05/18 土田 削除 **********↓
'Proc_End:
'    If Not objRS Is Nothing Then
'            objRS = Nothing
'    End If
'    If Not objCDBAccess Is Nothing Then
'        objCDBAccess = Nothing
'    End If
'    Exit Function
'********** 2005/05/18 土田 削除 **********↑
DefectErrorHandler:
'********** 2005/05/18 土田 削除 **********↓
'    If Not objRS Is Nothing Then
'            objRS = Nothing
'    End If
'    If Not objCDBAccess Is Nothing Then
'        objCDBAccess = Nothing
'    End If
'    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
'********** 2005/05/18 土田 削除 **********↑
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーですa。", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!     FncCheckDB_ChushiKbn
'*【処理概要】
'*?     発注中止理由区分が正しいか確認する。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P     strChushiKbn ; 発注中止区分を表す文字列
'*P     intNowLine ; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R     FncCheckDB_ChushiKbn ; Boolean チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/25; 浅野　正徳(USC); 新規作成;
'*M    2005/05/12; 土田　祐歩(USC); パフォーマンス向上の為､この関数内での切断レコード取得をやめ大域変数の切断レコードを利用する様修正;
'*< Comment End >**************************************************'
'
Private Function FncCheckDB_ChushiKbn(strChushiKbn As String, intNowLine As Integer) As Boolean

    '　発注中止理由区分がWF名称マスタテーブルに登録されているか確認する。
    '   　DB：      WF名称マスタ（m_wf_meisyou）
    '   　キー：        KEY区分（key_kbn）                          「M08」
    '           KEYコード（key_cd）                         CSVファイルから取得
    '
    '
'********** 2005/05/12土田 削除 **********↓
'この関数内での切断レコードセット取得処理削除
'    Dim strSql            As String
'
'    strSql = "SELECT key_cd "
'    strSql = strSql & " FROM m_wf_meisyou WITH(readcommitted)"
'    strSql = strSql & " WHERE key_kbn = '" & KEYCD_HATTYUU_TYUUSI_RIYU_KBN & "'"
'    strSql = strSql & " AND key_cd = '" & strChushiKbn & "'"
'
'    Dim IsOKChushiKbn As Boolean
'
'    IsOKChushiKbn = FncCheckDB_IsRecordExists("VSWF", strSql)
'
'    If IsOKChushiKbn = False Then
'        FncSetError (intNowLine, enmHacColNo.ChushiKbn, PARAM_NOTFINDMASTER
'    End If
'
'    FncCheckDB_ChushiKbn = IsOKChushiKbn
'********** 2005/05/12土田 削除 **********↑

'********** 2005/05/12土田 追加 **********↓
'大域変数の切断レコードセットを使用しエラーチェック
    With pHChkHacchuRS
'********** 2005/05/17 土田 変更 **********↓
'********** 2005/05/18 土田 変更 **********↓
'        If Not .EOF Then
       .Filter = adFilterNone
        If Not .BOF Then
'********** 2005/05/18 土田 変更 **********↑
            .MoveFirst
                .Find("key_cd = '" & strChushiKbn & "'")
        End If
'********** 2005/05/17 土田 変更 **********↑
    End With
    
    If pHChkHacchuRS.EOF Then
            FncSetError(intNowLine, enmHacColNo.ChushiKbn, PARAM_NOTFINDMASTER)
    End If

    FncCheckDB_ChushiKbn = Not pHChkHacchuRS.EOF
'********** 2005/05/12土田 追加 **********↑


End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!     FncCheckDB_DaiJANCode
'*【処理概要】
'*?     代替JANコードが有効かを確認する。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P     strDaiJANCode ; String ; 代替JANコードを表す文字列
'*P     intNowLine As Integer ; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R     FncCheckDB_DaiJANCode ; Boolean ;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/25; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncCheckDB_DaiJANCode(strDaiJANCode As String, intNowLine As Integer) As Boolean

    '　マスタ登録されているかは、統合マスタのKEY情報1テーブルで確認する。
    '   　DB：      統合マスタ KEY情報1（v_key_1）
    '   　キー：        JANコード（jan_cd）                         CSVファイルから取得
    
    Dim strKeyCode As String
    If FncCheckDB_EntryJANCode(strDaiJANCode, strKeyCode) = True Then
        FncCheckDB_DaiJANCode = True
    Else
            FncSetError(intNowLine, enmHacColNo.DaiJanCd, PARAM_NOTEXISTDB, strDaiJANCode)
        FncCheckDB_DaiJANCode = False
    End If
    
End Function

'*******************************************************************
' 発注中止日用チェックルーチンここまで
'*******************************************************************

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    FncHacTourokuAllJig

'*【処理概要】
'*?    事業所マスタから事業コードを取得して、事業コード毎にDBに取り込む

'*【パラメータ】
'*P    strErrMsg;        string; エラーメッセージ;
'*P    intN;             int; 何番目のレコードか;

'*【戻り値】
'*R    Boolean; TRUE:成功、FALSE:失敗;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/25; 黒沢(ユニシステム株式会社); 新規作成;
'*< Comment End >**************************************************'
Private Function FncHacTourokuAllJig(ByRef strErrMsg, ByVal intN) As Boolean
    Dim sql, objCDB, intSelNum, blnRet
    Dim rsSel As ADODB.Recordset
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTourokuAllJig(" & CStr(intN) + 1 & "行目：【事業コード：" & ALL_JIG_CD & "】)"
    
    FncHacTourokuAllJig = False

    '事業所マスタから事業コードを取得する。
        objCDB = CreateObject("COMDBUtil.CDBAccess")    'CDBAccessをインスタンス化

    'SQL文作成（条件：閉店日がNULLor一括登録日以降、KEY区分=W02、設定内容の４桁目が1）
    sql = "SELECT DISTINCT A.jigyou_cd AS jigcd " & vbCrLf
    sql = sql & "FROM v_jigyousyo AS A WITH(readcommitted) " & vbCrLf
    sql = sql & "INNER JOIN m_wf_meisyou AS B WITH(readcommitted) ON A.jigyou_cd = B.key_cd " & vbCrLf
    sql = sql & "WHERE " & vbCrLf
    sql = sql & " Substring(B.settei_naiyou,4,1) = '1' AND " & vbCrLf
    sql = sql & " B.key_kbn = '" & KEYCD_KIKAKU_DOUNYU_TAISYOU & "' " & vbCrLf
    
        rsSel = objCDB.GetDCRecordset("VSWF", sql)
        objCDB = Nothing
    If rsSel.EOF = True Then    '対象の事業コードが無い場合はエラー
        strErrMsg = CStr(intN + 1) & "行目：【事業コード：" & ALL_JIG_CD & "】展開している事業コードは１件もありませんでした。"
        rsSel.Close
            rsSel = Nothing
        Exit Function
    End If
    
    '事業ｺｰﾄﾞ毎にデータを商品選定WFのDBに取り込む（FncHacTourokuJig関数を呼ぶ）。
    Do Until rsSel.EOF
        blnRet = FncHacTourokuJig(strErrMsg, intN, rsSel.Fields("jigcd"))
        If blnRet = False Then
            rsSel.Close
                rsSel = Nothing
            Exit Function
        End If
        rsSel.MoveNext
    Loop
    
    FncHacTourokuAllJig = True
    
    'DBクローズ
    rsSel.Close
        rsSel = Nothing
    
    'GetObjectContext.SetComplete
Exit Function

DefectErrorHandler:
    If Not rsSel Is Nothing Then
            rsSel = Nothing
    End If
    If Not objCDB Is Nothing Then
            objCDB = Nothing
    End If
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    FncHacTourokuJig

'*【処理概要】
'*?    指定した事業コードのデータを、DBに取り込む

'*【パラメータ】
'*P    strErrMsg;        string; エラーメッセージ;
'*P    intN;             int; 何番目のレコードか;
'*P    strJigCd;         string; 事業コード;

'*【戻り値】
'*R    Boolean; TRUE:成功、FALSE:失敗;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/25; 黒沢(ユニシステム株式会社); 新規作成;
'*< Comment End >**************************************************'
Private Function FncHacTourokuJig(ByRef strErrMsg, ByVal intN, ByVal strJigCD) As Boolean
    Dim sql, objCDB, intNum, blnRet
    Dim strSelSql, strInsSql, strUpdSql
    Dim arrParam() As Object
    Dim lngP_Cnt
    Dim rsSel
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTourokuJig(" & CStr(intN) + 1 & "行目：【事業コード：" & strJigCD & "】)"
    
    FncHacTourokuJig = False

    '指定された事業コードのデータを追加or更新する。
        objCDB = CreateObject("COMDBUtil.CDBAccess")    'CDBAccessをインスタンス化

    lngP_Cnt = 0
    Erase arrParam
    'SQL文作成（条件：KEY区分=W02、閉店日がNULLor一括登録日以降、設定内容の４桁目が1）
    Call SubHacSqlSelJig(intN, strJigCD, strSelSql, arrParam, lngP_Cnt)
    Call SubHacSqlInsJig(intN, strJigCD, strInsSql, arrParam, lngP_Cnt)
    Call SubHacSqlUpdJig(intN, strJigCD, strUpdSql, arrParam, lngP_Cnt)
    sql = "If (" & strSelSql & ") = 0 " & vbCrLf        'データが既に存在するかチェック
    sql = sql & "BEGIN " & strInsSql & " END " & vbCrLf  '存在しない場合は、追加
    sql = sql & "ELSE " & vbCrLf
    sql = sql & "BEGIN " & strUpdSql & " END " & vbCrLf  '存在する場合は、更新
    
    intNum = objCDB.ExecuteSql("VSWF", sql, arrParam)
    If intNum <= 0 Then    '対象の事業コードの登録に失敗した場合、エラー
        strErrMsg = CStr(intN + 1) & "行目：【事業コード：" & strJigCD & "】のデータの登録に失敗しました。"
            objCDB = Nothing
        Exit Function
    End If
    
    FncHacTourokuJig = True
    
    'DBクローズ
        objCDB = Nothing
    
    'GetObjectContext.SetComplete
Exit Function

DefectErrorHandler:
    If Not rsSel Is Nothing Then
            rsSel = Nothing
    End If
    If Not objCDB Is Nothing Then
            objCDB = Nothing
    End If
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    FncTenTouroku

'*【処理概要】
'*?    配列に格納された店別情報のデータをDBに取り込む

'*【パラメータ】
'*P    strErrMsg;        string; エラーメッセージ;

'*【戻り値】
'*R    Boolean; TRUE:成功、FALSE:失敗;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/20; 黒沢(ユニシステム株式会社); 新規作成;
'*M    2005/05/11; 土田　祐歩(USC); パフォーマンス向上の為､DB参照エラーチェック用切断レコードを用意する処理追加;
'*M    2005/05/13; 田中　禎人(USC); パフォーマンス向上対策（追加）;
'*M    2006/09/13; 和田(IC);        商品選定ＷＦ問合せ撲滅改善;
'*< Comment End >**************************************************'
Private Function FncTenTouroku(ByRef strErrMsg) As Boolean
    Dim intN As Integer
    Dim blnRet As Boolean
    Dim strTenCd As String
    Dim intNowLine As Integer
    Dim strCSVRecord() As String
    Dim flgChkError As Boolean      'チェック時のエラーフラグ（TRUE:エラーあり）
    
    FncTenTouroku = False
    
    '全ての配列データのチェックを行う。
    flgChkError = False
    intNowLine = 0
    strErrMsg = ""
    
'********** 2005/05/13 田中 追加 **********↓
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(該当事業コードチェック)"
'********** 2005/05/13 田中 追加 **********↑
    
    
'********** 2005/05/11 土田 追加 **********↓
'DB参照エラーチェック用切断レコードセットの生成
    Dim strSQL          As String
    Dim objCDBAccess    As Object
        
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    '一括登録時、毎回一つユーザーのみ実行できる
    Call LockTableIkkatu(objCDBAccess)
    
    Call SetIkkatuLog("1-1", "店別情報のデータをDBに取り込む", "FncTenTouroku", "", "", pstrTantousyaId, "INS")
    
    
    Call SetIkkatuLog("1-1-1", "該当事業コードチェック用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")
'該当事業コードチェック用切断レコードセット生成
    'SQL文作成（条件：判別区分=5、閉店日がNULLor一括登録日以降、KEY区分=W02、設定内容の4桁目が1、店コード）
    strSQL = "SELECT A.jigyousyo_cd, A.jigyou_cd, "
    strSQL = strSQL & "C.baika_jyogai_kbn,C.genka_jyogai_kbn,C.hattyuu_tani_jyogai_kbn," & vbCrLf
    strSQL = strSQL & "C.torihikisaki_jyogai_kbn,C.center_jyogai_kbn" & vbCrLf
    strSQL = strSQL & "FROM v_jigyousyo AS A WITH(readcommitted) " & vbCrLf
    strSQL = strSQL & "INNER JOIN m_wf_meisyou AS B WITH(readcommitted) ON A.jigyou_cd = B.key_cd " & vbCrLf
    
    '---------------------------------
    strSQL = strSQL & "     LEFT OUTER JOIN t_syousen_jyogai_tenpo C WITH (READCOMMITTED) ON                                               " & vbCrLf
    strSQL = strSQL & "     A.jigyousyo_cd =C.ten_cd AND " & vbCrLf
    strSQL = strSQL & "     C.kikaku_syurui_cd = '" & pstrKikakuSyurui & "' AND " & vbCrLf
    strSQL = strSQL & "     C.nen_ren_no = '" & pstrKikakuNenrenban & "' AND                           " & vbCrLf
    strSQL = strSQL & "     C.eda_no = '" & pstrKikakuEdaban & "'         " & vbCrLf
    
    '---------------------------------
    
    strSQL = strSQL & "WHERE A.hanbetu_kbn = '" & HANBETU_KBN_TENPO & "' AND " & vbCrLf
    strSQL = strSQL & "(RTrim(A.heiten_date) = '' OR A.heiten_date is null OR A.heiten_date >= " & FncGetNowDate() & ") " & vbCrLf
    strSQL = strSQL & "AND Substring(B.settei_naiyou,4,1) = '1' " & vbCrLf
    strSQL = strSQL & "AND B.key_kbn = '" & KEYCD_KIKAKU_DOUNYU_TAISYOU & "' " & vbCrLf
    
    
    
    '該当事業コードチェック用　切断レコードセット取得
        pTChkJigyouRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkJigyouRS.Fields("jigyousyo_cd").Properties("Optimize") = True 'オプティマイザ
    
'展開店コードリストの生成
    
    With pTChkJigyouRS
'********** 2005/05/17 土田 変更 **********↓
'********** 2005/05/18 土田 変更 **********↓
'        If Not .EOF Then
        If Not .BOF Then
'********** 2005/05/18 土田 変更 **********↑
            .MoveFirst
            Do Until .EOF
                pTTenCodeList = pTTenCodeList & "'" & .Fields("jigyousyo_cd") & "',"
                .MoveNext
            Loop
            .MoveFirst
        End If
'********** 2005/05/17 土田 変更 **********↑
    End With
    
    pTTenCodeList = Left(pTTenCodeList, Len(pTTenCodeList) - 1)

        Call SetIkkatuLog("1-1-1", "", "", "", "OK", "", "UPD")



'********** 2005/05/19 土田 削除 **********↓
''********** 2005/05/19 土田 追加 **********↓
''統合マスタ登録済みJANコード紐付け主キーのリスト(主キー重複無し)作成
'    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(CSVファイル内JANコード紐付け主キーのリスト(主キー重複無し)作成)"
'    With pChkvk1JANCodeRS
'        Do Until .EOF
'            If InStr(pTKeyCDList, .Fields("key_cd")) = 0 Then
'                pTKeyCDList = pTKeyCDList & "'" & .Fields("key_cd") & "',"
'            End If
'            .MoveNext
'        Loop
'
'        If pTKeyCDList <> "" Then
'            pTKeyCDList = Left(pTKeyCDList, Len(pTKeyCDList) - 1)
'        Else
'            pTKeyCDList = "''"
'        End If
'    End With
''********** 2005/05/19 土田 追加 **********↑
'********** 2005/05/19 土田 削除 **********↑


'売価変更理由チェック用切断レコードセット生成
 Call SetIkkatuLog("1-1-2", "売価変更理由チェック用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")
 
'********** 2005/05/13 田中 追加 **********↓
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(売価変更理由チェック)"
'********** 2005/05/13 田中 追加 **********↑
    
   'SQL文作成（条件：KEY区分=025、売価変更理由コード）
    strSQL = "SELECT key_cd FROM v_meisyou WITH(readcommitted) "
    strSQL = strSQL & "WHERE key_kbn = '" & KEYCD_BAIKARIYUU & "'"

        pTChkBaikaHenkouRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkBaikaHenkouRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ
    
Call SetIkkatuLog("1-1-2", "", "", "", "OK", "", "UPD")
'取引先チェック用切断レコードセット生成
 
  Call SetIkkatuLog("1-1-3", "取引先チェック用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")
'********** 2005/05/13 田中 追加 **********↓
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(取引先チェック)"
'********** 2005/05/13 田中 追加 **********↑
    
   strSQL = "SELECT torihikisaki_cd FROM v_torihikisaki WITH(readcommitted)"
    strSQL = strSQL & "WHERE torikesi_date LIKE ' ' OR torikesi_date IS NULL"
    
        pTChkTorihikisakiRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkTorihikisakiRS.Fields("torihikisaki_cd").Properties("Optimize") = True 'オプティマイザ
Call SetIkkatuLog("1-1-3", "", "", "", "OK", "", "UPD")
'センター区分チェック用切断レコードセット生成
 
'********** 2005/05/13 田中 追加 **********↓
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(センター区分チェック)"
'********** 2005/05/13 田中 追加 **********↑
  Call SetIkkatuLog("1-1-4", "センター区分チェック用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")
  
   strSQL = "SELECT key_cd FROM m_wf_meisyou WITH(readcommitted) "
    strSQL = strSQL & "WHERE key_kbn = '" & KEYCD_CENTER_KBN & "' "
    
        pTChkCenterRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkCenterRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ
Call SetIkkatuLog("1-1-4", "", "", "", "OK", "", "UPD")
'********** 2005/05/11 土田 追加 **********↑


'********** 2005/05/13 田中 追加 **********↓
'先付変更情報の変更日チェック用切断レコードセット生成（統合マスタ・売価変更情報）
  Call SetIkkatuLog("1-1-5", "先付変更情報の変更日（統合マスタ・売価変更情報）", "SQL", "", "", pstrTantousyaId, "INS")
  
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(先付変更情報の変更日（統合マスタ・売価変更情報）)"
    strSQL = "SELECT key_cd, ten_cd, henkou_date "
    strSQL = strSQL & "FROM v_baika_henkou WITH(readcommitted) "
    'strSQL = strSQL & "where 1=0 "    '▼テスト
'********** 2005/05/19 土田 追加 後 削除**********↓
'    strSQL = strSQL & " WHERE key_cd IN(" & pTKeyCDList & ") "
'********** 2005/05/19 土田 追加 後 削除**********↓
    
        pTChkMstBaikaRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkMstBaikaRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ
    pTChkMstBaikaRS.Fields("henkou_date").Properties("Optimize") = True
Call SetIkkatuLog("1-1-5", "", "", "", "OK", "", "UPD")
  Call SetIkkatuLog("1-1-6", "先付変更情報の変更日チェック用切断レコードセット生成（統合マスタ・仕入変更情報）", "SQL", "", "", pstrTantousyaId, "INS")
  
'先付変更情報の変更日チェック用切断レコードセット生成（統合マスタ・仕入変更情報）
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(先付変更情報の変更日（統合マスタ・仕入変更情報）)"
    strSQL = "SELECT key_cd, ten_cd, henkou_date "
    strSQL = strSQL & "FROM v_siire_henkou WITH(readcommitted) "
    'strSQL = strSQL & "where 1=0 "    '▼テスト
'********** 2005/05/19 土田 追加 後 削除**********↓
'    strSQL = strSQL & " WHERE key_cd IN(" & pTKeyCDList & ") "
'********** 2005/05/19 土田 追加 後 削除**********↓

        pTChkMstSiireRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkMstSiireRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ
    pTChkMstSiireRS.Fields("henkou_date").Properties("Optimize") = True
Call SetIkkatuLog("1-1-6", "", "", "", "OK", "", "UPD")

  Call SetIkkatuLog("1-1-7", "先付変更情報の変更日（統合マスタ・センターコード変更情報）", "SQL", "", "", pstrTantousyaId, "INS")
  
'先付変更情報の変更日チェック用切断レコードセット生成（統合マスタ・センターコード変更情報）
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(先付変更情報の変更日（統合マスタ・センターコード変更情報）)"
    strSQL = "SELECT key_cd, ten_cd, henkou_date "
    strSQL = strSQL & "FROM v_tenpo_center_cd_henkou WITH(readcommitted) "
    'strSQL = strSQL & "where 1=0 "    '▼テスト
'********** 2005/05/19 土田 追加 後 削除**********↓
'    strSQL = strSQL & " WHERE key_cd IN(" & pTKeyCDList & ") "
'********** 2005/05/19 土田 追加 後 削除**********↓

        pTChkMstCenterRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkMstCenterRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ
    pTChkMstCenterRS.Fields("henkou_date").Properties("Optimize") = True
Call SetIkkatuLog("1-1-7", "", "", "", "OK", "", "UPD")
'********** 2005/05/13 田中 追加 **********↑

'********** 2005/05/18 土田 追加 **********↓
'CSVファイル内最大、最小日付の初期化
    pstrMinDate = paryCsvData(intNowLine)(enmTenColNo.HenkouDate)
    pstrMaxDate = paryCsvData(intNowLine)(enmTenColNo.HenkouDate)
'********** 2005/05/18 土田 追加 **********↑


'********** 2005/05/19 土田 移動 **********↓
'重複JANコードチェックレコードセットの排他ロックの時間を短くする為移動

    '重複JANコードチェック用切断レコードセット生成
    
    '********** 2005/05/13 田中 追加 **********↓
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(重複JANコードチェック)"
    '********** 2005/05/13 田中 追加 **********↑
        
'        strSQL = "SELECT jan_cd FROM t_syousen_wf_kanri WITH(readcommitted) WHERE "
'        strSQL = "SELECT jan_cd FROM t_syousen_wf_kanri WITH(UPDLOCK) WHERE "
'        strSQL = strSQL & " NOT("
'        strSQL = strSQL & "kikaku_syurui_cd = '" & pstrKikakuSyurui & "' "
'        strSQL = strSQL & " AND nen_ren_no = '" & pstrKikakuNenrenban & "' "
'        strSQL = strSQL & " AND eda_no = '" & pstrKikakuEdaban & "'"
'        strSQL = strSQL & ") "
'        strSQL = strSQL & ")  SET LOCK_TIMEOUT -1"
'
'            pChkSameJANCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
'        pChkSameJANCodeRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
'********** 2005/05/20土田 変更 **********↑

'********** 2005/09/12 高附 変更 **********↓
  Call SetIkkatuLog("1-1-8", "重複JANコードチェック用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")
  
    '重複JANコードチェック用切断レコードセット生成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(重複JANコードチェック用切断レコードセット生成)"
    strSQL = "SELECT twk.jan_cd, mk1.key_cd, twk.kikaku_no" & vbCrLf
    strSQL = strSQL & " FROM t_syousen_wf_kanri as twk WITH(readcommitted)" & vbCrLf
    strSQL = strSQL & "      INNER JOIN v_key_1 as mk1 WITH(readcommitted)" & vbCrLf
    strSQL = strSQL & "              ON(twk.jan_cd = mk1.jan_cd)" & vbCrLf
    strSQL = strSQL & " WHERE" & vbCrLf
    strSQL = strSQL & " NOT(" & vbCrLf
    strSQL = strSQL & "      kikaku_syurui_cd = '" & pstrKikakuSyurui & "' " & vbCrLf
    strSQL = strSQL & "      AND nen_ren_no = '" & pstrKikakuNenrenban & "' " & vbCrLf
    strSQL = strSQL & "      AND eda_no = '" & pstrKikakuEdaban & "'" & vbCrLf
    strSQL = strSQL & "    )" & vbCrLf
    strSQL = strSQL & " AND twk.touroku_kbn = '2'" & vbCrLf
    
        pChkSameJANCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pChkSameJANCodeRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    pChkSameJANCodeRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ
    
'********** 2005/09/12 高附 変更 **********↑
    
    '********** 2005/05/18 土田 追加 **********↓
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(CSVファイル内Janコードのリスト(Janコードの重複無し)作成)"
    'CSVファイル内Janコードのリスト(Janコードの重複無し)を作成する
        intNowLine = 0
        Do Until UBound(paryCsvData) < intNowLine
            
            strCSVRecord = paryCsvData(intNowLine)
             
            If InStr(pCsvJanCodeList, "'" & strCSVRecord(enmTenColNo.janCD) & "'") = 0 Then
                pCsvJanCodeList = pCsvJanCodeList & "'" & strCSVRecord(enmTenColNo.janCD) & "',"
            End If
    
            intNowLine = intNowLine + 1
        Loop
        
        If pCsvJanCodeList <> "" Then
            pCsvJanCodeList = Left(pCsvJanCodeList, Len(pCsvJanCodeList) - 1)
        Else
            pCsvJanCodeList = "''"
        End If
        
Call SetIkkatuLog("1-1-8", "", "", "", "OK", "", "UPD")

Call SetIkkatuLog("1-1-9", "統合マスタ登録済みJANコードチェック", "SQL", "", "", pstrTantousyaId, "INS")
        intNowLine = 0
    '********** 2005/05/18 土田 追加 **********↑
        
    '統合マスタ登録済みJANコードチェック用切断レコードセット生成
    '********** 2005/05/13 田中 追加 **********↓
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(統合マスタ登録済みJANコードチェック)"
    '********** 2005/05/13 田中 追加 **********↑
        
'********** 2005/05/20土田 変更 **********↓
'       strSQL = "SELECT jan_cd, key_cd  FROM v_key_1 WITH(readcommitted)"
       strSQL = "SELECT jan_cd, key_cd  FROM v_key_1 WITH(UPDLOCK)"
'********** 2005/05/20土田 変更 **********↑
    '********** 2005/05/18 土田 追加 **********↓
        strSQL = strSQL & " WHERE jan_cd IN(" & pCsvJanCodeList & ") "
'********** 2005/05/20土田 追加 **********↓
        strSQL = strSQL & " SET LOCK_TIMEOUT -1"
'********** 2005/05/20土田 追加 **********↑
    '********** 2005/05/18 土田 追加 **********↑
        
            pChkvk1JANCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
        pChkvk1JANCodeRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ

'↑重複JANコードチェックレコードセットの排他ロックの時間を短くする為移動
'********** 2005/05/19 土田 移動 **********↑


    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(エラーチェック)"
    
    Do Until UBound(paryCsvData) < intNowLine
        
        strCSVRecord = paryCsvData(intNowLine)
        
        '(1)DBに依存しないチェック：FncCheckParameter_CorrectForm
        '(2)DBを参照してチェック:FncCheckParameter_DataBase
'********** 2005/06/17 北澤 **********↓
'        If FncCheckParameter_CorrectForm(strCSVRecord, intNowLine + 1) = False Or _
'           FncCheckParameter_DataBase(strCSVRecord, intNowLine + 1) = False Then
'            flgChkError = True
'        End If
'アベンド回避の為｢DBに依存しないチェック｣でエラーが無い場合のみ｢DBを参照してチェック｣を行う
        If FncCheckParameter_CorrectForm(strCSVRecord, intNowLine + 1) = False Then
            flgChkError = True
        Else
            If FncCheckParameter_DataBase(strCSVRecord, intNowLine + 1) = False Then
                flgChkError = True
            End If
        End If
'********** 2005/06/17 北澤 **********↑
        
'********** 2005/05/17 土田 追加 **********↓
'CSVファイル内Janコードのリスト(Janコードの重複無し)を作成する
'        If InStr(pTCsvJanCodeList, strCSVRecord(enmTenColNo.JanCd)) = 0 Then
'            pTCsvJanCodeList = pTCsvJanCodeList & "'" & strCSVRecord(enmTenColNo.JanCd) & "',"
'        End If
'********** 2005/05/17 土田 追加 **********↑

        intNowLine = intNowLine + 1
    Loop
Call SetIkkatuLog("1-1-9", "", "", "", "OK", "", "UPD")
'********** 2005/05/17 土田 追加 **********↓
'CSVファイル内Janコードのリスト(Janコードの重複無し)を作成する
'    If pTCsvJanCodeList <> "" Then
'        pTCsvJanCodeList = Left(pTCsvJanCodeList, Len(pTCsvJanCodeList) - 1)
'    End If
'
'********** 2005/05/17 土田 追加 **********↑
    
'********** 2005/05/11 土田 追加 **********↓
'DB参照エラーチェック用切断レコードセットの開放
    
'********** 2005/05/11 土田 削除 **********↓
'データ登録時利用するので開放しない
'        pTChkJigyouRS.Close
'            pTChkJigyouRS = Nothing
'********** 2005/05/11 土田 削除 **********↑
        
'********** 2005/05/19 土田 削除 **********↓
'データ登録を済ませるまで排他ロックをかけるのでここでは開放しない
'        pChkSameJANCodeRS.Close
'            pChkSameJANCodeRS = Nothing
'
'********** 2005/05/19 土田 削除 **********↑
       
'********** 2005/05/17 土田 削除 **********↓
'データ登録時利用するので開放しない
'        pChkvk1JANCodeRS.Close
'            pChkvk1JANCodeRS = Nothing
'********** 2005/05/17 土田 削除 **********↑
       
        pTChkBaikaHenkouRS.Close
            pTChkBaikaHenkouRS = Nothing
       
        pTChkTorihikisakiRS.Close
            pTChkTorihikisakiRS = Nothing
       
        pTChkCenterRS.Close
            pTChkCenterRS = Nothing
       
'********** 2005/05/13 田中 追加 **********↓
'DB参照エラーチェック用切断レコードセットの開放（追加）
        pTChkMstBaikaRS.Close
            pTChkMstBaikaRS = Nothing
       
        pTChkMstSiireRS.Close
            pTChkMstSiireRS = Nothing
       
        pTChkMstCenterRS.Close
            pTChkMstCenterRS = Nothing
   
'********** 2005/05/13 田中 追加 **********↑

        'DBクローズ
        objCDBAccess = Nothing
        'GetObjectContext.SetComplete
    
'********** 2005/05/11 土田 追加 **********↑
   
   
    'データに１つでも問題がある場合は、ここで終了！！！
    If flgChkError = True Then
        'エラーがあればエラーメッセージを取得する。（終了しない）
        strErrMsg = FncGetErrorMsg()
        Exit Function
    End If
     
     'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(データ登録)"
        Call SetIkkatuLog("1-2", "データ登録", "FncTenTouroku", "", "", pstrTantousyaId, "INS")
'********** 2005/05/16 土田 仮追加 **********↓
    subIns_SetAllTables (paryCsvData)
'********** 2005/05/16 土田 仮追加 **********↑
        Call SetIkkatuLog("1-2", "", "", "", "OK", "", "UPD")

'********** 2005/05/20 土田 追加 **********↓
'排他ロック解除
        pChkSameJANCodeRS.Close
            pChkSameJANCodeRS = Nothing
        objCDBAccess = Nothing
        'GetObjectContext.SetComplete
        
        Call SetIkkatuLog("1-1", "", "", "", "OK", "", "UPD")
                
'********** 2005/05/20 土田 追加 **********↑
    
      'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(データ登録終了)"

'********** 2005/05/16 土田 仮削除 **********↓
'    '配列データをDBに取り込む
'    intN = 0
''********** 2005/04/25 黒沢 追加 **********↓
'    pflgJigJyoToukbn = False            '発注中止日用フラグも一応クリアする。（ここでは使用しない）
''********** 2005/04/25 黒沢 追加 **********↑
'
'
'    Do Until UBound(paryCsvData) < intN         '配列の終端までループを繰り返します。
'        '商品選定DB(WF管理項目)の先付変更の登録区分更新確認用フラグをクリア
'        pflgBaikaToukbn = False
'        pflgSiireToukbn = False
'        pflgCenterToukbn = False
'
'        strTenCd = paryCsvData(intN)(enmTenColNo.TenCd)      '該当店コードを取得
'
'        '該当店コード別に処理を分ける。
'        If strTenCd = ALL_TEN_CD Then           '全店
'            blnRet = FncTenTourokuJig(strErrMsg, intN, strTenCd)
'        ElseIf strTenCd = ALL_JIG_CD Then       '全事業
'            blnRet = FncTenTourokuJig(strErrMsg, intN, strTenCd)
'        ElseIf Len(strTenCd) = 2 Then           '事業
'            blnRet = FncTenTourokuJig(strErrMsg, intN, strTenCd)
'        ElseIf Len(strTenCd) = 4 Then           '店
'            blnRet = FncTenTourokuTen(strErrMsg, intN, strTenCd)
'        End If
'        If blnRet = False Then      'エラー処理
'            Exit Function
'        End If
'
''********** 2005/04/27 北澤 **********↓
''VSWF表示仕様テーブルにレコードが存在しない場合はVMST一般基本テーブルから情報を取得しレコードを追加する
''(照会などで商品名が表示されない為)
'        blnRet = Insert_t_syousen_hyouji_siyou(strErrMsg, intN _
'                , pstrKikakuSyurui, pstrKikakuNenrenban, pstrKikakuEdaban _
'                , paryCsvData(intN)(enmTenColNo.JanCd) _
'                , pstrTantousyaId)
'        If blnRet = False Then
'            Exit Function
'        End If
''********** 2005/04/27 北澤 **********↑
'
'        '商品選定DB(WF管理項目)の更新
''********** 2005/04/26 黒沢 変更 **********↓
'        If FncSetWFKanri(strErrMsg, intN, RET_FLD_TENB) = False Then
''        If FncUpdWFKanri(strErrMsg, intN) = False Then
''********** 2005/04/26 黒沢 変更 **********↑
'            Exit Function           'エラー処理
'        End If
'
'        intN = intN + 1
'    Loop
'
'********** 2005/05/16 土田 仮削除 **********↑

    FncTenTouroku = True
    
'********** 2005/05/13 田中 追加 **********↓

Exit Function

DefectErrorHandler:
    
    If Not pTChkJigyouRS Is Nothing Then
            pTChkJigyouRS = Nothing
    End If
    If Not pChkSameJANCodeRS Is Nothing Then
            pChkSameJANCodeRS = Nothing
    End If
    If Not pChkvk1JANCodeRS Is Nothing Then
            pChkvk1JANCodeRS = Nothing
    End If
    If Not pTChkBaikaHenkouRS Is Nothing Then
            pTChkBaikaHenkouRS = Nothing
    End If
    If Not pTChkTorihikisakiRS Is Nothing Then
            pTChkTorihikisakiRS = Nothing
    End If
    If Not pTChkCenterRS Is Nothing Then
            pTChkCenterRS = Nothing
    End If
    If Not pTChkMstBaikaRS Is Nothing Then
            pTChkMstBaikaRS = Nothing
    End If
    If Not pTChkMstSiireRS Is Nothing Then
            pTChkMstSiireRS = Nothing
    End If
    If Not pTChkMstCenterRS Is Nothing Then
            pTChkMstCenterRS = Nothing
    End If
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
    
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    Call SetIkkatuLog("1-1", "", "", Left(Err.Description, 400), "NG", "", "UPD")
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description
'********** 2005/05/13 田中 追加 **********↑

End Function

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    FncModosiKakaku

'*【処理概要】
'*?    配列に格納された店別情報のデータをDBに取り込む

'*【パラメータ】
'*P    strErrMsg;        string; エラーメッセージ;

'*【戻り値】
'*R    Boolean; TRUE:成功、FALSE:失敗;

'******************************************************************'
'*【変更履歴】
'*M    2009/06/06; Supattra T.(Tostem Thai); 新規作成;
'*< Comment End >**************************************************'
Private Function FncModosiKakaku(ByRef strErrMsg) As Boolean
    Dim intN As Integer
    Dim blnRet As Boolean
    Dim strTenCd As String
    Dim intNowLine As Integer
    Dim strCSVRecord() As String
    Dim paryCheckData() As Object  '「変更日」、「JANコード」が同一の行単位でチェックするレコード
    Dim flgChkError As Boolean      'チェック時のエラーフラグ（TRUE:エラーあり）
    
    FncModosiKakaku = False
    
    '全ての配列データのチェックを行う。
    flgChkError = False
    intNowLine = 0
    strErrMsg = ""
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(該当事業コードチェック)"
    
    Dim strSQL          As String
    Dim objCDBAccess    As Object
        
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    '一括登録時、毎回一つユーザーのみ実行できる
    Call LockTableIkkatu(objCDBAccess)
    
    Call SetIkkatuLog("1-1", "店別情報のデータをDBに取り込む", "FncModosiKakaku", "", "", pstrTantousyaId, "INS")
    Call SetIkkatuLog("1-1-1", "該当事業コードチェック用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")
    
    '該当事業コードチェック用切断レコードセット生成
    'SQL文作成（条件：判別区分=5、閉店日がNULLor一括登録日以降、KEY区分=W02、設定内容の4桁目が1、店コード）
    strSQL = "SELECT A.jigyousyo_cd, A.jigyou_cd, "
    strSQL = strSQL & "C.baika_jyogai_kbn,C.genka_jyogai_kbn,C.hattyuu_tani_jyogai_kbn," & vbCrLf
    strSQL = strSQL & "C.torihikisaki_jyogai_kbn,C.center_jyogai_kbn" & vbCrLf
    strSQL = strSQL & "FROM v_jigyousyo AS A WITH(readcommitted) " & vbCrLf
    strSQL = strSQL & "INNER JOIN m_wf_meisyou AS B WITH(readcommitted) ON A.jigyou_cd = B.key_cd " & vbCrLf
    
    '---------------------------------
    strSQL = strSQL & "     LEFT OUTER JOIN t_syousen_jyogai_tenpo C WITH (READCOMMITTED) ON                                               " & vbCrLf
    strSQL = strSQL & "     A.jigyousyo_cd =C.ten_cd AND " & vbCrLf
    strSQL = strSQL & "     C.kikaku_syurui_cd = '" & pstrKikakuSyurui & "' AND " & vbCrLf
    strSQL = strSQL & "     C.nen_ren_no = '" & pstrKikakuNenrenban & "' AND                           " & vbCrLf
    strSQL = strSQL & "     C.eda_no = '" & pstrKikakuEdaban & "'         " & vbCrLf
    
    '---------------------------------
    
    strSQL = strSQL & "WHERE A.hanbetu_kbn = '" & HANBETU_KBN_TENPO & "' AND " & vbCrLf
    strSQL = strSQL & "(RTrim(A.heiten_date) = '' OR A.heiten_date is null OR A.heiten_date >= " & FncGetNowDate() & ") " & vbCrLf
    strSQL = strSQL & "AND Substring(B.settei_naiyou,4,1) = '1' " & vbCrLf
    strSQL = strSQL & "AND B.key_kbn = '" & KEYCD_KIKAKU_DOUNYU_TAISYOU & "' " & vbCrLf
    
    
    
    '該当事業コードチェック用　切断レコードセット取得
        pTChkJigyouRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkJigyouRS.Fields("jigyousyo_cd").Properties("Optimize") = True 'オプティマイザ
    
'展開店コードリストの生成
    
    With pTChkJigyouRS
        If Not .BOF Then
            .MoveFirst
            Do Until .EOF
                pTTenCodeList = pTTenCodeList & "'" & .Fields("jigyousyo_cd") & "',"
                .MoveNext
            Loop
            .MoveFirst
        End If
    End With
    
    pTTenCodeList = Left(pTTenCodeList, Len(pTTenCodeList) - 1)
        Call SetIkkatuLog("1-1-1", "", "", "", "OK", "", "UPD")


        Call SetIkkatuLog("1-1-2", "売価変更理由チェック", "SQL", "", "", pstrTantousyaId, "INS")
    '売価変更理由チェック用切断レコードセット生成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(売価変更理由チェック)"
    
   'SQL文作成（条件：KEY区分=025、売価変更理由コード）
    strSQL = "SELECT key_cd FROM v_meisyou WITH(readcommitted) "
    strSQL = strSQL & "WHERE key_kbn = '" & KEYCD_BAIKARIYUU & "'"

        pTChkBaikaHenkouRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkBaikaHenkouRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ
        Call SetIkkatuLog("1-1-2", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-1-3", "取引先チェック", "SQL", "", "", pstrTantousyaId, "INS")
    '取引先チェック用切断レコードセット生成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(取引先チェック)"
    
    strSQL = "SELECT torihikisaki_cd FROM v_torihikisaki WITH(readcommitted)"
    strSQL = strSQL & "WHERE torikesi_date LIKE ' ' OR torikesi_date IS NULL"
    
        pTChkTorihikisakiRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkTorihikisakiRS.Fields("torihikisaki_cd").Properties("Optimize") = True 'オプティマイザ
        Call SetIkkatuLog("1-1-3", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-1-4", "センター区分チェック", "SQL", "", "", pstrTantousyaId, "INS")
   'センター区分チェック用切断レコードセット生成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(センター区分チェック)"
    
   strSQL = "SELECT key_cd FROM m_wf_meisyou WITH(readcommitted) "
    strSQL = strSQL & "WHERE key_kbn = '" & KEYCD_CENTER_KBN & "' "
    
        pTChkCenterRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkCenterRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ
        Call SetIkkatuLog("1-1-4", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-1-5", "先付変更情報の変更日（統合マスタ・売価変更情報）", "SQL", "", "", pstrTantousyaId, "INS")
   '先付変更情報の変更日チェック用切断レコードセット生成（統合マスタ・売価変更情報）
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(先付変更情報の変更日（統合マスタ・売価変更情報）)"
    'strSQL = "SELECT key_cd, ten_cd, henkou_date "
    strSQL = "SELECT key_cd, henkou_date, ten_cd "
    strSQL = strSQL & "FROM v_baika_henkou WITH(readcommitted) "
    
        pTChkMstBaikaRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkMstBaikaRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ
    'pTChkMstBaikaRS.Fields("henkou_date").Properties("Optimize") = True
        Call SetIkkatuLog("1-1-5", "", "", "", "OK", "", "UPD")
        Call SetIkkatuLog("1-1-6", "先付変更情報の変更日（統合マスタ・仕入変更情報)", "SQL", "", "", pstrTantousyaId, "INS")
   '先付変更情報の変更日チェック用切断レコードセット生成（統合マスタ・仕入変更情報）
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(先付変更情報の変更日（統合マスタ・仕入変更情報）)"
    'strSQL = "SELECT key_cd, ten_cd, henkou_date "
    strSQL = "SELECT key_cd, henkou_date, ten_cd "
    strSQL = strSQL & "FROM v_siire_henkou WITH(readcommitted) "

        pTChkMstSiireRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkMstSiireRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ
    'pTChkMstSiireRS.Fields("henkou_date").Properties("Optimize") = True
        Call SetIkkatuLog("1-1-6", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-1-7", "先付変更情報の変更日（統合マスタ・センターコード変更情報）", "SQL", "", "", pstrTantousyaId, "INS")
   '先付変更情報の変更日チェック用切断レコードセット生成（統合マスタ・センターコード変更情報）
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(先付変更情報の変更日（統合マスタ・センターコード変更情報）)"
    'strSQL = "SELECT key_cd, ten_cd, henkou_date "
    strSQL = "SELECT key_cd, henkou_date, ten_cd "
    strSQL = strSQL & "FROM v_tenpo_center_cd_henkou WITH(readcommitted) "

        pTChkMstCenterRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkMstCenterRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ
    'pTChkMstCenterRS.Fields("henkou_date").Properties("Optimize") = True

   'CSVファイル内最大、最小日付の初期化
    pstrMinDate = paryCsvData(intNowLine)(enmTenColNo.HenkouDate)
    pstrMaxDate = paryCsvData(intNowLine)(enmTenColNo.HenkouDate)
    Call SetIkkatuLog("1-1-7", "", "", "", "OK", "", "UPD")


        Call SetIkkatuLog("1-1-8", "複JANコードチェック", "SQL", "", "", pstrTantousyaId, "INS")
    '重複JANコードチェックレコードセットの排他ロックの時間を短くする為移動
    '重複JANコードチェック用切断レコードセット生成
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(重複JANコードチェック)"

    '重複JANコードチェック用切断レコードセット生成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(重複JANコードチェック用切断レコードセット生成)"
    strSQL = "SELECT twk.jan_cd, mk1.key_cd, twk.kikaku_no" & vbCrLf
    strSQL = strSQL & " FROM t_syousen_wf_kanri as twk WITH(readcommitted)" & vbCrLf
    strSQL = strSQL & "      INNER JOIN v_key_1 as mk1 WITH(readcommitted)" & vbCrLf
    strSQL = strSQL & "              ON(twk.jan_cd = mk1.jan_cd)" & vbCrLf
    strSQL = strSQL & " WHERE" & vbCrLf
    strSQL = strSQL & " NOT(" & vbCrLf
    strSQL = strSQL & "      kikaku_syurui_cd = '" & pstrKikakuSyurui & "' " & vbCrLf
    strSQL = strSQL & "      AND nen_ren_no = '" & pstrKikakuNenrenban & "' " & vbCrLf
    strSQL = strSQL & "      AND eda_no = '" & pstrKikakuEdaban & "'" & vbCrLf
    strSQL = strSQL & "    )" & vbCrLf
    strSQL = strSQL & " AND twk.touroku_kbn = '2'" & vbCrLf
    
        pChkSameJANCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pChkSameJANCodeRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    pChkSameJANCodeRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ
        
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(CSVファイル内Janコードのリスト(Janコードの重複無し)作成)"
    'CSVファイル内Janコードのリスト(Janコードの重複無し)を作成する
        intNowLine = 0
        Do Until UBound(paryCsvData) < intNowLine
            
            strCSVRecord = paryCsvData(intNowLine)
             
            If InStr(pCsvJanCodeList, "'" & strCSVRecord(enmTenColNo.janCD) & "'") = 0 Then
                pCsvJanCodeList = pCsvJanCodeList & "'" & strCSVRecord(enmTenColNo.janCD) & "',"
            End If
    
            intNowLine = intNowLine + 1
        Loop
        
        If pCsvJanCodeList <> "" Then
            pCsvJanCodeList = Left(pCsvJanCodeList, Len(pCsvJanCodeList) - 1)
        Else
            pCsvJanCodeList = "''"
        End If
        intNowLine = 0
Call SetIkkatuLog("1-1-8", "", "", "", "OK", "", "UPD")

Call SetIkkatuLog("1-1-9", "統合マスタ登録済みJANコードチェック", "SQL", "", "", pstrTantousyaId, "INS")
    '   統合マスタ登録済みJANコードチェック用切断レコードセット生成
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(統合マスタ登録済みJANコードチェック)"
        
       strSQL = "SELECT jan_cd, key_cd  FROM v_key_1 WITH(UPDLOCK)"
        strSQL = strSQL & " WHERE jan_cd IN(" & pCsvJanCodeList & ") "
        strSQL = strSQL & " SET LOCK_TIMEOUT -1"
        
            pChkvk1JANCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
        pChkvk1JANCodeRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ

    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(エラーチェック)"
    
    '********** 2012/08/03 楊双（大連）入力チェックの追加 **********↓
    Dim flgFirst As Boolean
    flgFirst = False
    '********** 2012/08/03 楊双（大連）入力チェックの追加 **********↑
    Do Until UBound(paryCsvData) < intNowLine
        
        strCSVRecord = paryCsvData(intNowLine)
        
        If FncCheckParameter_CorrectForm_Modosi(strCSVRecord, intNowLine + 1) = False Then
            flgChkError = True
        Else
            If FncCheckParameter_DataBase(strCSVRecord, intNowLine + 1) = False Then
                flgChkError = True
            End If
            '********** 2012/08/03 楊双（大連）入力チェックの追加 **********↓
            If FncCheckParameter_PDC_Modosi(strCSVRecord, intNowLine + 1, paryCheckData, flgFirst) = False Then
                flgChkError = True
            End If
            '********** 2012/08/03 楊双（大連）入力チェックの追加 **********↑
        End If

        intNowLine = intNowLine + 1
    Loop

    '********** 2012/08/03 楊双（大連）入力チェックの追加 **********↓

    Dim intLoopCnt As Integer
    
   If flgFirst = True Then
    For intLoopCnt = 0 To UBound(paryCheckData, 1)
        If InStr(paryCheckData(intLoopCnt)(4), "'0000'") = 0 Then
            'エラーの場合
                    FncSetErrorPDC(intNowLine, enmModColNo.TenCd, PARAM_HONBU_NAI, paryCheckData(intLoopCnt)(0) & "," & paryCheckData(intLoopCnt)(1))
            flgChkError = True
        End If
                                
    Next
    End If

    Dim aryAllLineInfo As Object
    Dim aryEachLineInfo As Object
    Dim flgFirstAllData As Boolean
    Dim flgSameGroup As Boolean
    Dim intLoopCnt1 As Integer

    flgFirstAllData = False
    intNowLine = 0
    Do Until UBound(paryCsvData) < intNowLine

        strCSVRecord = paryCsvData(intNowLine)

        If Not flgFirstAllData Then
            '配列領域定義
            ReDim aryAllLineInfo(0)
            'ワーク変数
            ReDim aryEachLineInfo(2)
            aryEachLineInfo(0) = strCSVRecord(enmModColNo.HenkouDate)
            aryEachLineInfo(1) = strCSVRecord(enmModColNo.janCD)
            If strCSVRecord(enmModColNo.Hacchuutani) <> "" Or strCSVRecord(enmModColNo.Genka) <> "" Then
                aryEachLineInfo(2) = "1"
            Else
                aryEachLineInfo(2) = "0"
            End If
            aryAllLineInfo(0) = aryEachLineInfo
            flgFirstAllData = True
        Else
            flgSameGroup = False

            For intLoopCnt1 = 0 To UBound(aryAllLineInfo, 1)
                If aryAllLineInfo(intLoopCnt1)(0) = strCSVRecord(enmModColNo.HenkouDate) _
                And aryAllLineInfo(intLoopCnt1)(1) = strCSVRecord(enmModColNo.janCD) Then
                                        
                    flgSameGroup = True
                    If strCSVRecord(enmModColNo.Hacchuutani) <> "" Or strCSVRecord(enmModColNo.Genka) <> "" Then
                        aryAllLineInfo(intLoopCnt1)(2) = "1"
                    End If

                    Exit For
                    
                End If
            Next
            If Not flgSameGroup Then
            
            ReDim Preserve aryAllLineInfo(UBound(aryAllLineInfo) + 1)
                'ワーク変数
                ReDim aryEachLineInfo(2)
                aryEachLineInfo(0) = strCSVRecord(enmModColNo.HenkouDate)
                aryEachLineInfo(1) = strCSVRecord(enmModColNo.janCD)
                If strCSVRecord(enmModColNo.Hacchuutani) <> "" Or strCSVRecord(enmModColNo.Genka) <> "" Then
                    aryEachLineInfo(2) = "1"
                Else
                    aryEachLineInfo(2) = "0"
                End If
        
                aryAllLineInfo(UBound(aryAllLineInfo)) = aryEachLineInfo
                
            End If
        End If

        intNowLine = intNowLine + 1
    Loop
    
Call SetIkkatuLog("1-1-9", "", "", "", "OK", "", "UPD")
Call SetIkkatuLog("1-1-10", "(発注単位、原価が一致チェック)", "SQL", "", "", pstrTantousyaId, "INS")

    intNowLine = 0
        
    Dim strHattyuuTaniHonbu As String
    Dim strGenkaHonbu As String
    Dim strHattyuuTani As String
    Dim strGenka As String

    For intLoopCnt = 0 To UBound(aryAllLineInfo, 1)
    '「原価」がともに入力がない場合は以下のチェックは行わない。
    If aryAllLineInfo(intLoopCnt)(2) = "1" Then
    
        strHattyuuTaniHonbu = ""
        strGenkaHonbu = ""
    
        '一括登録票以外の店舗で店別情報(v_mise_betu)にセンター区分が2：PDCの店舗の発注単位、原価
        'と
        '一括登録票の発注単位、原価が一致しない場合
        '発注単位、原価が一致チェック用切断レコードセット生成
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(発注単位、原価が一致チェック)"
            
        '発注単位、原価が一致チェック用切断レコードセット生成(本部)
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(発注単位、原価が一致チェック用切断レコードセット生成(本部))"
        strSQL = ""
        strSQL = strSQL & "SELECT                                                   " & vbCrLf
        strSQL = strSQL & "    vsk.hyoujyun_hattyuu_tani AS hattyuu_tani            " & vbCrLf
        strSQL = strSQL & "    ,vsk.hyoujyun_genka AS genka                         " & vbCrLf
        strSQL = strSQL & "FROM                                                     " & vbCrLf
        strSQL = strSQL & "    v_syanai_kihon vsk                                   " & vbCrLf
        strSQL = strSQL & "    INNER JOIN v_key_1 AS vk1                            " & vbCrLf
        strSQL = strSQL & "    ON vk1.key_cd=vsk.key_cd                             " & vbCrLf
        strSQL = strSQL & "    AND vk1.jan_cd='" & aryAllLineInfo(intLoopCnt)(1) & "'" & vbCrLf
        
            pChkHonbuGenkaHattyuuTaniRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
        With pChkHonbuGenkaHattyuuTaniRS
            If Not .BOF Then
                .MoveFirst
                '本部の発注単位、原価の取得
                strHattyuuTaniHonbu = .Fields("hattyuu_tani")
                strGenkaHonbu = .Fields("genka")
            End If
        End With
    
        intNowLine = 0
        Do Until UBound(paryCsvData) < intNowLine

            strCSVRecord = paryCsvData(intNowLine)
            
            If strCSVRecord(enmModColNo.HenkouDate) = aryAllLineInfo(intLoopCnt)(0) _
                And strCSVRecord(enmModColNo.janCD) = aryAllLineInfo(intLoopCnt)(1) Then
            
                If strCSVRecord(enmModColNo.TenCd) = "0000" Then
                
                    '本部の発注単位、原価のセット
                    If strCSVRecord(enmModColNo.Hacchuutani) <> "" Then
                        strHattyuuTaniHonbu = strCSVRecord(enmModColNo.Hacchuutani)
                    End If
                    If strCSVRecord(enmModColNo.Genka) <> "" Then
                        strGenkaHonbu = strCSVRecord(enmModColNo.Genka)
                    End If
                                           
                End If
    
            End If
            
            intNowLine = intNowLine + 1
        Loop
    
    
        '発注単位、原価が一致チェック用切断レコードセット生成
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(発注単位、原価が一致チェック用切断レコードセット生成)"
        strSQL = "SELECT VMB.ten_cd,VMB.hattyuu_tani,VMB.genka,VMB.center_kbn,VJ.jigyou_cd " & vbCrLf
        strSQL = strSQL & " FROM v_mise_betu as VMB WITH (READCOMMITTED)" & vbCrLf
        strSQL = strSQL & "      INNER JOIN v_key_1 as VK1 WITH(READCOMMITTED)" & vbCrLf
        strSQL = strSQL & "      ON VK1.jan_cd = '" & aryAllLineInfo(intLoopCnt)(1) & "'" & vbCrLf
        strSQL = strSQL & "      AND VK1.key_cd = VMB.key_cd" & vbCrLf
        strSQL = strSQL & "      INNER JOIN v_jigyousyo VJ WITH(READCOMMITTED)" & vbCrLf
        strSQL = strSQL & "      ON VJ.jigyousyo_cd=VMB.ten_cd " & vbCrLf
        strSQL = strSQL & "      INNER JOIN m_wf_meisyou MWM WITH (READCOMMITTED) " & vbCrLf
        strSQL = strSQL & "      ON MWM.key_cd = VJ.jigyou_cd " & vbCrLf
        strSQL = strSQL & "      AND MWM.key_kbn = 'W02' " & vbCrLf
        strSQL = strSQL & "      AND SUBSTRING(MWM.settei_naiyou,4, 1) = '1' " & vbCrLf
        strSQL = strSQL & " WHERE" & vbCrLf
        strSQL = strSQL & "      VMB.center_kbn = '2'" & vbCrLf
        strSQL = strSQL & "      AND VJ.hanbetu_kbn = '5'" & vbCrLf
        
            pChkSameGenkaHattyuuTaniRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
        pChkSameGenkaHattyuuTaniRS.Fields("ten_cd").Properties("Optimize") = True 'オプティマイザ
        
        With pChkSameGenkaHattyuuTaniRS
            If Not .BOF Then
                .MoveFirst
                '店別、事業所の場合
                Do Until .EOF
                    intNowLine = 0
                    Do Until UBound(paryCsvData) < intNowLine
    
                        strCSVRecord = paryCsvData(intNowLine)
                        
                        If strCSVRecord(enmModColNo.HenkouDate) = aryAllLineInfo(intLoopCnt)(0) _
                            And strCSVRecord(enmModColNo.janCD) = aryAllLineInfo(intLoopCnt)(1) Then
                        
                            If Trim(.Fields("ten_cd")) = strCSVRecord(enmModColNo.TenCd) Or Trim(.Fields("jigyou_cd")) = strCSVRecord(enmModColNo.TenCd) Then
                                
                                If strCSVRecord(enmModColNo.CenterKbn) = "" Then
                                    'センターが変更しない場合
                                    If strCSVRecord(enmModColNo.Hacchuutani) <> "" Then
                                        .Fields("hattyuu_tani") = strCSVRecord(enmModColNo.Hacchuutani)
                                    End If
                                    If strCSVRecord(enmModColNo.Genka) <> "" Then
                                        .Fields("genka") = strCSVRecord(enmModColNo.Genka)
                                    End If
                                Else
                                    'センターを変更する場合
                                    .Fields("center_kbn") = strCSVRecord(enmModColNo.CenterKbn)
                                    If strCSVRecord(enmModColNo.Hacchuutani) <> "" Then
                                        .Fields("hattyuu_tani") = strCSVRecord(enmModColNo.Hacchuutani)
                                    End If
                                    If strCSVRecord(enmModColNo.Genka) <> "" Then
                                        .Fields("genka") = strCSVRecord(enmModColNo.Genka)
                                    End If

                                End If
                                
                            End If
                
                        End If
                        
                        intNowLine = intNowLine + 1
                    Loop
                    
                    .MoveNext
                Loop
                
                '9999の場合
                intNowLine = 0
                Do Until UBound(paryCsvData) < intNowLine
                    
                    strCSVRecord = paryCsvData(intNowLine)
                    
                    If strCSVRecord(enmModColNo.HenkouDate) = aryAllLineInfo(intLoopCnt)(0) _
                        And strCSVRecord(enmModColNo.janCD) = aryAllLineInfo(intLoopCnt)(1) Then
                    
                        If strCSVRecord(enmModColNo.TenCd) = "9999" Then
                            
                            If strCSVRecord(enmModColNo.CenterKbn) = "" Then
                                .MoveFirst
                                Do Until .EOF
                                    'センターが変更しない場合
                                    If strCSVRecord(enmModColNo.Hacchuutani) <> "" Then
                                        .Fields("hattyuu_tani") = strCSVRecord(enmModColNo.Hacchuutani)
                                    End If
                                    If strCSVRecord(enmModColNo.Genka) <> "" Then
                                        .Fields("genka") = strCSVRecord(enmModColNo.Genka)
                                    End If
                                    .MoveNext
                                Loop

                            Else
                                .MoveFirst
                                Do Until .EOF
                                    'センターを変更する場合
                                    .Fields("center_kbn") = strCSVRecord(enmModColNo.CenterKbn)
                                    If strCSVRecord(enmModColNo.Hacchuutani) <> "" Then
                                        .Fields("hattyuu_tani") = strCSVRecord(enmModColNo.Hacchuutani)
                                    End If
                                    If strCSVRecord(enmModColNo.Genka) <> "" Then
                                        .Fields("genka") = strCSVRecord(enmModColNo.Genka)
                                    End If
                                    .MoveNext
                                Loop

                            End If
                        End If
                    End If
                        
                    intNowLine = intNowLine + 1
                Loop
                    
                strHattyuuTani = ""
                strGenka = ""
                .MoveFirst
                Do Until .EOF
                    If .Fields("center_kbn") = "2" Then
                        strHattyuuTani = Trim(.Fields("hattyuu_tani"))
                        strGenka = Trim(.Fields("genka"))
                        .MoveLast
                    End If
                    .MoveNext
                Loop
                    
                .MoveFirst
                Do Until .EOF
                    If .Fields("center_kbn") = "2" Then
                        If Trim(.Fields("hattyuu_tani")) <> strHattyuuTani _
                        Or Trim(.Fields("genka")) <> strGenka _
                        Or strHattyuuTani <> strHattyuuTaniHonbu _
                        Or strGenka <> strGenkaHonbu Then
                            'エラーの場合
                                    FncSetErrorPDC(intNowLine, enmModColNo.TenCd, PARAM_GENKA_HATTYUUTANI, aryAllLineInfo(intLoopCnt)(0) & "," & aryAllLineInfo(intLoopCnt)(1))
                            flgChkError = True
                            .MoveLast
                        End If
                    End If
                    .MoveNext
                Loop
                
            End If
        End With
    End If
    Next
Call SetIkkatuLog("1-1-10", "", "", "", "OK", "", "UPD")

    '********** 2012/08/03 楊双（大連）入力チェックの追加 **********↑
        pTChkBaikaHenkouRS.Close
            pTChkBaikaHenkouRS = Nothing
       
        pTChkTorihikisakiRS.Close
            pTChkTorihikisakiRS = Nothing
       
        pTChkCenterRS.Close
            pTChkCenterRS = Nothing
       
        pTChkMstBaikaRS.Close
            pTChkMstBaikaRS = Nothing
       
        pTChkMstSiireRS.Close
            pTChkMstSiireRS = Nothing
       
        pTChkMstCenterRS.Close
            pTChkMstCenterRS = Nothing
        
        '********** 2012/08/03 楊双（大連）入力チェックの追加 **********↓
        If Not pChkSameGenkaHattyuuTaniRS Is Nothing Then
            pChkSameGenkaHattyuuTaniRS.Close
                pChkSameGenkaHattyuuTaniRS = Nothing
        End If
        
        If Not pChkHonbuGenkaHattyuuTaniRS Is Nothing Then
            pChkHonbuGenkaHattyuuTaniRS.Close
                pChkHonbuGenkaHattyuuTaniRS = Nothing
        End If
        '********** 2012/08/03 楊双（大連）入力チェックの追加 **********↑
        
        'DBクローズ
        objCDBAccess = Nothing
        'GetObjectContext.SetComplete
   
    'データに１つでも問題がある場合は、ここで終了！！！
    If flgChkError = True Then
        'エラーがあればエラーメッセージを取得する。（終了しない）
        strErrMsg = FncGetErrorMsg_Modosi()
        '********** 2012/08/03 楊双（大連）入力チェックの追加 **********↓
        If strErrMsg = "" Then
            strErrMsg = FncGetErrorMsg_PDC()
        Else
            strErrMsg = strErrMsg & FncGetErrorMsg_PDC()
        End If
        '********** 2012/08/03 楊双（大連）入力チェックの追加 **********↑
        Exit Function
    End If
     
     'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(データ登録)"


    '企画種類
    If pstrKikakuSyurui = "012" Or pstrKikakuSyurui = "015" Then
    
                Call SetIkkatuLog("1-2", "商品選定WF(戻し価格)", "InsertModosiKakaku", "", "", pstrTantousyaId, "INS")
        InsertModosiKakaku (paryCsvData)
        Call SetIkkatuLog("1-2", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-3", "商品選定WF(戻し価格)", "subIns_SetAllTablesMosi", "", "", pstrTantousyaId, "INS")
        subIns_SetAllTablesMosi (paryCsvData)
        Call SetIkkatuLog("1-3", "", "", "", "OK", "", "UPD")
        
    Else
        Call SetIkkatuLog("1-2", "商品選定WF(戻し価格)", "subIns_SetAllTablesMosi", "", "", pstrTantousyaId, "INS")
        subIns_SetAllTables (paryCsvData)
        Call SetIkkatuLog("1-2", "", "", "", "OK", "", "UPD")
        Call SetIkkatuLog("1-3", "なし", "", "", "", pstrTantousyaId, "INS")
        Call SetIkkatuLog("1-3", "", "", "", "OK", "", "UPD")
    End If

    
    '一括登録
    If pstrKikakuSyurui = "015" Then
        Call SetIkkatuLog("1-4", "一括登録", "UpdGenka_HattuTani", "", "", pstrTantousyaId, "INS")
        UpdGenka_HattuTani (paryCsvData)
        Call SetIkkatuLog("1-4", "", "", "", "OK", "", "UPD")
    End If

    '排他ロック解除
    pChkSameJANCodeRS.Close
        pChkSameJANCodeRS = Nothing
    objCDBAccess = Nothing
    'GetObjectContext.SetComplete
    Call SetIkkatuLog("1-1", "", "", "", "OK", "", "UPD")
    
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(データ登録終了)"

    FncModosiKakaku = True

Exit Function

DefectErrorHandler:
    
    If Not pTChkJigyouRS Is Nothing Then
            pTChkJigyouRS = Nothing
    End If
    If Not pChkSameJANCodeRS Is Nothing Then
            pChkSameJANCodeRS = Nothing
    End If
    If Not pChkvk1JANCodeRS Is Nothing Then
            pChkvk1JANCodeRS = Nothing
    End If
    If Not pTChkBaikaHenkouRS Is Nothing Then
            pTChkBaikaHenkouRS = Nothing
    End If
    If Not pTChkTorihikisakiRS Is Nothing Then
            pTChkTorihikisakiRS = Nothing
    End If
    If Not pTChkCenterRS Is Nothing Then
            pTChkCenterRS = Nothing
    End If
    If Not pTChkMstBaikaRS Is Nothing Then
            pTChkMstBaikaRS = Nothing
    End If
    If Not pTChkMstSiireRS Is Nothing Then
            pTChkMstSiireRS = Nothing
    End If
    If Not pTChkMstCenterRS Is Nothing Then
            pTChkMstCenterRS = Nothing
    End If
    '********** 2012/08/03 楊双（大連）入力チェックの追加 **********↓
    If Not pChkSameGenkaHattyuuTaniRS Is Nothing Then
            pChkSameGenkaHattyuuTaniRS = Nothing
    End If
    
    If Not pChkHonbuGenkaHattyuuTaniRS Is Nothing Then
            pChkHonbuGenkaHattyuuTaniRS = Nothing
    End If
    '********** 2012/08/03 楊双（大連）入力チェックの追加 **********↑

    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
    
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    Call SetIkkatuLog("1-1", "", "", Left(Err.Description, 400), "NG", "", "UPD")
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description

End Function
'******************************************************************'
'*【プロシージャ名】
'*!    InsertModosiKakaku
'*【処理概要】
'*?    INSERT文作成（商品選定WF(戻し価格)）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*【戻り値】
'*    < Name Type Comment >
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2009/06/08; Supattra T(Tostem Thai); 新規作成;
'*M    2010/03/18; 営業システム 和田; 建デポ対応;
'*< Comment End >**************************************************'
Private Sub InsertModosiKakaku(ByVal arryData As Object)
        Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim strMethodCallInfo As String
    Dim lstrKikakuSyurui As String
    Dim iRecord As Long
    Dim strTenCd As String
    Dim strTenCdInsert As String
    Dim janCD As String
    Dim strGaitouMiseCD As String
    Dim btnCheck As Boolean
    Dim i As Integer
    Dim strTenCdArry As Object
    Dim strChkData As Boolean
    
    lstrKikakuSyurui = pstrKikakuSyurui
    
    Dim lstrKikakuNenrenban As String
    lstrKikakuNenrenban = pstrKikakuNenrenban
    
    Dim lstrKikakuEdaban As String
    lstrKikakuEdaban = pstrKikakuEdaban
    
    Dim lstrTantousyaId As String
    lstrTantousyaId = pstrTantousyaId
    
    Dim lstrNow As Object
    lstrNow = Format(Now(), "yyyy/mm/dd hh:nn:ss")
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.InsertModosiKakaku()"
    
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    On Error GoTo DefectErrorHandler
    
    '事業所マスタ(v_jigyousyo)より、店舗を取得する。
    Dim strSQL  As String
    
    '2009.07.02 Add by Supattra T.(Tostem Thai)Start.
    
    Dim strTenkai_kaisi_ym_date As String
            
    '企画情報DB(企画詳細情報).展開開始年月日の年月を取得する。
    strSQL = ""
    strSQL = strSQL & " SELECT  " & vbCrLf
    strSQL = strSQL & " LEFT(CONVERT(CHAR,tenkai_kaisi_ym_date,112),6) AS tenkai_kaisi_ym_date " & vbCrLf
    strSQL = strSQL & " ,DATEADD(MONTH,-1,tenkai_kaisi_ym_date) AS tenkai_kaisi_ym_date_1 " & vbCrLf
    strSQL = strSQL & " FROM t_kikaku_jyouhou WITH(READCOMMITTED) " & vbCrLf
    strSQL = strSQL & " WHERE kikaku_syurui_cd ='" & lstrKikakuSyurui & "'" & vbCrLf
    strSQL = strSQL & " AND nen_ren_no = '" & lstrKikakuNenrenban & "'" & vbCrLf
    strSQL = strSQL & " AND eda_no ='" & lstrKikakuEdaban & "'" & vbCrLf

    Dim tenkaiDateRS As ADODB.Recordset
        tenkaiDateRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    
    'tenkaiDateRS.Fields("tenkai_kaisi_ym_date").Properties("Optimize") = True 'オプティマイザ
    
    strTenkai_kaisi_ym_date = tenkaiDateRS.Fields("tenkai_kaisi_ym_date").Value
    
    If Not tenkaiDateRS Is Nothing Then
            tenkaiDateRS = Nothing
    End If
    
    
    '事業所マスタ(v_jigyousyo)より、店舗を取得する。
    strSQL = ""
    strSQL = strSQL & "  SELECT " & vbCrLf
    strSQL = strSQL & "  A.jigyousyo_cd " & vbCrLf
    strSQL = strSQL & "  ,B.ten_cd " & vbCrLf
    strSQL = strSQL & "  ,A.key_cd " & vbCrLf
    strSQL = strSQL & "  ,CASE WHEN B.ten_cd IS NULL  THEN '1' ELSE '0' END Flg " & vbCrLf
    strSQL = strSQL & "  FROM ( " & vbCrLf
    strSQL = strSQL & "  SELECT V.jigyousyo_cd  " & vbCrLf
    strSQL = strSQL & "         ,A.key_cd  " & vbCrLf
    strSQL = strSQL & "   FROM v_jigyousyo V WITH(READCOMMITTED)  " & vbCrLf
    strSQL = strSQL & "  INNER JOIN " & vbCrLf
    strSQL = strSQL & "  (SELECT " & vbCrLf
    strSQL = strSQL & "    key_cd " & vbCrLf
    strSQL = strSQL & "   FROM " & vbCrLf
    strSQL = strSQL & "    m_wf_meisyou  WITH(READCOMMITTED) " & vbCrLf
    strSQL = strSQL & "   WHERE " & vbCrLf
    strSQL = strSQL & "    key_kbn = 'W02'" & vbCrLf
    strSQL = strSQL & "   AND " & vbCrLf
    strSQL = strSQL & "    substring(rtrim(settei_naiyou),4,1) = '1' " & vbCrLf
    strSQL = strSQL & "  ) as A " & vbCrLf
    strSQL = strSQL & "  ON V.jigyou_cd = A.key_cd " & vbCrLf
    strSQL = strSQL & "  WHERE V.hanbetu_kbn = '5' " & vbCrLf
    strSQL = strSQL & "  AND (V.heiten_date is null or V.heiten_date = '') or(V.heiten_date  >= CONVERT(CHAR,GETDATE(),112))  " & vbCrLf
    strSQL = strSQL & ") A  " & vbCrLf
    
    strSQL = strSQL & " Left Join " & vbCrLf
    strSQL = strSQL & " (" & vbCrLf
    strSQL = strSQL & " SELECT  " & vbCrLf
    strSQL = strSQL & "  ten_cd " & vbCrLf
    strSQL = strSQL & " FROM  " & vbCrLf
    strSQL = strSQL & "  t_syousen_jyogai_tenpo WITH(READCOMMITTED) " & vbCrLf
    strSQL = strSQL & " WHERE  " & vbCrLf
    strSQL = strSQL & "     kikaku_syurui_cd ='" & lstrKikakuSyurui & "'" & vbCrLf
    strSQL = strSQL & " AND nen_ren_no = '" & lstrKikakuNenrenban & "'" & vbCrLf
    strSQL = strSQL & " AND eda_no ='" & lstrKikakuEdaban & "'" & vbCrLf
    strSQL = strSQL & " AND baika_jyogai_kbn ='1'" & vbCrLf
    strSQL = strSQL & " AND genka_jyogai_kbn  ='1'" & vbCrLf
    strSQL = strSQL & " AND hattyuu_tani_jyogai_kbn  = '1'" & vbCrLf
    
    strSQL = strSQL & " ) B" & vbCrLf
    strSQL = strSQL & " ON " & vbCrLf
    strSQL = strSQL & " A.jigyousyo_cd = B.ten_cd " & vbCrLf
    strSQL = strSQL & " Order By " & vbCrLf
    strSQL = strSQL & " A.key_cd " & vbCrLf
    
    Dim baikaRS As ADODB.Recordset
        baikaRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    
        'baikaRS.Fields("jigyousyo_cd").Properties("Optimize") = True 'オプティマイザ
    
    Dim strTenCdInsert01 As String
    Dim strTenCdInsert03 As String
    Dim strTenCdInsert06 As String
    Dim strTenCdInsert07 As String
    Dim strTenCdInsert9999 As String
    Dim blnChck01 As Boolean
    Dim blnChck03 As Boolean
    Dim blnChck06 As Boolean
    Dim blnChck07 As Boolean
    blnChck01 = False
    blnChck03 = False
    blnChck06 = False
    blnChck07 = False
    
    'Mai2
    If baikaRS.EOF = False Then
        strChkData = True
        Do Until baikaRS.EOF
            If baikaRS.Fields("key_cd").Value = "01" Then
                blnChck01 = True
                If baikaRS.Fields("Flg").Value = "1" Then
                    strTenCdInsert01 = strTenCdInsert01 & "'" & baikaRS.Fields("jigyousyo_cd").Value & "',"
                End If
            ElseIf baikaRS.Fields("key_cd").Value = "03" Then
                 blnChck03 = True
                If baikaRS.Fields("Flg").Value = "1" Then
                    strTenCdInsert03 = strTenCdInsert03 & "'" & baikaRS.Fields("jigyousyo_cd").Value & "',"
                End If
            ElseIf baikaRS.Fields("key_cd").Value = "06" Then
                blnChck06 = True
                If baikaRS.Fields("Flg").Value = "1" Then
                    strTenCdInsert06 = strTenCdInsert06 & "'" & baikaRS.Fields("jigyousyo_cd").Value & "',"
                End If
            ElseIf baikaRS.Fields("key_cd").Value = "07" Then
                blnChck07 = True
                If baikaRS.Fields("Flg").Value = "1" Then
                    strTenCdInsert07 = strTenCdInsert07 & "'" & baikaRS.Fields("jigyousyo_cd").Value & "',"
                End If
            End If
            If baikaRS.Fields("Flg").Value = "1" Then
                 strTenCdInsert9999 = strTenCdInsert9999 & "'" & baikaRS.Fields("jigyousyo_cd").Value & "',"
            End If
            baikaRS.MoveNext
        Loop
        
    Else
        strChkData = False
    End If
    'Mai2

    '変更除外店舗を登録対象外にする。
    '商品選定DB(変更除外店舗)を企画Noで検索する。
    strSQL = ""
    strSQL = strSQL & " SELECT  " & vbCrLf
    strSQL = strSQL & "  ten_cd " & vbCrLf
    strSQL = strSQL & "  ,baika_jyogai_kbn " & vbCrLf
    strSQL = strSQL & "  ,genka_jyogai_kbn " & vbCrLf
    strSQL = strSQL & "  ,hattyuu_tani_jyogai_kbn " & vbCrLf
    strSQL = strSQL & " FROM  " & vbCrLf
    strSQL = strSQL & "  t_syousen_jyogai_tenpo WITH(READCOMMITTED) " & vbCrLf
    strSQL = strSQL & " WHERE  " & vbCrLf
    strSQL = strSQL & "  kikaku_syurui_cd ='" & lstrKikakuSyurui & "'" & vbCrLf
    strSQL = strSQL & " AND nen_ren_no = '" & lstrKikakuNenrenban & "'" & vbCrLf
    strSQL = strSQL & " AND eda_no ='" & lstrKikakuEdaban & "'" & vbCrLf
    
    Dim tenCdRS As ADODB.Recordset
        tenCdRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    'tenCdRS.Fields("baika_jyogai_kbn").Properties("Optimize") = True 'オプティマイザ
        'tenCdRS.Fields("ten_cd").Properties("Optimize") = True 'オプティマイザ
    
    '2009.07.02 Add by Supattra T.(Tostem Thai)End.
    
        Dim nengetuRS As New ADODB.Recordset

    For iRecord = LBound(arryData) To UBound(arryData)

        janCD = arryData(iRecord)(enmTenColNo.janCD)
        strGaitouMiseCD = arryData(iRecord)(enmTenColNo.TenCd)
        
        '店舗コード（strGaitouMiseCD）が、"9999：全店"、"01：VH"、"03：VCS"、"06：SVH"の場合、店舗コードに展開する。
        '"99：全事業"、"07：建デポ"を追加 2010/03/18
        If Trim(strGaitouMiseCD) = "9999" Or Trim(strGaitouMiseCD) = "01" Or Trim(strGaitouMiseCD) = "03" _
            Or Trim(strGaitouMiseCD) = "06" Or Trim(strGaitouMiseCD) = "99" Or Trim(strGaitouMiseCD) = "07" Then
            
            If strChkData = True Then
                If Trim(strGaitouMiseCD) = "01" Then
                    If blnChck01 = False Then
                        strTenCdInsert = "'" & arryData(iRecord)(enmTenColNo.TenCd) & "',"
                    Else
                         strTenCdInsert = strTenCdInsert01
                    End If
                ElseIf Trim(strGaitouMiseCD) = "03" Then
                    If blnChck03 = False Then
                        strTenCdInsert = "'" & arryData(iRecord)(enmTenColNo.TenCd) & "',"
                    Else
                         strTenCdInsert = strTenCdInsert03
                    End If
                ElseIf Trim(strGaitouMiseCD) = "06" Then
                    If blnChck06 = False Then
                        strTenCdInsert = "'" & arryData(iRecord)(enmTenColNo.TenCd) & "',"
                    Else
                         strTenCdInsert = strTenCdInsert06
                    End If
                ElseIf Trim(strGaitouMiseCD) = "07" Then
                    If blnChck07 = False Then
                        strTenCdInsert = "'" & arryData(iRecord)(enmTenColNo.TenCd) & "',"
                    Else
                         strTenCdInsert = strTenCdInsert07
                    End If
                Else
                    strTenCdInsert = strTenCdInsert9999
                End If
            Else
                strTenCdInsert = "'" & arryData(iRecord)(enmTenColNo.TenCd) & "',"
            End If
       
            If strTenCdInsert <> "" Then
                strTenCdInsert = Left(strTenCdInsert, Len(strTenCdInsert) - 1)
                btnCheck = True
            Else
                 btnCheck = False
            End If
            
        Else
            strTenCd = "'" & arryData(iRecord)(enmTenColNo.TenCd) & "'"
                
            With tenCdRS
                If Not .BOF Then
                    .MoveFirst
                    .Find ("ten_cd = " & strTenCd)
                End If
                
                If tenCdRS.EOF = False Then
                    If tenCdRS.Fields("baika_jyogai_kbn").Value = "1" And tenCdRS.Fields("genka_jyogai_kbn").Value = "1" And tenCdRS.Fields("hattyuu_tani_jyogai_kbn").Value = "1" Then
                         btnCheck = False
                    Else
                         strTenCdInsert = strTenCd
                         btnCheck = True
                    End If
                Else
                    strTenCdInsert = strTenCd
                    btnCheck = True
                End If
                
            End With
        End If
       
        '2009.07.02 Add by Supattra T.(Tostem Thai)Start.
        ' 取得した店舗コード数分で繰り返し。
        strSQL = ""
        strSQL = strSQL & " SELECT " & vbCrLf
        strSQL = strSQL & " ten_cd " & vbCrLf
        strSQL = strSQL & ",nengetu " & vbCrLf
        strSQL = strSQL & ",jan_cd " & vbCrLf
        strSQL = strSQL & ",genka " & vbCrLf
        strSQL = strSQL & ",baika " & vbCrLf
        strSQL = strSQL & ",hattyuu_tani " & vbCrLf
        strSQL = strSQL & ",pop_size " & vbCrLf
        strSQL = strSQL & ",touroku_user_id " & vbCrLf
        strSQL = strSQL & ",kousin_user_id " & vbCrLf
        strSQL = strSQL & ",touroku_ymd_hms " & vbCrLf
        strSQL = strSQL & ",kousin_ymd_hms " & vbCrLf
        strSQL = strSQL & "FROM t_modosi_kakaku  " & vbCrLf
        strSQL = strSQL & " WHERE nengetu ='" & strTenkai_kaisi_ym_date & "'" & vbCrLf
        strSQL = strSQL & " AND jan_cd = '" & janCD & "'" & vbCrLf
        


            nengetuRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
            
            'nengetuRS.Fields("ten_cd").Properties("Optimize") = True 'オプティマイザ
        '2009.07.02 Add by Supattra T.(Tostem Thai)End.
        
        If btnCheck = True Then
            
            strTenCdArry = Split(strTenCdInsert, ",")
                
            For i = 0 To UBound(strTenCdArry)
                strTenCdArry(i) = Left(strTenCdArry(i), Len(strTenCdArry(i)) - 1)
                strTenCdArry(i) = Right(strTenCdArry(i), Len(strTenCdArry(i)) - 1)
                
                With nengetuRS
                    If Not .BOF Then
                        .MoveFirst
                        .Find ("ten_cd = '" & strTenCdArry(i) & "'")
                    End If
                End With

                    If nengetuRS.EOF Then
                        With nengetuRS
                            'Insert
                            .AddNew()
                            .Fields.Item("nengetu") = strTenkai_kaisi_ym_date

                            .Fields("jan_cd") = janCD
                            .Fields("ten_cd") = strTenCdArry(i)
                            .Fields("genka") = arryData(iRecord)(enmModColNo.MoGenka)
                            If Trim(arryData(iRecord)(enmModColNo.MoBaika)) <> "" Then
                                .Fields("baika") = arryData(iRecord)(enmModColNo.MoBaika)
                            End If
                            .Fields("hattyuu_tani") = arryData(iRecord)(enmModColNo.HattyuuTani)
                            If Trim(arryData(iRecord)(enmModColNo.PopSize)) <> "" Then
                                .Fields("pop_size") = arryData(iRecord)(enmModColNo.PopSize)
                            End If
                            .Fields("touroku_user_id") = lstrTantousyaId
                            .Fields("kousin_user_id") = lstrTantousyaId
                            .Fields("touroku_ymd_hms") = lstrNow
                            .Fields("kousin_ymd_hms") = lstrNow
                            .Update()
                        End With
                    Else
                        With nengetuRS
                            'Update
                            .Fields("genka") = arryData(iRecord)(enmModColNo.MoGenka)
                            If Trim(arryData(iRecord)(enmModColNo.MoBaika)) <> "" Then
                                .Fields("baika") = arryData(iRecord)(enmModColNo.MoBaika)
                            End If
                            .Fields("hattyuu_tani") = arryData(iRecord)(enmModColNo.HattyuuTani)
                            If Trim(arryData(iRecord)(enmModColNo.PopSize)) <> "" Then
                                .Fields("pop_size") = arryData(iRecord)(enmModColNo.PopSize)
                            End If
                            .Fields("kousin_user_id") = lstrTantousyaId
                            .Fields("kousin_ymd_hms") = lstrNow
                            .Update()
                        End With
                    End If
            Next

            Call objCDBAccess.UpdateDCRecordset("VSWF", nengetuRS)
        End If
        
    Next iRecord
    

    '「変更日」を、まる得企画担当者.変更日(新規)又は、変更日(継続)に変更する。
    
    strSQL = ""
    strSQL = strSQL & " SELECT  " & vbCrLf
    strSQL = strSQL & "     jan_cd " & vbCrLf
    strSQL = strSQL & "     ,ten_cd " & vbCrLf
    strSQL = strSQL & " FROM " & vbCrLf
    strSQL = strSQL & "     (SELECT " & vbCrLf
    strSQL = strSQL & "         b_rireki.jan_cd " & vbCrLf
    strSQL = strSQL & "         ,b_rireki.ten_cd " & vbCrLf
    strSQL = strSQL & "     FROM " & vbCrLf
    strSQL = strSQL & "         t_syousen_baika_henkou_rireki b_rireki " & vbCrLf
    strSQL = strSQL & "         INNER JOIN " & vbCrLf
    strSQL = strSQL & "             (SELECT " & vbCrLf
    strSQL = strSQL & "                 kikaku_syurui_cd " & vbCrLf
    strSQL = strSQL & "                 ,nen_ren_no " & vbCrLf
    strSQL = strSQL & "                 ,eda_no " & vbCrLf
    strSQL = strSQL & "             FROM " & vbCrLf
    strSQL = strSQL & "                 t_kikaku_jyouhou " & vbCrLf
    strSQL = strSQL & "             WHERE " & vbCrLf
    strSQL = strSQL & "                 kikaku_syurui_cd = '" & lstrKikakuSyurui & "'" & vbCrLf
    strSQL = strSQL & "                 AND CONVERT(CHAR(6),tenkai_kaisi_ym_date,112) = " & vbCrLf
    strSQL = strSQL & "                     (SELECT " & vbCrLf
    strSQL = strSQL & "                         CONVERT(CHAR(6),DATEADD(Month,-1,tenkai_kaisi_ym_date),112) " & vbCrLf
    strSQL = strSQL & "                     FROM " & vbCrLf
    strSQL = strSQL & "                         t_kikaku_jyouhou " & vbCrLf
    strSQL = strSQL & "                     WHERE " & vbCrLf
    strSQL = strSQL & "                         kikaku_syurui_cd = '" & lstrKikakuSyurui & "'" & vbCrLf
    strSQL = strSQL & "                         AND nen_ren_no = '" & lstrKikakuNenrenban & "'" & vbCrLf
    strSQL = strSQL & "                         AND eda_no = '" & lstrKikakuEdaban & "'" & vbCrLf
    strSQL = strSQL & "                     ) " & vbCrLf
    strSQL = strSQL & "             ) kikaku " & vbCrLf
    strSQL = strSQL & "         ON " & vbCrLf
    strSQL = strSQL & "             b_rireki.kikaku_syurui_cd = kikaku.kikaku_syurui_cd " & vbCrLf
    strSQL = strSQL & "             AND b_rireki.nen_ren_no = kikaku.nen_ren_no " & vbCrLf
    strSQL = strSQL & "             AND b_rireki.eda_no = kikaku.eda_no " & vbCrLf
    strSQL = strSQL & "     WHERE " & vbCrLf
    strSQL = strSQL & "         b_rireki.baika IS NOT NULL " & vbCrLf

    strSQL = strSQL & "     UNION " & vbCrLf

    strSQL = strSQL & "     SELECT " & vbCrLf
    strSQL = strSQL & "         s_rireki.jan_cd " & vbCrLf
    strSQL = strSQL & "         ,s_rireki.ten_cd " & vbCrLf
    strSQL = strSQL & "     FROM " & vbCrLf
    strSQL = strSQL & "         t_syousen_siire_henkou_rireki s_rireki " & vbCrLf
    strSQL = strSQL & "         INNER JOIN " & vbCrLf
    strSQL = strSQL & "             (SELECT " & vbCrLf
    strSQL = strSQL & "                 kikaku_syurui_cd " & vbCrLf
    strSQL = strSQL & "                 ,nen_ren_no " & vbCrLf
    strSQL = strSQL & "                 ,eda_no " & vbCrLf
    strSQL = strSQL & "             FROM " & vbCrLf
    strSQL = strSQL & "                 t_kikaku_jyouhou " & vbCrLf
    strSQL = strSQL & "             WHERE" & vbCrLf
    strSQL = strSQL & "                 kikaku_syurui_cd ='" & lstrKikakuSyurui & "'" & vbCrLf
    strSQL = strSQL & "                 AND CONVERT(CHAR(6),tenkai_kaisi_ym_date,112) = " & vbCrLf
    strSQL = strSQL & "                     (SELECT " & vbCrLf
    strSQL = strSQL & "                         CONVERT(CHAR(6),DATEADD(Month,-1,tenkai_kaisi_ym_date),112) " & vbCrLf
    strSQL = strSQL & "                     FROM " & vbCrLf
    strSQL = strSQL & "                         t_kikaku_jyouhou " & vbCrLf
    strSQL = strSQL & "                     WHERE " & vbCrLf
    strSQL = strSQL & "                         kikaku_syurui_cd = '" & lstrKikakuSyurui & "'" & vbCrLf
    strSQL = strSQL & "                         AND nen_ren_no = '" & lstrKikakuNenrenban & "'" & vbCrLf
    strSQL = strSQL & "                         AND eda_no = '" & lstrKikakuEdaban & "'" & vbCrLf
    strSQL = strSQL & "                     ) " & vbCrLf
    strSQL = strSQL & "             ) kikaku " & vbCrLf
    strSQL = strSQL & "         ON " & vbCrLf
    strSQL = strSQL & "             s_rireki.kikaku_syurui_cd = kikaku.kikaku_syurui_cd " & vbCrLf
    strSQL = strSQL & "             AND s_rireki.nen_ren_no = kikaku.nen_ren_no " & vbCrLf
    strSQL = strSQL & "             AND s_rireki.eda_no = kikaku.eda_no " & vbCrLf
    strSQL = strSQL & "     WHERE " & vbCrLf
    strSQL = strSQL & "         s_rireki.genka IS NOT NULL " & vbCrLf
    strSQL = strSQL & "         OR s_rireki.hattyuu_tani IS NOT NULL " & vbCrLf
    strSQL = strSQL & "     ) uni " & vbCrLf
    strSQL = strSQL & " WHERE " & vbCrLf
    strSQL = strSQL & "     jan_cd IN(" & pCsvJanCodeList & ") " & vbCrLf

        pTChkkikakuRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkkikakuRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    
                
    '配列の変更日項目を上記内容によって、変更する。
    strSQL = ""
    strSQL = strSQL & " SELECT " & vbCrLf
    strSQL = strSQL & "     TOP 1 " & vbCrLf
    strSQL = strSQL & "     tan.henkou_date_sinki " & vbCrLf
    strSQL = strSQL & "     ,tan.henkou_date_keizoku " & vbCrLf
    strSQL = strSQL & " FROM " & vbCrLf
    strSQL = strSQL & "     t_marutoku_kikaku_tantousya tan WITH (READCOMMITTED) " & vbCrLf
    strSQL = strSQL & "     INNER JOIN " & vbCrLf
    strSQL = strSQL & "         (SELECT " & vbCrLf
    strSQL = strSQL & "             CONVERT(CHAR(6),tenkai_kaisi_ym_date,112) nengetu " & vbCrLf
    strSQL = strSQL & "             ,kikaku_syurui_cd " & vbCrLf
    strSQL = strSQL & "         FROM " & vbCrLf
    strSQL = strSQL & "             t_kikaku_jyouhou WITH (READCOMMITTED) " & vbCrLf
    strSQL = strSQL & "         WHERE " & vbCrLf
    strSQL = strSQL & "             kikaku_syurui_cd ='" & lstrKikakuSyurui & "'" & vbCrLf
    strSQL = strSQL & "             AND nen_ren_no = '" & lstrKikakuNenrenban & "'" & vbCrLf
    strSQL = strSQL & "             AND eda_no ='" & lstrKikakuEdaban & "'" & vbCrLf
    strSQL = strSQL & "         ) jyou " & vbCrLf
    strSQL = strSQL & "     ON " & vbCrLf
    strSQL = strSQL & "         tan.nengetu = jyou.nengetu " & vbCrLf
    strSQL = strSQL & "         AND tan.kikaku_syurui_cd = jyou.kikaku_syurui_cd " & vbCrLf

    'strSQL = strSQL & " SELECT TOP 1 " & vbCrLf
    'strSQL = strSQL & " MKT.henkou_date_sinki " & vbCrLf
    'strSQL = strSQL & " ,MKT.henkou_date_keizoku " & vbCrLf
    'strSQL = strSQL & " FROM t_marutoku_kikaku_tantousya MKT WITH(READCOMMITTED) " & vbCrLf
    'strSQL = strSQL & " LEFT JOIN t_kikaku_jyouhou TKJ WITH(READCOMMITTED)  " & vbCrLf
    'strSQL = strSQL & " ON TKJ.kikaku_syurui_cd = MKT.kikaku_syurui_cd " & vbCrLf
    'strSQL = strSQL & " AND left(CONVERT(CHAR,TKJ.tenkai_kaisi_ym_date,112),6) = MKT.nengetu " & vbCrLf
      
    Dim marutokuRS As ADODB.Recordset
        marutokuRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    
    If Not marutokuRS.EOF Then
        '継続商品の場合、変更日(継続)（kikaku_mei_hosoku）で配列の変更日項目を変更する。
         pCsvhenkouDateKeizoku = marutokuRS.Fields("henkou_date_keizoku").Value
            
        '新規商品の場合、変更日(新規)（ henkou_date_sinki ）で配列の変更日項目を変更する。
         pCsvHenkouDate = marutokuRS.Fields("henkou_date_sinki").Value
    Else
        pCsvhenkouDateKeizoku = ""
        pCsvHenkouDate = ""
    End If
        

    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If

    If Not baikaRS Is Nothing Then
            baikaRS = Nothing
    End If

    If Not tenCdRS Is Nothing Then
            tenCdRS = Nothing
    End If
    
    If Not nengetuRS Is Nothing Then
            nengetuRS = Nothing
    End If
    
'    If Not kikakuRS Is Nothing Then
'            kikakuRS = Nothing
'    End If
    
    If Not marutokuRS Is Nothing Then
            marutokuRS = Nothing
    End If
    
    'GetObjectContext.SetComplete
  
    Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
    
    If Not baikaRS Is Nothing Then
            baikaRS = Nothing
    End If

    If Not tenCdRS Is Nothing Then
            tenCdRS = Nothing
    End If
    
    If Not tenkaiDateRS Is Nothing Then
            tenkaiDateRS = Nothing
    End If
    
        
    If Not nengetuRS Is Nothing Then
            nengetuRS = Nothing
    End If
'    If Not kikakuRS Is Nothing Then
'            kikakuRS = Nothing
'    End If

    
    If Not marutokuRS Is Nothing Then
            marutokuRS = Nothing
    End If
    
    'If Not GetObjectContext Is Nothing Then
        'GetObjectContext.SetAbort
        'End If

    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
            'Err.Number, Err.Source, Err.Description
End Sub
  '********** 2005/05/16 土田 仮作成 **********↓
Private Sub subIns_SetAllTables(ByVal arryData As Object)
        Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim strMethodCallInfo As String
    Dim lngCount As Long
    Dim aryJanCd() As String
 
    Dim strNVHFilter As String
    Dim rsNVH As ADODB.Recordset
 
'********** 2005/05/18 土田 テスト追加 **********↓
'大域変数の値をローカルにコピーし使用する

    Dim lstrKikakuSyurui As String
    lstrKikakuSyurui = pstrKikakuSyurui
    
    Dim lstrKikakuNenrenban As String
    lstrKikakuNenrenban = pstrKikakuNenrenban
    
    Dim lstrKikakuEdaban As String
    lstrKikakuEdaban = pstrKikakuEdaban

    Dim lstrTantousyaId As String
    lstrTantousyaId = pstrTantousyaId

'日付を始めに取得
    Dim lstrNow As Object
    lstrNow = Format(Now(), "yyyy/mm/dd hh:nn:ss")
    
'********** 2005/05/18 土田 テスト追加 **********上

    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables()"

    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    '---------------
    '   内部処理
    '---------------
    On Error GoTo DefectErrorHandler

    'RS作成用SQLの条件部分（企画種類コード、年連番、枝番の指定）
    Dim strWhereSQL As String
    strWhereSQL = "WHERE kikaku_syurui_cd = '" & lstrKikakuSyurui & "' AND nen_ren_no ='" & lstrKikakuNenrenban & "' AND eda_no ='" & lstrKikakuEdaban & "' "

'********** 2005/05/18 土田 追加 **********↓
     'RS作成用SQLの条件部分（変更日付の範囲指定）
    Dim strHenkouSQL As String
    strHenkouSQL = " AND henkou_date BETWEEN '" & pstrMinDate & "' AND '" & pstrMaxDate & "' "
'********** 2005/05/18 土田 追加 **********↑

    'RS作成用SQL
    Dim strSQL  As String

    '売価変更情報RS作成
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(売価変更情報RS作成)"
    strSQL = "SELECT * FROM " & TBL_BAI_HEN & " with (UPDLOCK) " & strWhereSQL
'********** 2005/05/18 土田 追加 **********↓
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") " & strHenkouSQL
'********** 2005/05/18 土田 追加 **********↑

    Dim baikaRS As ADODB.Recordset
        baikaRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
        'baikaRS.Fields("ten_cd").Properties("Optimize") = True 'オプティマイザ
    
    '仕入変更情報RS作成
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(仕入変更情報RS作成)"
    strSQL = "SELECT * FROM " & TBL_SII_HEN & " with (UPDLOCK) " & strWhereSQL
'********** 2005/05/18 土田 追加 **********↓
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") " & strHenkouSQL
'********** 2005/05/18 土田 追加 **********↑
    
    Dim siireRS As ADODB.Recordset
        siireRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    siireRS.Fields("ten_cd").Properties("Optimize") = True 'オプティマイザ
    
    'センターコード変更情報RS作成
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(センターコード変更情報RS作成)"
    strSQL = "SELECT * FROM " & TBL_CEN_HEN & " with (UPDLOCK) " & strWhereSQL
'********** 2005/05/18 土田 追加 **********↓
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") " & strHenkouSQL
'********** 2005/05/18 土田 追加 **********↑
    
    Dim centerRS As ADODB.Recordset
        centerRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    centerRS.Fields("ten_cd").Properties("Optimize") = True 'オプティマイザ
    
    'WF管理項目RS作成
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(WF管理項目RS作成)"
    strSQL = "SELECT * FROM t_syousen_wf_kanri with (UPDLOCK) " & strWhereSQL
    
    Dim wfRS As ADODB.Recordset
        wfRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    wfRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    
'********** 2005/05/17 土田 削除 **********↓
'    '主キーRS生成
'    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(主キーRS生成)"
'    strSQL = "SELECT key_cd , jan_cd FROM v_key_1 WITH(readcommitted)"
'
'    Dim vk1RS As ADODB.Recordset
'        vk1RS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
'    vk1RS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
'********** 2005/05/17 土田 削除 **********↑


'********** 2005/05/17 土田 追加 **********↓
    'VSWF表示仕様テーブルRS
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(VSWF表示仕様テーブルRS生成)"
    strSQL = _
    "SELECT *" & _
    "  FROM t_syousen_hyouji_siyou WITH(UPDLOCK)" & _
    strWhereSQL
    
    Dim hyoujiSiyouRS As ADODB.Recordset
        hyoujiSiyouRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    hyoujiSiyouRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
         
    'VMST一般基本テーブルRS
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(VMST一般基本テーブルRS生成)"
    strSQL = _
    "SELECT v_key_1.jan_cd, v_ippan_kihon.*" & _
    "  FROM v_ippan_kihon WITH(READCOMMITTED)" & _
    " INNER JOIN v_key_1 WITH(READCOMMITTED) ON v_ippan_kihon.key_cd = v_key_1.key_cd" & _
    " AND v_key_1.jan_cd IN(" & pCsvJanCodeList & ") "
    '↑janコードでの絞込み追加
    
    Dim ippanKihonRS As ADODB.Recordset
        ippanKihonRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    ippanKihonRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
        
    'センター区分を元にして、センターコードを取得する為のRS
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(センター区分を元にして、センターコードを取得する為のRS生成)"
    strSQL = _
    "SELECT jigyousyo_cd, center_cd " & _
    "  FROM v_jigyousyo WITH(READCOMMITTED)"
    
    Dim centerCDRS As ADODB.Recordset
        centerCDRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    centerCDRS.Fields("jigyousyo_cd").Properties("Optimize") = True 'オプティマイザ
    
    
'********** 2005/05/17 土田 追加 **********↑

    'WF更新用フラグ(WF更新時の効率化の為Stringで"0"又は"1"を持つ)
    Dim strBaikaAdded As String     '売価変更RSが更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    Dim strSiireAdded As String     '仕入変更RSが更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    Dim strCenterAdded As String    'センターコード変更RSが更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    
    ' ************** 1レコードずつRSを更新 *******************
    
    '********** 2008/06/04 ラン峰（大連） 追加 **********↓
    Dim strBaikaFlg As String       '売価が更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    Dim strSiireFlg As String       '仕入が更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    Dim strCenterFlg As String    'センターコードが更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    '********** 2008/06/04 ラン峰（大連） 追加 **********↑
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(1レコードずつRSを更新 開始)"
    Dim iRecord As Long
'    For iRecord = 0 To UBound(arryData)
    For iRecord = LBound(arryData) To UBound(arryData)
    
        'WF更新用フラグの初期化
        strBaikaAdded = "0"
        strSiireAdded = "0"
        strCenterAdded = "0"
        
        '********** 2008/05/27 ラン峰（大連） 追加 **********↓
        '売価、仕入情報、センターコードが更新(新規又は更新)されたかのフラグ
        strBaikaFlg = "0"
        strSiireFlg = "0"
        strCenterFlg = "0"
        '********** 2008/05/27 ラン峰（大連） 追加 **********↑
        '売価、仕入、センターコードRSをJanCDと変更日でフィルタリング
        Dim strFilter As String
        strFilter = "jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "' AND henkou_date = '" & arryData(iRecord)(enmTenColNo.HenkouDate) & "'"
        
        baikaRS.Filter = strFilter
        siireRS.Filter = strFilter
        centerRS.Filter = strFilter
        
        'データ更新対象店コードの取得
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(データ更新対象店コードの取得)"
        
        'このレコードの店コード
        Dim TenCd As String
        TenCd = arryData(iRecord)(enmTenColNo.TenCd)
        
        'データ更新対象店コードの取得に該当事業コードチェック用RSを再利用する
        With pTChkJigyouRS
            'フィルターを解除しておく
            .Filter = adFilterNone
'********** 2005/05/17 土田 削除 **********↓
'            .MoveFirst
'********** 2005/05/17 土田 削除 **********↑
            
'********** 2005/05/20 土田 変更 **********↓
'本部("0000")の場合もフィルタリングしない(下の"データ更新対象店コードのループ"に入りたい為）
'            If tenCD = ALL_TEN_CD Or tenCD = ALL_JIG_CD Then
            If TenCd = ALL_TEN_CD Or TenCd = ALL_JIG_CD Or TenCd = HON_TEN_CD Then
                '全店、全事業、本部の場合なにもしない
'********** 2005/05/20 土田 変更 **********↑
            ElseIf Len(TenCd) = 2 Then
                '各事業の場合
                '該当事業コードチェック用RSを事業コードでフィルタリングする
                
                '***** 2008/09/19 和田(IC) 変更 *****↓
                strNVHFilter = ""
                strNVHFilter = strNVHFilter & "jigyou_cd = '" & TenCd & "'"
                If arryData(iRecord)(enmTenColNo.TenCd) = "01" Then
                        rsNVH = GetNVHInfo(arryData(iRecord)(enmTenColNo.janCD))
                    Do While Not rsNVH.EOF
                        strNVHFilter = strNVHFilter & " and jigyousyo_cd <> '" & rsNVH.Fields("ten_cd").Value & "'"
                        rsNVH.MoveNext
                    Loop
                    rsNVH.Close
                        rsNVH = Nothing
                End If
                
                .Filter = strNVHFilter
                '***** 2008/09/19 和田(IC) 変更 *****↑

            ElseIf Len(TenCd) = 4 Then

                '各店の場合、該当事業コードチェック用RSを店コードでフィルタリングする
                .Filter = "jigyousyo_cd = '" & TenCd & "'"
            End If
        End With
        
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(データ更新対象店コードのループ 開始) "
' ************** データ更新対象店コードのループ *******************
'        Dim iTenCD As Integer
'        For iTenCD = LBound(arryTenCD) To UBound(arryTenCD)

        Dim strTenCd As String
        Dim i As Integer
        Do Until pTChkJigyouRS.EOF
           
            
'********** 2005/05/20 土田 変更 **********↓
'本部("0000")の場合は店コードに"0000"を直接入れる(レコードセットからは取得しない)
'            strTenCD = pTChkJigyouRS.Fields("jigyousyo_cd")
            If (TenCd = HON_TEN_CD) Then
                strTenCd = HON_TEN_CD
            Else
                strTenCd = pTChkJigyouRS.Fields("jigyousyo_cd")
            
            End If
'********** 2005/05/20 土田 変更 **********↑
                        
           '売価、売価変更理由のいずれかがある場合、売価変更情報RSに追加か更新
            If arryData(iRecord)(enmTenColNo.Baika) <> "" Or arryData(iRecord)(enmTenColNo.BaikaRiyuu) <> "" Then
'                'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(売価変更情報RSに追加か更新) " & strTenCD
                '売価変更情報RSに追加か更新
                With baikaRS
                    
'                    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(売価変更情報RSに追加か更新) movefirst: " & strTenCD
        '********** 2005/05/18 土田 変更 **********↓
        '            If Not .EOF Then
                    If Not .BOF Then
        '********** 2005/05/18 土田 変更 **********↑
                        .MoveFirst
'                        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(売価変更情報RSに追加か更新) find: " & strTenCD
                        .Find ("ten_cd = '" & strTenCd & "'")
                        
                    End If
                    
                    
                    If .EOF Then
'                        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(新規追加)"
                        '新規追加
                        .AddNew
                        .Fields("kikaku_syurui_cd").Value = lstrKikakuSyurui
                        .Fields("nen_ren_no").Value = lstrKikakuNenrenban
                        .Fields("eda_no").Value = lstrKikakuEdaban
                        .Fields("jan_cd").Value = arryData(iRecord)(enmTenColNo.janCD)
'                        .Fields("ten_cd").Value = arryData(iRecord)(enmTenColNo.tenCD)
                        .Fields("ten_cd").Value = strTenCd
                        .Fields("henkou_date").Value = arryData(iRecord)(enmTenColNo.HenkouDate)
                      
                        '売価
                        '********** 2008/05/27 ラン峰（大連） 追加 **********↓
                        '商品選定WF(変更除外店舗)テーブルの｢売価除外区分｣フィールドは｢除外｣になると
                        'CSVから取得した｢売価｣が登録・更新されない(NULLにセット)。
                                                                                            
                        '商品選定WF(変更除外店舗)テーブルの｢売価除外区分｣フィールドは｢除外｣ではなければ
                        '商品選定DB（売価変更情報）テーブルへCSVから取得した｢売価｣が登録・更新される。
                       
                        If (strTenCd <> HON_TEN_CD) And _
                            (pTChkJigyouRS.Fields("baika_jyogai_kbn").Value = "1") Then
                        '商品選定WF(変更除外店舗)テーブルの｢売価除外区分｣フィールドは｢除外｣になると
                        'NULLにセット
                            .Fields("baika").Value = dbNull.value
                        Else
                        
                        '********** 2008/05/27 ラン峰（大連） 追加 **********↑
                            If arryData(iRecord)(enmTenColNo.Baika) <> "" Then
                                .Fields("baika").Value = CLng(arryData(iRecord)(enmTenColNo.Baika))
                                strBaikaFlg = "1"
    '********** 2005/05/19 土田 追加 **********↓
                            Else
                                .Fields("baika").Value = dbNull.value
    '********** 2005/05/19 土田 追加 **********↑
                            End If
                        End If
                        
                        
                        '売価変更理由コード
'********** 2005/05/19 土田 追加 **********↓
'                        .Fields("baika_henkou_riyuu_cd").Value = arryData(iRecord)(enmTenColNo.BaikaRiyuu)
                        If arryData(iRecord)(enmTenColNo.BaikaRiyuu) <> "" Then
                            .Fields("baika_henkou_riyuu_cd").Value = arryData(iRecord)(enmTenColNo.BaikaRiyuu)
                        Else
                            .Fields("baika_henkou_riyuu_cd").Value = dbNull.value
                        End If
'********** 2005/05/19 土田 追加 **********↑
'********** 2005/06/15 北澤 **********↓
'                        .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
'売価先付変更情報が基幹DBに既存の場合は登録区分｢2｣をセットする(承認されると先付変更情報を上書するデータとなる)
                        If InStr(1, arrTourokuKbn2(iRecord, conColBaika), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
'********** 2005/06/15 北澤 **********↑
                        .Fields("touroku_user_id").Value = lstrTantousyaId
                        .Fields("touroku_ymd_hms").Value = lstrNow
                        
                    Else
'                        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(更新処理)"
                        ' 更新処理
                        '売価
                        '********** 2008/05/27 ラン峰（大連） 追加 **********↓
                        '商品選定WF(変更除外店舗)テーブルの｢売価除外区分｣フィールドは｢除外｣になると
                        'CSVから取得した｢売価｣が登録・更新されない(NULLにセット)。
                                                                                            
                        '商品選定WF(変更除外店舗)テーブルの｢売価除外区分｣フィールドは｢除外｣ではなければ
                        '商品選定DB（売価変更情報）テーブルへCSVから取得した｢売価｣が登録・更新される。
                       
                       
                        If (strTenCd <> HON_TEN_CD) And _
                            (pTChkJigyouRS.Fields("baika_jyogai_kbn").Value = "1") Then
                       
                        '商品選定WF(変更除外店舗)テーブルの｢売価除外区分｣フィールドは｢除外｣になると
                        'NULLにセット
                            .Fields("baika").Value = dbNull.value
                        Else
                    '********** 2008/05/27 ラン峰（大連） 追加 **********↑
                            If arryData(iRecord)(enmTenColNo.Baika) <> "" Then
                                .Fields("baika").Value = CLng(arryData(iRecord)(enmTenColNo.Baika))
                                strBaikaFlg = "1"
                            End If
                        End If
                       
                        '売価変更理由コード
                        If arryData(iRecord)(enmTenColNo.BaikaRiyuu) <> "" Then
                            .Fields("baika_henkou_riyuu_cd").Value = arryData(iRecord)(enmTenColNo.BaikaRiyuu)
                        End If
'********** 2005/06/15 北澤 **********↓
'売価先付変更情報が基幹DBに既存の場合は登録区分｢2｣をセットする(承認されると先付変更情報を上書するデータとなる)
                        If InStr(1, arrTourokuKbn2(iRecord, conColBaika), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
'********** 2005/06/15 北澤 **********↑
                        .Fields("kousin_user_id").Value = lstrTantousyaId
                        .Fields("kousin_ymd_hms").Value = lstrNow
                    End If
                    .Update
                   
                End With
                strBaikaAdded = "1"

            End If
            
            
'            'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(原価、取引先、発注単位のどれかがあるかチェックする。)"
            '原価、取引先、発注単位のどれかがあるかチェックする。
            If arryData(iRecord)(enmTenColNo.Genka) <> "" _
                Or arryData(iRecord)(enmTenColNo.Torihikisaki) <> "" _
                Or arryData(iRecord)(enmTenColNo.Hacchuutani) <> "" Then
                
'                'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(仕入変更情報RS更新)"
                '仕入変更情報RS更新
                With siireRS
'********** 2005/05/18 土田 変更 **********↓
'                   If Not .EOF Then
                   If Not .BOF Then
'********** 2005/05/18 土田 変更 **********↑
                        .MoveFirst
                        .Find ("ten_cd = '" & strTenCd & "'")
                    End If
                    If .EOF Then
                         '新規追加
                        .AddNew
                        .Fields("kikaku_syurui_cd").Value = lstrKikakuSyurui
                        .Fields("nen_ren_no").Value = lstrKikakuNenrenban
                        .Fields("eda_no").Value = lstrKikakuEdaban
                        .Fields("jan_cd").Value = arryData(iRecord)(enmTenColNo.janCD)
'                        .Fields("ten_cd").Value = arryData(iRecord)(enmTenColNo.tenCD)
                        .Fields("ten_cd").Value = strTenCd
                        .Fields("henkou_date").Value = arryData(iRecord)(enmTenColNo.HenkouDate)
                        
                        '取引先コード
'********** 2005/05/19 土田 変更 **********↓
'                        .Fields("torihikisaki_cd").Value = arryData(iRecord)(enmTenColNo.Torihikisaki)
                        '********** 2008/05/28 ラン峰（大連） 追加 **********↓
                        '商品選定WF(変更除外店舗)テーブルの｢取引先除外区分｣は｢除外｣ではない場合
                        '取引先コードは登録される。
                       
                        If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("torihikisaki_jyogai_kbn").Value = "1" Then
                       
                            .Fields("torihikisaki_cd").Value = dbNull.value
                            
                        Else
                    '********** 2008/05/28 ラン峰（大連） 追加 **********↑
                            If arryData(iRecord)(enmTenColNo.Torihikisaki) <> "" Then
                                .Fields("torihikisaki_cd").Value = arryData(iRecord)(enmTenColNo.Torihikisaki)
                                strSiireFlg = "1"
                            Else
                                .Fields("torihikisaki_cd").Value = dbNull.value
                            End If
                        End If
                       
                        
                    
                        
'********** 2005/05/19 土田 変更 **********↑
                        '発注単位
                         '********** 2008/05/28 ラン峰（大連） 追加 **********↓
                         '商品選定WF(変更除外店舗)テーブルの｢発注単位除外区分｣は｢除外｣ではない場合
                         '発注単位は登録される。
                 
                        If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("hattyuu_tani_jyogai_kbn").Value = "1" Then
                       
                            .Fields("hattyuu_tani").Value = dbNull.value
                            
                        Else
                            '********** 2008/05/28 ラン峰（大連） 追加 **********↑
                            If arryData(iRecord)(enmTenColNo.Hacchuutani) <> "" Then
                                .Fields("hattyuu_tani").Value = CLng(arryData(iRecord)(enmTenColNo.Hacchuutani))
     '********** 2005/05/19 土田 追加 **********↓
                                strSiireFlg = "1"
                            Else
                                .Fields("hattyuu_tani").Value = dbNull.value
    '********** 2005/05/19 土田 追加 **********↑
                            End If
                         
                        End If
                        
                        '原価
                         '********** 2008/05/28 ラン峰（大連） 追加 **********↓
                         '商品選定WF(変更除外店舗)テーブルの｢原価除外区分｣は｢除外｣ではない場合
                         '原価は登録される。
                         
                        If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("genka_jyogai_kbn").Value = "1" Then
                            .Fields("genka").Value = dbNull.value
                        Else
                        '********** 2008/05/28 ラン峰（大連） 追加 **********↑
                            If arryData(iRecord)(enmTenColNo.Genka) <> "" Then
    '********** 2005/05/19 土田 変更 **********↓
    '                            .Fields("genka").Value = CLng(arryData(iRecord)(enmTenColNo.Genka))
                                .Fields("genka").Value = CDbl(arryData(iRecord)(enmTenColNo.Genka))
    '********** 2005/05/19 土田 追加 **********↓
                                strSiireFlg = "1"
                            Else
                                .Fields("genka").Value = dbNull.value
    '********** 2005/05/19 土田 追加 **********↑
    '********** 2005/05/19 土田 変更 **********↑
                            End If
                        End If
                       
'********** 2005/06/15 北澤 **********↓
'                        .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
'仕入先付変更情報が基幹DBに既存の場合は登録区分｢2｣をセットする(承認されると先付変更情報を上書するデータとなる)
                        If InStr(1, arrTourokuKbn2(iRecord, conColSiire), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
'********** 2005/06/15 北澤 **********↑
                        .Fields("touroku_user_id").Value = lstrTantousyaId
                        .Fields("touroku_ymd_hms").Value = lstrNow
                    Else
                        '更新処理
                        '取引先コード
                        '********** 2008/05/28 ラン峰（大連） 追加 **********↓
                        '商品選定WF(変更除外店舗)テーブルの｢取引先除外区分｣は｢除外｣ではない場合
                        '取引先コードは更新される。
                        If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("torihikisaki_jyogai_kbn").Value = "1" Then
                    
                            .Fields("torihikisaki_cd").Value = dbNull.value
                        Else
                        
                            If arryData(iRecord)(enmTenColNo.Torihikisaki) <> "" Then
    '********** 2005/05/19 土田 変更 **********↓
    '                            .Fields("torihikisaki_cd").Value = CLng(arryData(iRecord)(enmTenColNo.Torihikisaki))
                                 .Fields("torihikisaki_cd").Value = arryData(iRecord)(enmTenColNo.Torihikisaki)
    '********** 2005/05/19 土田 変更 **********↑
                                 strSiireFlg = "1"
                            End If
                        End If
                        
                        '********** 2008/05/28 ラン峰（大連） 追加 **********↑
                        '発注単位
                        '********** 2008/05/28 ラン峰（大連） 追加 **********↓
                        '商品選定WF(変更除外店舗)テーブルの｢発注単位除外区分｣は｢除外｣ではない場合
                        '発注単位は更新される。
                        If (strTenCd <> HON_TEN_CD) And _
                                pTChkJigyouRS.Fields("hattyuu_tani_jyogai_kbn").Value = "1" Then
                            .Fields("hattyuu_tani").Value = dbNull.value
                        Else
                            If arryData(iRecord)(enmTenColNo.Hacchuutani) <> "" Then
                                .Fields("hattyuu_tani").Value = arryData(iRecord)(enmTenColNo.Hacchuutani)
                                strSiireFlg = "1"
                            End If
                        End If
                        
                        
                        '********** 2008/05/28 ラン峰（大連） 追加 **********↑
                        
                        '原価
                        '********** 2008/05/28 ラン峰（大連） 追加 **********↓
                        '商品選定WF(変更除外店舗)テーブルの｢原価除外区分｣は｢除外｣ではない場合
                        '原価は更新される。
                         If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("genka_jyogai_kbn").Value = "1" Then
                            .Fields("genka").Value = dbNull.value
                                
                        Else
                            
                            If arryData(iRecord)(enmTenColNo.Genka) <> "" Then
    '********** 2005/05/19 土田 変更 **********↓
    '                            .Fields("genka").Value = arryData(iRecord)(enmTenColNo.Genka)
                                .Fields("genka").Value = CDbl(arryData(iRecord)(enmTenColNo.Genka))
    '********** 2005/05/19 土田 変更 **********↑
                                strSiireFlg = "1"
                            End If
                        End If
                        
                        
'********** 2005/06/15 北澤 **********↓
'仕入先付変更情報が基幹DBに既存の場合は登録区分｢2｣をセットする(承認されると先付変更情報を上書するデータとなる)
                        If InStr(1, arrTourokuKbn2(iRecord, conColSiire), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
'********** 2005/06/15 北澤 **********↑
                        .Fields("kousin_user_id").Value = lstrTantousyaId
                        .Fields("kousin_ymd_hms").Value = lstrNow
                   End If
                End With
                strSiireAdded = "1"
            End If
        
'            'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(センター区分があるかチェックする)"
            'センター区分があるかチェックする。
            If arryData(iRecord)(enmTenColNo.CenterKbn) <> "" Then
'                'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(センターコード変更情報RS更新)"
                'センターコード変更情報RS更新
                
                'センターコードを決める：
                'センター区分が0の場合はNULL、1の場合は'0151'、2の場合は事業所マスタから取得する
'********** 2005/05/19 土田 変更 **********↓
'   センター区分が０の場合は""ではなくNULLを入れるよう修正
'                Dim strCenterCD As String
                Dim centerCD
                Select Case arryData(iRecord)(enmTenColNo.CenterKbn)
                    Case "0":       'センター区分が0の場合、Nullをセット
'                        strCenterCD = ""
                        centerCD = dbNull.value
                    
                    '2015/09/15 wadak7 upd
                    'Case "1":       'センター区分が1の場合、0151(児玉センター)をセット
                    Case "1", "3":      'センター区分が1,3の場合
                        'centerCD = CENTERCD_KODAMA
                        '登録票のセンターコードをセット
                        centerCD = arryData(iRecord)(enmModColNo.centerCD)
                    Case "2":       'センター区分が2の場合、店コードを元に事業所マスタから取得したセンターコードをセット
                        With centerCDRS
'********** 2005/05/19 土田 変更 **********↓
'                            If Not .EOF Then
                            If Not .BOF Then
'********** 2005/05/19 土田 変更 **********↑
                                .MoveFirst
                                .Find ("jigyousyo_cd = '" & strTenCd & "'")
                            End If
                        
                            'データが無かったらNullをセットする。
                            If .EOF Then
'                                strCenterCD = ""
                                centerCD = dbNull.value
                            Else
                                centerCD = centerCDRS.Fields("center_cd")
                            End If
                        End With
                        
                    Case Else:       'センター区分が0,1,2以外の場合、Nullをセット
'                       strCenterCD = ""
                        centerCD = dbNull.value
'********** 2005/05/19 土田 変更 **********↑
                        
                End Select
                
                With centerRS
'********** 2005/05/18 土田 変更 **********↓
'                    If Not .EOF Then
                    If Not .BOF Then
'********** 2005/05/18 土田 変更 **********↑
                        .MoveFirst
                        .Find ("ten_cd = '" & strTenCd & "'")
                    End If
                    If .EOF Then
                         '新規追加
                        .AddNew
                        .Fields("kikaku_syurui_cd").Value = lstrKikakuSyurui
                        .Fields("nen_ren_no").Value = lstrKikakuNenrenban
                        .Fields("eda_no").Value = lstrKikakuEdaban
                        .Fields("jan_cd").Value = arryData(iRecord)(enmTenColNo.janCD)
'                        .Fields("ten_cd").Value = arryData(iRecord)(enmTenColNo.tenCD)
                        .Fields("ten_cd").Value = strTenCd
                        .Fields("henkou_date").Value = arryData(iRecord)(enmTenColNo.HenkouDate)
                      
                        'センター区分
                        '********** 2008/05/28 ラン峰（大連） 追加 **********↓
                
               
                        If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("center_jyogai_kbn").Value = "1" Then
                        '商品選定WF(変更除外店舗)テーブルの｢センター除外区分｣は｢除外｣であれば、NULLにセット
                            .Fields("center_kbn").Value = dbNull.value
                        Else
                    '********** 2008/05/28 ラン峰（大連） 追加 **********↑
                            .Fields("center_kbn").Value = arryData(iRecord)(enmTenColNo.CenterKbn)
                            'センターコード：センター区分が0の場合はNULL、1の場合は'0151'、2の場合は事業所マスタから取得する
    '                        .Fields("center_cd").Value = FncGetCenterCd(arryData(iRecord)(enmTenColNo.CenterKbn), strTenCD)
    '********** 2005/05/18 土田 変更 **********↓
    '                        .Fields("center_cd").Value = strCenterCD
                            .Fields("center_cd").Value = centerCD
    '********** 2005/05/18 土田 変更 **********↑
                            strCenterFlg = "1"
                        End If
                       
'********** 2005/06/15 北澤 **********↓
'                        .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
'センター先付変更情報が基幹DBに既存の場合は登録区分｢2｣をセットする(承認されると先付変更情報を上書するデータとなる)
                        If InStr(1, arrTourokuKbn2(iRecord, conColCenter), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
'********** 2005/06/15 北澤 **********↑
                        .Fields("touroku_user_id").Value = lstrTantousyaId
                        .Fields("touroku_ymd_hms").Value = lstrNow
                    Else
                        ' 更新処理
                        'センター区分
                        '********** 2008/05/28 ラン峰（大連） 追加 **********↓
                        If (strTenCd <> HON_TEN_CD) And _
                                pTChkJigyouRS.Fields("center_jyogai_kbn").Value = "1" Then
                            '商品選定WF(変更除外店舗)テーブルの｢センター除外区分｣は｢除外｣であれば、NULLにセット
                                .Fields("center_kbn").Value = dbNull.value
                        Else
                    '********** 2008/05/28 ラン峰（大連） 追加 **********↑
                            If arryData(iRecord)(enmTenColNo.CenterKbn) <> "" Then
                                .Fields("center_kbn").Value = arryData(iRecord)(enmTenColNo.CenterKbn)
                                'センターコード：センター区分が0の場合はNULL、1の場合は'0151'、2の場合は事業所マスタから取得する
        '                        .Fields("center_cd").Value = FncGetCenterCd(arryData(iRecord)(enmTenColNo.CenterKbn), strTenCD)
    '********** 2005/05/18 土田 変更 **********↓
        '                        .Fields("center_cd").Value = strCenterCD
                                .Fields("center_cd").Value = centerCD
    '********** 2005/05/18 土田 変更 **********↑
                                strCenterFlg = "1"
                            End If
                        End If
                        
                        
'********** 2005/06/15 北澤 **********↓
'センター先付変更情報が基幹DBに既存の場合は登録区分｢2｣をセットする(承認されると先付変更情報を上書するデータとなる)
                        If InStr(1, arrTourokuKbn2(iRecord, conColCenter), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
'********** 2005/06/15 北澤 **********↑
                        .Fields("kousin_user_id").Value = lstrTantousyaId
                        .Fields("kousin_ymd_hms").Value = lstrNow
                    End If
                End With
                strCenterAdded = "1"
           End If
'********** 2005/05/20 土田 追加 **********↓
'本部("0000")の場合は一回でDoループを抜ける
            If (strTenCd = HON_TEN_CD) Then Exit Do
'********** 2005/05/20 土田 追加 **********↑
            pTChkJigyouRS.MoveNext
            
        Loop
            
        '売価、仕入、センターコードRSのいずれかで追加処理があった場合
        If (strBaikaAdded = "1" Or strCenterAdded = "1" Or strSiireAdded = "1") Then
'            'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(売価、仕入、センターコードRSのいずれかで追加処理があった場合)"


            'JANコードと紐付く主キーを取得する
            Dim strKey_cd
'********** 2005/05/17 土田 変更 **********↓
'            With vk1RS
            With pChkvk1JANCodeRS
'********** 2005/05/18 土田 変更 **********↓
'                If Not .EOF Then
                If Not .BOF Then
'********** 2005/05/18 土田 変更 **********↑
                    .MoveFirst
                    .Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
                    strKey_cd = .Fields("key_cd")
                End If
'********** 2005/05/17 土田 変更 **********↑
            End With

            
            'WF管理項目を最新情報に更新する
            With wfRS
'********** 2005/05/17 土田 変更 **********↓
'********** 2005/05/18 土田 変更 **********↓
'                If Not .EOF Then
                If Not .BOF Then
'********** 2005/05/18 土田 変更 **********↑
                    .MoveFirst
                    .Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
                End If
'********** 2005/05/17 土田 変更 **********↑
                'WF管理項目に無い場合は、新規追加
                If .EOF Then
                    .AddNew
                    .Fields("kikaku_syurui_cd").Value = lstrKikakuSyurui
                    .Fields("nen_ren_no").Value = lstrKikakuNenrenban
                    .Fields("eda_no").Value = lstrKikakuEdaban
                    .Fields("jan_cd").Value = arryData(iRecord)(enmTenColNo.janCD)
                    .Fields("jan_umu_kbn").Value = 1
                    .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                    .Fields("jyogai_kbn").Value = 0
                    .Fields("touroku_tantousya").Value = lstrTantousyaId
                    .Fields("keikakusuu").Value = 0
                    .Fields("kikaku_no").Value = lstrKikakuSyurui & lstrKikakuNenrenban & "-" & lstrKikakuEdaban
                    .Fields("kikaku_kahi_kbn").Value = 0
                    .Fields("kikaku_syouninsya_comment").Value = dbNull.value
                    .Fields("hinsyu_kahi_kbn").Value = 0
                    .Fields("hinsyu_syouninsya_comment").Value = dbNull.value
                    .Fields("naibu_cd").Value = strKey_cd
                    .Fields("hyouji_siyou_touroku_kbn").Value = 1
                    .Fields("sekisai_size_touroku_kbn").Value = 0
                    .Fields("pop_touroku_kbn").Value = 0
                    .Fields("syanai_kihon_touroku_kbn").Value = 0
                    .Fields("kowake_touroku_kbn").Value = 0
                    .Fields("jigyoubetu_touroku_kbn").Value = 0
                    .Fields("hanbai_touroku_kbn").Value = 0
                    .Fields("siire_touroku_kbn").Value = 0
                    .Fields("hinsyu_taikei_touroku_kbn").Value = 0
                    .Fields("uriba_touroku_kbn").Value = 0
                    '********** 2008/05/27 ラン峰（大連） 削除 **********↓
                    '売価変更情報登録区分
                    '.Fields("baika_henkou_touroku_kbn").Value = strBaikaAdded
                    '仕入変更情報登録区分
                    '.Fields("siire_henkou_touroku_kbn").Value = strSiireAdded
                    'ｾﾝﾀｰｺｰﾄﾞ変更情報登録区分
                    '.Fields("center_cd_henkou_touroku_kbn").Value = strCenterAdded
                    '********** 2008/05/27 ラン峰（大連） 削除 **********↑
                    
                    '売価変更情報登録区分
                    '********** 2008/05/27 ラン峰（大連） 追加 **********↓
                    If strBaikaFlg = "1" Then
                        .Fields("baika_henkou_touroku_kbn").Value = strBaikaAdded
                    Else
                        .Fields("baika_henkou_touroku_kbn").Value = 0
                    End If
                    '********** 2008/05/27 ラン峰（大連） 追加 **********↑
                    
                    '仕入変更情報登録区分
                    '********** 2008/05/27 ラン峰（大連） 追加 **********↓
                    If strSiireFlg = "1" Then
                        .Fields("siire_henkou_touroku_kbn").Value = strSiireAdded
                    Else
                        .Fields("siire_henkou_touroku_kbn").Value = 0
                    End If
                    
                    '********** 2008/05/27 ラン峰（大連） 追加 **********↑
                    
                    'ｾﾝﾀｰｺｰﾄﾞ変更情報登録区分
                    '********** 2008/05/27 ラン峰（大連） 追加 **********↓
                    If strCenterFlg = "1" Then
                        .Fields("center_cd_henkou_touroku_kbn").Value = strCenterAdded
                    Else
                        .Fields("center_cd_henkou_touroku_kbn").Value = 0
                    End If
                    
                    '********** 2008/05/27 ラン峰（大連） 追加 **********↑
                    .Fields("kowake_henkou_touroku_kbn").Value = 0
                    .Fields("hanki_henkou_touroku_kbn").Value = 0
                    .Fields("touroku_user_id").Value = lstrTantousyaId
                    .Fields("kousin_user_id").Value = dbNull.value
                    .Fields("touroku_ymd_hms").Value = lstrNow
                    .Fields("kousin_ymd_hms").Value = dbNull.value
        
                'WF管理項目にある場合は、更新
                Else
                    .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                    .Fields("touroku_tantousya").Value = lstrTantousyaId
                    .Fields("kikaku_no").Value = lstrKikakuSyurui & lstrKikakuNenrenban & "-" & lstrKikakuEdaban
                    .Fields("naibu_cd").Value = strKey_cd
                    .Fields("hyouji_siyou_touroku_kbn").Value = 1
'********** 2005/06/16 北澤 **********↓
'                    .Fields("baika_henkou_touroku_kbn").Value = strBaikaAdded
'                    .Fields("siire_henkou_touroku_kbn").Value = strSiireAdded
'                    .Fields("center_cd_henkou_touroku_kbn").Value = strCenterAdded
'更新時は、登録区分｢0｣をセットする事はあり得ない
                    '********** 2008/05/27 ラン峰（大連） 削除 **********↓
                    'If strBaikaAdded = "1" Then
                        '売価変更情報登録区分
                    '    .Fields("baika_henkou_touroku_kbn").Value = strBaikaAdded
                    'End If
                    'If strSiireAdded = "1" Then
                        '仕入変更情報登録区分
                    '    .Fields("siire_henkou_touroku_kbn").Value = strSiireAdded
                    'End If
                    'If strCenterAdded = "1" Then
                        'ｾﾝﾀｰｺｰﾄﾞ変更情報登録区分
                    '    .Fields("center_cd_henkou_touroku_kbn").Value = strCenterAdded
                    'End If
                    '********** 2008/05/27 ラン峰（大連） 削除 **********↑
                    
                    '********** 2008/05/27 ラン峰（大連） 追加 **********↓
                    If strBaikaFlg = "1" Then
                        If strBaikaAdded = "1" Then
                            '売価変更情報登録区分
                            .Fields("baika_henkou_touroku_kbn").Value = strBaikaAdded
                        End If
                    'Else
                    '    .Fields("baika_henkou_touroku_kbn").Value = dbNull.value
                        
                    End If
                    If strSiireFlg = "1" Then
                        If strSiireAdded = "1" Then
                            '仕入変更情報登録区分
                            .Fields("siire_henkou_touroku_kbn").Value = strSiireAdded
                        End If
                    'Else
                    '    .Fields("siire_henkou_touroku_kbn").Value = dbNull.value
                    End If
                    If strCenterFlg = "1" Then
                        If strCenterAdded = "1" Then
                            'ｾﾝﾀｰｺｰﾄﾞ変更情報登録区分
                            .Fields("center_cd_henkou_touroku_kbn").Value = strCenterAdded
                        End If
                    'Else
                    '    .Fields("center_cd_henkou_touroku_kbn").Value = dbNull.value
                    End If
                    '********** 2008/05/27 ラン峰（大連） 追加 **********↑
'********** 2005/06/16 北澤 **********↑
                    .Fields("kousin_user_id").Value = lstrTantousyaId
                    .Fields("kousin_ymd_hms").Value = lstrNow
                End If
                .Update
            End With
        End If
        
        '各RSのフィルター解除
'        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(各RSのフィルター解除)"
            baikaRS .Filter = adFilterNone
        siireRS.Filter = adFilterNone
        centerRS.Filter = adFilterNone
        
'********** 2005/05/17 土田 追加 **********↓
'VMST一般基本テーブルより商品情報を取得し、VSWF表示仕様テーブルにレコードを追加する
'        'strMethodCallInfo = CStr(iRecord) & "行目：VMST一般基本テーブルより商品情報を取得し、VSWF表示仕様テーブルにレコードを追加する"
        With hyoujiSiyouRS
'********** 2005/05/18 土田 変更 **********↓
'            If Not .EOF Then
            If Not .BOF Then
'********** 2005/05/18 土田 変更 **********↑
                .MoveFirst
                .Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
            End If
'            'strMethodCallInfo = CStr(iRecord) & "行目：hyoujiSiyouRS"
            
            'VSWF表示仕様テーブルにレコードが存在しない場合はVMST一般基本テーブルより商品情報を取得しVSWF表示仕様テーブルにレコードを追加する
            If .EOF Then
'                'strMethodCallInfo = CStr(iRecord) & "行目：VMST一般基本テーブルより商品情報を取得し"
                With ippanKihonRS
'********** 2005/05/18 土田 変更 **********↓
'                    If Not .EOF Then
                    If Not .BOF Then
'********** 2005/05/18 土田 変更 **********↑
                        .MoveFirst
                        .Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
                    End If
'                    'strMethodCallInfo = CStr(iRecord) & "行目：jan_cdでＦＩＮＤ"
                    If .EOF Then
'                        'strMethodCallInfo = CStr(iRecord) & "行目：VMST一般基本テーブルにデータが存在しない為VSWF表示仕様テーブルにレコードを追加出来ませんでした。"
                        GoTo DefectErrorHandler
                    Else
                        'VSWF表示仕様テーブルRSを更新する
'                        'strMethodCallInfo = CStr(iRecord) & "行目：VSWF表示仕様テーブルRSを更新する"
                        With hyoujiSiyouRS
                            .AddNew
                            .Fields("kikaku_syurui_cd") = lstrKikakuSyurui
                            .Fields("nen_ren_no") = lstrKikakuNenrenban
                            .Fields("eda_no") = lstrKikakuEdaban
                            .Fields("jan_cd") = arryData(iRecord)(enmTenColNo.janCD)
                            .Fields("hyoujiyou_syouhin_mei_kana") = ippanKihonRS("hyoujiyou_syouhin_mei_kana")
                            .Fields("hyoujiyou_syouhin_mei_kanji") = ippanKihonRS("hyoujiyou_syouhin_mei_kanji")
                            .Fields("receipt_mei_kana") = ippanKihonRS("receipt_mei_kana")
                            .Fields("receipt_mei_kanji") = ippanKihonRS("receipt_mei_kanji")
                            .Fields("syouhin_mei_kanji") = ippanKihonRS("syouhin_mei_kanji")
                            .Fields("maker_mei_kanji") = ippanKihonRS("maker_mei_kanji")
                            .Fields("kikaku_size_kanji") = ippanKihonRS("kikaku_size_kanji")
                            .Fields("sina_no_kata_no_kanji") = ippanKihonRS("sina_no_kata_no_kanji")
                            .Fields("iro_gara_kanji") = ippanKihonRS("iro_gara_kanji")
                            .Fields("taste_kanji") = ippanKihonRS("taste_kanji")
                            .Fields("seisanti_mei_kanji") = ippanKihonRS("seisanti_mei_kanji")
                            .Fields("tokutyou") = ippanKihonRS("tokutyou")
                            .Fields("syoumi_siyoukigen_um_kbn") = ippanKihonRS("syoumi_siyoukigen_um_kbn")
                            .Fields("eco_syouhin_kbn") = ippanKihonRS("eco_syouhin_kbn")
                            .Fields("green_tekigou_syouhin_kbn") = ippanKihonRS("green_tekigou_syouhin_kbn")
                            .Fields("gpn_db_keisai_syouhin_kbn") = ippanKihonRS("gpn_db_keisai_syouhin_kbn")
                            .Fields("touroku_user_id") = lstrTantousyaId
                            .Fields("touroku_ymd_hms") = lstrNow
                            .Update
                        End With
                    End If
                End With
            End If
        End With
'********** 2005/05/17 土田 追加 **********↑
    Next iRecord

    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(RSのバッチ更新0)"
    'RSのバッチ更新
    Call objCDBAccess.UpdateDCRecordset("VSWF", baikaRS)
    Call objCDBAccess.UpdateDCRecordset("VSWF", siireRS)
    Call objCDBAccess.UpdateDCRecordset("VSWF", centerRS)
    Call objCDBAccess.UpdateDCRecordset("VSWF", wfRS)
'********** 2005/05/17 土田 追加 **********↓
    Call objCDBAccess.UpdateDCRecordset("VSWF", hyoujiSiyouRS)
'********** 2005/05/17 土田 追加 **********↑
    
    'GetObjectContext.SetComplete
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(オブジェクト開放)"
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If

    If Not baikaRS Is Nothing Then
            baikaRS = Nothing
    End If

    If Not siireRS Is Nothing Then
            siireRS = Nothing
    End If

    If Not centerRS Is Nothing Then
            centerRS = Nothing
    End If

    If Not wfRS Is Nothing Then
            wfRS = Nothing
    End If

'********** 2005/05/17 土田 削除 **********↓
'    If Not vk1RS Is Nothing Then
'            vk1RS = Nothing
'    End If
'********** 2005/05/17 土田 削除 **********↑

'********** 2005/05/17 土田 追加 **********↓
    If Not hyoujiSiyouRS Is Nothing Then
            hyoujiSiyouRS = Nothing
    End If

    If Not ippanKihonRS Is Nothing Then
            ippanKihonRS = Nothing
    End If
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(登録時利用したのでここで開放する)"
'登録時利用したのでここで開放する(作成はFncTenTourokuで行っている)
    If Not pTChkJigyouRS Is Nothing Then
        pTChkJigyouRS.Close
            pTChkJigyouRS = Nothing
    End If
    
'登録時利用したのでここで開放する(作成はFncTenTourokuで行っている)
    If Not pChkvk1JANCodeRS Is Nothing Then
        pChkvk1JANCodeRS.Close
            pChkvk1JANCodeRS = Nothing
    End If
    
    If Not centerCDRS Is Nothing Then
        centerCDRS.Close
            centerCDRS = Nothing
    End If
    
'********** 2005/05/17 土田 追加 **********↑
    Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If

    If Not baikaRS Is Nothing Then
            baikaRS = Nothing
    End If

    If Not siireRS Is Nothing Then
            siireRS = Nothing
    End If

    If Not centerRS Is Nothing Then
            centerRS = Nothing
    End If

    If Not wfRS Is Nothing Then
            wfRS = Nothing
    End If
    
'********** 2005/05/17 土田 削除 **********↓
'    If Not vk1RS Is Nothing Then
'            vk1RS = Nothing
'    End If
'********** 2005/05/17 土田 削除 **********↑

'********** 2005/05/17 土田 追加 **********↓
    If Not hyoujiSiyouRS Is Nothing Then
            hyoujiSiyouRS = Nothing
    End If

    If Not ippanKihonRS Is Nothing Then
            ippanKihonRS = Nothing
    End If

'登録時利用したのでここで開放する(作成自体はFncTenTourokuで行っている)
    If Not pTChkJigyouRS Is Nothing Then
       pTChkJigyouRS.Close
            pTChkJigyouRS = Nothing
    End If

'登録時利用したのでここで開放する(作成はFncTenTourokuで行っている)
    If Not pChkvk1JANCodeRS Is Nothing Then
        pChkvk1JANCodeRS.Close
            pChkvk1JANCodeRS = Nothing
    End If
    
    If Not centerCDRS Is Nothing Then
        centerCDRS.Close
            centerCDRS = Nothing
    End If
    
'********** 2005/05/17 土田 追加 **********↑
    
    'If Not GetObjectContext Is Nothing Then
        'GetObjectContext.SetAbort
        'End If

    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
            'Err.Number, Err.Source, Err.Description
End Sub
'********** 2005/05/16 土田 仮作成 **********↑

'******************************************************************'
'*【プロシージャ名】
'*!    subIns_SetAllTablesMosi
'*【処理概要】
'*?    INSERT文作成（商品選定WF(戻し価格)）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*【戻り値】
'*    < Name Type Comment >
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2009/06/08; Supattra T(Tostem Thai); 新規作成;
'*M    2011/02/16; 営業システム/和田; センターコード追加;
'*M    2015/09/15; 流通システムG/和田; TCシステム構築;
'*< Comment End >**************************************************'
Private Sub subIns_SetAllTablesMosi(ByVal arryData As Object)
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    
    Dim strMethodCallInfo As String
    Dim lngCount As Long
    Dim aryJanCd() As String
 
    Dim strNVHFilter As String
    Dim rsNVH As ADODB.Recordset

    Dim lstrKikakuSyurui As String
    lstrKikakuSyurui = pstrKikakuSyurui
    
    Dim lstrKikakuNenrenban As String
    lstrKikakuNenrenban = pstrKikakuNenrenban
    
    Dim lstrKikakuEdaban As String
    lstrKikakuEdaban = pstrKikakuEdaban

    Dim lstrTantousyaId As String
    lstrTantousyaId = pstrTantousyaId

'日付を始めに取得
    Dim lstrNow As Object
    lstrNow = Format(Now(), "yyyy/mm/dd hh:nn:ss")

    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi()"

    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    '---------------
    '   内部処理
    '---------------
    On Error GoTo DefectErrorHandler

    'RS作成用SQLの条件部分（企画種類コード、年連番、枝番の指定）
    Dim strWhereSQL As String
    strWhereSQL = "WHERE kikaku_syurui_cd = '" & lstrKikakuSyurui & "' AND nen_ren_no ='" & lstrKikakuNenrenban & "' AND eda_no ='" & lstrKikakuEdaban & "' "

     'RS作成用SQLの条件部分（変更日付の範囲指定）
    Dim strHenkouSQL As String
    If pCsvHenkouDate <> "" And pCsvhenkouDateKeizoku <> "" Then
        strHenkouSQL = " AND henkou_date BETWEEN '" & pCsvHenkouDate & "' AND '" & pCsvhenkouDateKeizoku & "' "
    Else
        strHenkouSQL = " AND henkou_date BETWEEN '" & pstrMinDate & "' AND '" & pstrMaxDate & "' "
    End If

    'RS作成用SQL
    Dim strSQL  As String

    '売価変更情報RS作成
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(売価変更情報RS作成)"
    strSQL = "SELECT "
    strSQL = strSQL & " jan_cd"
    strSQL = strSQL & " ,ten_cd"
    strSQL = strSQL & " ,kikaku_syurui_cd"
    strSQL = strSQL & " ,nen_ren_no"
    strSQL = strSQL & " ,eda_no"
    strSQL = strSQL & " ,henkou_date"
    strSQL = strSQL & " ,baika"
    strSQL = strSQL & " ,baika_henkou_riyuu_cd"
    strSQL = strSQL & " ,touroku_kbn"
    strSQL = strSQL & " ,touroku_user_id"
    strSQL = strSQL & " ,kousin_user_id"
    strSQL = strSQL & " ,touroku_ymd_hms"
    strSQL = strSQL & " ,kousin_ymd_hms"
    strSQL = strSQL & " FROM " & TBL_BAI_HEN & " with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") " & strHenkouSQL

    Dim baikaRS As ADODB.Recordset
        baikaRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    baikaRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    'baikaRS.Fields("ten_cd").Properties("Optimize") = True 'オプティマイザ
    
    '仕入変更情報RS作成
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(仕入変更情報RS作成)"
    strSQL = "SELECT"
    strSQL = strSQL & " jan_cd"
    strSQL = strSQL & " ,ten_cd"
    strSQL = strSQL & " ,kikaku_syurui_cd"
    strSQL = strSQL & " ,nen_ren_no"
    strSQL = strSQL & " ,eda_no"
    strSQL = strSQL & " ,henkou_date"
    strSQL = strSQL & " ,torihikisaki_cd"
    strSQL = strSQL & " ,hattyuu_tani"
    strSQL = strSQL & " ,genka"
    strSQL = strSQL & " ,touroku_kbn"
    strSQL = strSQL & " ,touroku_user_id"
    strSQL = strSQL & " ,kousin_user_id"
    strSQL = strSQL & " ,touroku_ymd_hms"
    strSQL = strSQL & " ,kousin_ymd_hms"
    strSQL = strSQL & " FROM " & TBL_SII_HEN & " with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") " & strHenkouSQL
    
    Dim siireRS As ADODB.Recordset
        siireRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    siireRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    'siireRS.Fields("ten_cd").Properties("Optimize") = True 'オプティマイザ
    
    'センターコード変更情報RS作成
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(センターコード変更情報RS作成)"
    strSQL = "SELECT"
    strSQL = strSQL & " jan_cd"
    strSQL = strSQL & " ,ten_cd"
    strSQL = strSQL & " ,kikaku_syurui_cd"
    strSQL = strSQL & " ,nen_ren_no"
    strSQL = strSQL & " ,eda_no"
    strSQL = strSQL & " ,henkou_date"
    strSQL = strSQL & " ,center_kbn"
    strSQL = strSQL & " ,center_cd"
    strSQL = strSQL & " ,touroku_kbn"
    strSQL = strSQL & " ,touroku_user_id"
    strSQL = strSQL & " ,kousin_user_id"
    strSQL = strSQL & " ,touroku_ymd_hms"
    strSQL = strSQL & " ,kousin_ymd_hms"
    strSQL = strSQL & " FROM " & TBL_CEN_HEN & " with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") " & strHenkouSQL
    
    Dim centerRS As ADODB.Recordset
        centerRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    centerRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    'centerRS.Fields("ten_cd").Properties("Optimize") = True 'オプティマイザ
    
    'WF管理項目RS作成
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(WF管理項目RS作成)"
    strSQL = "SELECT"
    strSQL = strSQL & " jan_cd"
    strSQL = strSQL & " ,kikaku_syurui_cd"
    strSQL = strSQL & " ,nen_ren_no"
    strSQL = strSQL & " ,eda_no"
    strSQL = strSQL & " ,jan_umu_kbn"
    strSQL = strSQL & " ,touroku_kbn"
    strSQL = strSQL & " ,jyogai_kbn"
    strSQL = strSQL & " ,touroku_tantousya"
    strSQL = strSQL & " ,keikakusuu"
    strSQL = strSQL & " ,kikaku_no"
    strSQL = strSQL & " ,kikaku_kahi_kbn"
    strSQL = strSQL & " ,kikaku_syouninsya_comment"
    strSQL = strSQL & " ,hinsyu_kahi_kbn"
    strSQL = strSQL & " ,hinsyu_syouninsya_comment"
    strSQL = strSQL & " ,naibu_cd"
    strSQL = strSQL & " ,hyouji_siyou_touroku_kbn"
    strSQL = strSQL & " ,sekisai_size_touroku_kbn"
    strSQL = strSQL & " ,pop_touroku_kbn"
    strSQL = strSQL & " ,syanai_kihon_touroku_kbn"
    strSQL = strSQL & " ,kowake_touroku_kbn"
    strSQL = strSQL & " ,jigyoubetu_touroku_kbn"
    strSQL = strSQL & " ,hanbai_touroku_kbn"
    strSQL = strSQL & " ,siire_touroku_kbn"
    strSQL = strSQL & " ,hinsyu_taikei_touroku_kbn"
    strSQL = strSQL & " ,uriba_touroku_kbn"
    strSQL = strSQL & " ,baika_henkou_touroku_kbn"
    strSQL = strSQL & " ,siire_henkou_touroku_kbn"
    strSQL = strSQL & " ,center_cd_henkou_touroku_kbn"
    strSQL = strSQL & " ,kowake_henkou_touroku_kbn"
    strSQL = strSQL & " ,hanki_henkou_touroku_kbn"
    strSQL = strSQL & " ,touroku_user_id"
    strSQL = strSQL & " ,kousin_user_id"
    strSQL = strSQL & " ,touroku_ymd_hms"
    strSQL = strSQL & " ,kousin_ymd_hms"
    strSQL = strSQL & " FROM t_syousen_wf_kanri with (UPDLOCK) " & strWhereSQL
    
    Dim wfRS As ADODB.Recordset
        wfRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    wfRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    
    'VSWF表示仕様テーブルRS
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(VSWF表示仕様テーブルRS生成)"
    strSQL = "SELECT"
    strSQL = strSQL & "  jan_cd"
    strSQL = strSQL & " ,kikaku_syurui_cd"
    strSQL = strSQL & " ,nen_ren_no"
    strSQL = strSQL & " ,eda_no"
    strSQL = strSQL & " ,hyoujiyou_syouhin_mei_kana"
    strSQL = strSQL & " ,hyoujiyou_syouhin_mei_kanji"
    strSQL = strSQL & " ,receipt_mei_kana"
    strSQL = strSQL & " ,receipt_mei_kanji"
    strSQL = strSQL & " ,syouhin_mei_kanji"
    strSQL = strSQL & " ,maker_mei_kanji"
    strSQL = strSQL & " ,kikaku_size_kanji"
    strSQL = strSQL & " ,sina_no_kata_no_kanji"
    strSQL = strSQL & " ,iro_gara_kanji"
    strSQL = strSQL & " ,taste_kanji"
    strSQL = strSQL & " ,seisanti_mei_kanji"
    strSQL = strSQL & " ,tokutyou"
    strSQL = strSQL & " ,syoumi_siyoukigen_um_kbn"
    strSQL = strSQL & " ,eco_syouhin_kbn"
    strSQL = strSQL & " ,green_tekigou_syouhin_kbn"
    strSQL = strSQL & " ,gpn_db_keisai_syouhin_kbn"
    strSQL = strSQL & " ,touroku_user_id"
    strSQL = strSQL & " ,kousin_user_id"
    strSQL = strSQL & " ,touroku_ymd_hms"
    strSQL = strSQL & " ,kousin_ymd_hms"
    strSQL = strSQL & " FROM t_syousen_hyouji_siyou WITH(UPDLOCK) " & strWhereSQL
    
    Dim hyoujiSiyouRS As ADODB.Recordset
        hyoujiSiyouRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    hyoujiSiyouRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
         
    'VMST一般基本テーブルRS
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(VMST一般基本テーブルRS生成)"
    strSQL = _
    "SELECT v_key_1.jan_cd, v_ippan_kihon.*" & _
    "  FROM v_ippan_kihon WITH(READCOMMITTED)" & _
    " INNER JOIN v_key_1 WITH(READCOMMITTED) ON v_ippan_kihon.key_cd = v_key_1.key_cd" & _
    " AND v_key_1.jan_cd IN(" & pCsvJanCodeList & ") "
    '↑janコードでの絞込み追加
    
    Dim ippanKihonRS As ADODB.Recordset
        ippanKihonRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    ippanKihonRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
        
    'センター区分を元にして、センターコードを取得する為のRS
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(センター区分を元にして、センターコードを取得する為のRS生成)"
    strSQL = _
    "SELECT jigyousyo_cd, center_cd " & _
    "  FROM v_jigyousyo WITH(READCOMMITTED)"
    
    Dim centerCDRS As ADODB.Recordset
        centerCDRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    centerCDRS.Fields("jigyousyo_cd").Properties("Optimize") = True 'オプティマイザ

    'WF更新用フラグ(WF更新時の効率化の為Stringで"0"又は"1"を持つ)
    Dim strBaikaAdded As String     '売価変更RSが更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    Dim strSiireAdded As String     '仕入変更RSが更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    Dim strCenterAdded As String    'センターコード変更RSが更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    
    ' ************** 1レコードずつRSを更新 *******************
    
    Dim strBaikaFlg As String       '売価が更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    Dim strSiireFlg As String       '仕入が更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    Dim strCenterFlg As String    'センターコードが更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(1レコードずつRSを更新 開始)"
    Dim iRecord As Long
    
    For iRecord = LBound(arryData) To UBound(arryData)
    
    
        'WF更新用フラグの初期化
        strBaikaAdded = "0"
        strSiireAdded = "0"
        strCenterAdded = "0"
        
        '売価、仕入情報、センターコードが更新(新規又は更新)されたかのフラグ
        strBaikaFlg = "0"
        strSiireFlg = "0"
        strCenterFlg = "0"
        
         '2009.06.25 Mod by Supattra T.(Tostem Thai)Start.

        '売価、仕入、センターコードRSをJanCDと変更日でフィルタリング
        Dim strFilter As String
        
'        Dim arrypTTenCode As Object
'        Dim arrypTJanCode As Object
'        Dim arrypTJanCodeNew As Object
'        Dim Count As Integer
'        Dim inti As Integer
'        Dim CountTenCd As Integer
        
        If pCsvHenkouDate <> "" And pCsvhenkouDateKeizoku <> "" Then
            With pTChkkikakuRS
                 If Not .BOF Then
                    .MoveFirst
                    .Filter = "jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'"
                 End If
                 
            End With
        End If
        
'
'        If pTJanCode = "" Then
'             paryCsvData(iRecord)(enmTenColNo.HenkouDate) = pCsvHenkouDate
'             strFilter = "jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "' AND henkou_date = '" & pCsvHenkouDate & "'"
'        End If
'
'        arrypTJanCode = Split(pTJanCode, ",")
'
'        For Count = LBound(arrypTJanCode) To UBound(arrypTJanCode)
'            arrypTJanCodeNew = Split(arrypTJanCode(Count), "*")
'             For inti = 0 To UBound(arrypTJanCodeNew)
'                If arrypTJanCodeNew(inti) = arryData(iRecord)(enmTenColNo.janCD) Then
'                    paryCsvData(iRecord)(enmTenColNo.HenkouDate) = pCsvhenkouDateKeizoku
'                    strFilter = "jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "' AND henkou_date = '" & pCsvhenkouDateKeizoku & "'"
'                    Exit For
'                Else
'                    paryCsvData(iRecord)(enmTenColNo.HenkouDate) = pCsvHenkouDate
'                    strFilter = "jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "' AND henkou_date = '" & pCsvHenkouDate & "'"
'                End If
'             Next
'        Next Count
         
        '2009.06.25 Mod by Supattra T.(Tostem Thai)End.

        strFilter = "jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'"
        
        baikaRS.Filter = strFilter
        siireRS.Filter = strFilter
        centerRS.Filter = strFilter
        
        'データ更新対象店コードの取得
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(データ更新対象店コードの取得)"
        
        'このレコードの店コード
        Dim TenCd As String
        TenCd = arryData(iRecord)(enmTenColNo.TenCd)
        
        'データ更新対象店コードの取得に該当事業コードチェック用RSを再利用する
        With pTChkJigyouRS
            'フィルターを解除しておく
            .Filter = adFilterNone

            If TenCd = ALL_TEN_CD Or TenCd = ALL_JIG_CD Or TenCd = HON_TEN_CD Then
                '全店、全事業、本部の場合なにもしない
            ElseIf Len(TenCd) = 2 Then
                '各事業の場合
                '該当事業コードチェック用RSを事業コードでフィルタリングする

                strNVHFilter = ""
                strNVHFilter = strNVHFilter & "jigyou_cd = '" & TenCd & "'"
                If arryData(iRecord)(enmTenColNo.TenCd) = "01" Then
                        rsNVH = GetNVHInfo(arryData(iRecord)(enmTenColNo.janCD))
                    Do While Not rsNVH.EOF
                        strNVHFilter = strNVHFilter & " and jigyousyo_cd <> '" & rsNVH.Fields("ten_cd").Value & "'"
                        rsNVH.MoveNext
                    Loop
                    rsNVH.Close
                        rsNVH = Nothing
                End If
                
                .Filter = strNVHFilter

            ElseIf Len(TenCd) = 4 Then

                '各店の場合、該当事業コードチェック用RSを店コードでフィルタリングする
                .Filter = "jigyousyo_cd = '" & TenCd & "'"
            End If
        End With
        
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(データ更新対象店コードのループ 開始) "
' ************** データ更新対象店コードのループ *******************

        Dim strTenCd As String
        Dim i As Integer
        Dim strCaseHenkouDate As String
        
        Do Until pTChkJigyouRS.EOF
           
'本部("0000")の場合は店コードに"0000"を直接入れる(レコードセットからは取得しない)
            If (TenCd = HON_TEN_CD) Then
                strTenCd = HON_TEN_CD
            Else
                strTenCd = pTChkJigyouRS.Fields("jigyousyo_cd")
            
            End If
                        
           '売価、売価変更理由のいずれかがある場合、売価変更情報RSに追加か更新
            If arryData(iRecord)(enmTenColNo.Baika) <> "" Or arryData(iRecord)(enmTenColNo.BaikaRiyuu) <> "" Then
                '売価変更情報RSに追加か更新
                With baikaRS
                    If Not .BOF Then
                        .MoveFirst
                        .Find ("ten_cd = '" & Trim(strTenCd) & "'")
                    End If
                    
                    If .EOF Then
                        '新規追加
                        .AddNew
                        .Fields("kikaku_syurui_cd").Value = lstrKikakuSyurui
                        .Fields("nen_ren_no").Value = lstrKikakuNenrenban
                        .Fields("eda_no").Value = lstrKikakuEdaban
                        .Fields("jan_cd").Value = arryData(iRecord)(enmTenColNo.janCD)
                        .Fields("ten_cd").Value = strTenCd
                        '2009.06.25 Mod by Supattra T.(Tostem Thai)Start.
'                        .Fields("henkou_date").Value = pCsvHenkouDate
                        
                        
                        If pCsvHenkouDate <> "" And pCsvhenkouDateKeizoku <> "" Then
                            With pTChkkikakuRS
                                 If Not .BOF Then
                                    .MoveFirst
                                    '.Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
                                    .Find ("ten_cd = '" & strTenCd & "'")
                                 End If
                                 If .EOF Then
                                      strCaseHenkouDate = pCsvHenkouDate
                                 Else
                                      strCaseHenkouDate = pCsvhenkouDateKeizoku
                                 End If
                            End With
                        Else
                            strCaseHenkouDate = arryData(iRecord)(enmTenColNo.HenkouDate)
                        End If
                            
                                        
'                        If pTJanCode = "" Then
'                            strCaseHenkouDate = pCsvHenkouDate
'                        End If
'
'                        For Count = LBound(arrypTJanCode) To UBound(arrypTJanCode)
'                            arrypTJanCodeNew = Split(arrypTJanCode(Count), "*")
'                             For inti = 0 To UBound(arrypTJanCodeNew)
'                                If arrypTJanCodeNew(inti) = arryData(iRecord)(enmTenColNo.janCD) Then
'                                    If strTenCd = arrypTJanCodeNew(inti + 1) Then
'                                            strCaseHenkouDate = pCsvhenkouDateKeizoku
'                                        Exit For
'                                    End If
'                                Else
'                                   strCaseHenkouDate = pCsvHenkouDate
'                                End If
'                             Next inti
'                        Next Count

                         .Fields("henkou_date").Value = strCaseHenkouDate
                        '2009.06.25 Mod by Supattra T.(Tostem Thai)End.
                        
                        
                        '売価
                        '********** 2008/05/27 ラン峰（大連） 追加 **********↓
                        '商品選定WF(変更除外店舗)テーブルの｢売価除外区分｣フィールドは｢除外｣になると
                        'CSVから取得した｢売価｣が登録・更新されない(NULLにセット)。
                                                                                            
                        '商品選定WF(変更除外店舗)テーブルの｢売価除外区分｣フィールドは｢除外｣ではなければ
                        '商品選定DB（売価変更情報）テーブルへCSVから取得した｢売価｣が登録・更新される。
                       
                        If (strTenCd <> HON_TEN_CD) And _
                            (pTChkJigyouRS.Fields("baika_jyogai_kbn").Value = "1") Then
                        '商品選定WF(変更除外店舗)テーブルの｢売価除外区分｣フィールドは｢除外｣になると
                        'NULLにセット
                            .Fields("baika").Value = dbNull.value
                        Else
                        
                            If arryData(iRecord)(enmTenColNo.Baika) <> "" Then
                                .Fields("baika").Value = CLng(arryData(iRecord)(enmTenColNo.Baika))
                                strBaikaFlg = "1"
                            Else
                                .Fields("baika").Value = dbNull.value
                            End If
                        End If
                        
                        
                        '売価変更理由コード
                        If arryData(iRecord)(enmTenColNo.BaikaRiyuu) <> "" Then
                            .Fields("baika_henkou_riyuu_cd").Value = arryData(iRecord)(enmTenColNo.BaikaRiyuu)
                        Else
                            .Fields("baika_henkou_riyuu_cd").Value = dbNull.value
                        End If
'売価先付変更情報が基幹DBに既存の場合は登録区分｢2｣をセットする(承認されると先付変更情報を上書するデータとなる)
                        If InStr(1, arrTourokuKbn2(iRecord, conColBaika), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
                        .Fields("touroku_user_id").Value = lstrTantousyaId
                        .Fields("touroku_ymd_hms").Value = lstrNow
                        
                    Else
'                        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(更新処理)"
                        ' 更新処理
                        '売価
                        '********** 2008/05/27 ラン峰（大連） 追加 **********↓
                        '商品選定WF(変更除外店舗)テーブルの｢売価除外区分｣フィールドは｢除外｣になると
                        'CSVから取得した｢売価｣が登録・更新されない(NULLにセット)。
                                                                                            
                        '商品選定WF(変更除外店舗)テーブルの｢売価除外区分｣フィールドは｢除外｣ではなければ
                        '商品選定DB（売価変更情報）テーブルへCSVから取得した｢売価｣が登録・更新される。
                       
                       
                        If (strTenCd <> HON_TEN_CD) And _
                            (pTChkJigyouRS.Fields("baika_jyogai_kbn").Value = "1") Then
                       
                        '商品選定WF(変更除外店舗)テーブルの｢売価除外区分｣フィールドは｢除外｣になると
                        'NULLにセット
                            .Fields("baika").Value = dbNull.value
                        Else
                            If arryData(iRecord)(enmTenColNo.Baika) <> "" Then
                                .Fields("baika").Value = CLng(arryData(iRecord)(enmTenColNo.Baika))
                                strBaikaFlg = "1"
                            End If
                        End If
                       
                        '売価変更理由コード
                        If arryData(iRecord)(enmTenColNo.BaikaRiyuu) <> "" Then
                            .Fields("baika_henkou_riyuu_cd").Value = arryData(iRecord)(enmTenColNo.BaikaRiyuu)
                        End If
'売価先付変更情報が基幹DBに既存の場合は登録区分｢2｣をセットする(承認されると先付変更情報を上書するデータとなる)
                        If InStr(1, arrTourokuKbn2(iRecord, conColBaika), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
                        .Fields("kousin_user_id").Value = lstrTantousyaId
                        .Fields("kousin_ymd_hms").Value = lstrNow
                    End If
                    .Update
                   
                End With
                strBaikaAdded = "1"

            End If
            
            
            '原価、取引先、発注単位のどれかがあるかチェックする。
            If arryData(iRecord)(enmTenColNo.Genka) <> "" _
                Or arryData(iRecord)(enmTenColNo.Torihikisaki) <> "" _
                Or arryData(iRecord)(enmTenColNo.Hacchuutani) <> "" Then
                
                '仕入変更情報RS更新
                With siireRS
'                   If Not .EOF Then
                    If Not .BOF Then
                        .MoveFirst
                        .Find ("ten_cd = '" & strTenCd & "'")
                    End If
                    
                    If .EOF Then
                         '新規追加
                        .AddNew
                        .Fields("kikaku_syurui_cd").Value = lstrKikakuSyurui
                        .Fields("nen_ren_no").Value = lstrKikakuNenrenban
                        .Fields("eda_no").Value = lstrKikakuEdaban
                        .Fields("jan_cd").Value = arryData(iRecord)(enmTenColNo.janCD)
                        .Fields("ten_cd").Value = strTenCd
                        
                        If pCsvHenkouDate <> "" And pCsvhenkouDateKeizoku <> "" Then
                            With pTChkkikakuRS
                                 If Not .BOF Then
                                    .MoveFirst
                                    '.Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
                                    .Find ("ten_cd = '" & strTenCd & "'")
                                 End If
                                 If .EOF Then
                                      strCaseHenkouDate = pCsvHenkouDate
                                 Else
                                      strCaseHenkouDate = pCsvhenkouDateKeizoku
                                 End If
                            End With
                        Else
                            strCaseHenkouDate = arryData(iRecord)(enmTenColNo.HenkouDate)
                        End If
                        
                        .Fields("henkou_date").Value = strCaseHenkouDate
                        
                        '取引先コード
                        '商品選定WF(変更除外店舗)テーブルの｢取引先除外区分｣は｢除外｣ではない場合
                        '取引先コードは登録される。
                       
                        If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("torihikisaki_jyogai_kbn").Value = "1" Then
                       
                            .Fields("torihikisaki_cd").Value = dbNull.value
                            
                        Else
                            If arryData(iRecord)(enmTenColNo.Torihikisaki) <> "" Then
                                .Fields("torihikisaki_cd").Value = arryData(iRecord)(enmTenColNo.Torihikisaki)
                                strSiireFlg = "1"
                            Else
                                .Fields("torihikisaki_cd").Value = dbNull.value
                            End If
                        End If
                       
                        '発注単位
                         '商品選定WF(変更除外店舗)テーブルの｢発注単位除外区分｣は｢除外｣ではない場合
                         '発注単位は登録される。
                 
                        If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("hattyuu_tani_jyogai_kbn").Value = "1" Then
                       
                            .Fields("hattyuu_tani").Value = dbNull.value
                            
                        Else
                            If arryData(iRecord)(enmTenColNo.Hacchuutani) <> "" Then
                                .Fields("hattyuu_tani").Value = CLng(arryData(iRecord)(enmTenColNo.Hacchuutani))
                                strSiireFlg = "1"
                            Else
                                .Fields("hattyuu_tani").Value = dbNull.value
                            End If
                         
                        End If
                        
                        '原価
                         '商品選定WF(変更除外店舗)テーブルの｢原価除外区分｣は｢除外｣ではない場合
                         '原価は登録される。
                         
                        If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("genka_jyogai_kbn").Value = "1" Then
                            .Fields("genka").Value = dbNull.value
                        Else
                            If arryData(iRecord)(enmTenColNo.Genka) <> "" Then
                                .Fields("genka").Value = CDbl(arryData(iRecord)(enmTenColNo.Genka))
                                strSiireFlg = "1"
                            Else
                                .Fields("genka").Value = dbNull.value
                            End If
                        End If
                       
'仕入先付変更情報が基幹DBに既存の場合は登録区分｢2｣をセットする(承認されると先付変更情報を上書するデータとなる)
                        If InStr(1, arrTourokuKbn2(iRecord, conColSiire), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
                        .Fields("touroku_user_id").Value = lstrTantousyaId
                        .Fields("touroku_ymd_hms").Value = lstrNow
                    Else
                        '更新処理
                        '取引先コード
                        '商品選定WF(変更除外店舗)テーブルの｢取引先除外区分｣は｢除外｣ではない場合
                        '取引先コードは更新される。
                        If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("torihikisaki_jyogai_kbn").Value = "1" Then
                    
                            .Fields("torihikisaki_cd").Value = dbNull.value
                        Else
                        
                            If arryData(iRecord)(enmTenColNo.Torihikisaki) <> "" Then
                                 .Fields("torihikisaki_cd").Value = arryData(iRecord)(enmTenColNo.Torihikisaki)
                                 strSiireFlg = "1"
                            End If
                        End If
                        
                        '発注単位
                        '商品選定WF(変更除外店舗)テーブルの｢発注単位除外区分｣は｢除外｣ではない場合
                        '発注単位は更新される。
                        If (strTenCd <> HON_TEN_CD) And _
                                pTChkJigyouRS.Fields("hattyuu_tani_jyogai_kbn").Value = "1" Then
                            .Fields("hattyuu_tani").Value = dbNull.value
                        Else
                            If arryData(iRecord)(enmTenColNo.Hacchuutani) <> "" Then
                                .Fields("hattyuu_tani").Value = arryData(iRecord)(enmTenColNo.Hacchuutani)
                                strSiireFlg = "1"
                            End If
                        End If
                        
                        '原価
                        '商品選定WF(変更除外店舗)テーブルの｢原価除外区分｣は｢除外｣ではない場合
                        '原価は更新される。
                         If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("genka_jyogai_kbn").Value = "1" Then
                            .Fields("genka").Value = dbNull.value
                                
                        Else
                            
                            If arryData(iRecord)(enmTenColNo.Genka) <> "" Then
                                .Fields("genka").Value = CDbl(arryData(iRecord)(enmTenColNo.Genka))
                                strSiireFlg = "1"
                            End If
                        End If
                        
'仕入先付変更情報が基幹DBに既存の場合は登録区分｢2｣をセットする(承認されると先付変更情報を上書するデータとなる)
                        If InStr(1, arrTourokuKbn2(iRecord, conColSiire), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
                        .Fields("kousin_user_id").Value = lstrTantousyaId
                        .Fields("kousin_ymd_hms").Value = lstrNow
                   End If
                End With
                strSiireAdded = "1"
            End If
        
            'センター区分があるかチェックする。
            If arryData(iRecord)(enmTenColNo.CenterKbn) <> "" Then
                'センターコード変更情報RS更新
                'センターコードを決める：
                'センター区分が0の場合はNULL、1の場合は'0151'、2の場合は事業所マスタから取得する
'   センター区分が０の場合は""ではなくNULLを入れるよう修正
                Dim centerCD
                Select Case arryData(iRecord)(enmTenColNo.CenterKbn)
                    Case "0":       'センター区分が0の場合、Nullをセット
'                        strCenterCD = ""
                        centerCD = dbNull.value
                    
                    '2015/09/15 wadak7 upd
                    'Case "1":       'センター区分が1の場合、0151(児玉センター)をセット
                    Case "1", "3":      'センター区分が1,3の場合
                        'centerCD = CENTERCD_KODAMA
                        '登録票のセンターコードをセット
                        centerCD = arryData(iRecord)(enmModColNo.centerCD)
                    
                    Case "2":       'センター区分が2の場合、店コードを元に事業所マスタから取得したセンターコードをセット
                        With centerCDRS
                            If Not .BOF Then
                                .MoveFirst
                                .Find ("jigyousyo_cd = '" & strTenCd & "'")
                            End If
                        
                            'データが無かったらNullをセットする。
                            If .EOF Then
                                centerCD = dbNull.value
                            Else
                                centerCD = centerCDRS.Fields("center_cd")
                            End If
                        End With
                        
                    Case Else:       'センター区分が0,1,2以外の場合、Nullをセット
                        centerCD = dbNull.value
                        
                End Select
                
                With centerRS
                    If Not .BOF Then
                        .MoveFirst
                        .Find ("ten_cd = '" & strTenCd & "'")
                    End If
                    
                    If .EOF Then
                         '新規追加
                        .AddNew
                        .Fields("kikaku_syurui_cd").Value = lstrKikakuSyurui
                        .Fields("nen_ren_no").Value = lstrKikakuNenrenban
                        .Fields("eda_no").Value = lstrKikakuEdaban
                        .Fields("jan_cd").Value = arryData(iRecord)(enmTenColNo.janCD)
                        .Fields("ten_cd").Value = strTenCd
                        .Fields("henkou_date").Value = strCaseHenkouDate
                        If strCaseHenkouDate = "" Then
                            .Fields("henkou_date").Value = arryData(iRecord)(enmTenColNo.HenkouDate)
                        End If
               
                        If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("center_jyogai_kbn").Value = "1" Then
                        '商品選定WF(変更除外店舗)テーブルの｢センター除外区分｣は｢除外｣であれば、NULLにセット
                            .Fields("center_kbn").Value = dbNull.value
                        Else
                            .Fields("center_kbn").Value = arryData(iRecord)(enmTenColNo.CenterKbn)
                            'センターコード：センター区分が0の場合はNULL、1の場合は'0151'、2の場合は事業所マスタから取得する
                            .Fields("center_cd").Value = centerCD
                            strCenterFlg = "1"
                        End If
                       
'センター先付変更情報が基幹DBに既存の場合は登録区分｢2｣をセットする(承認されると先付変更情報を上書するデータとなる)
                        If InStr(1, arrTourokuKbn2(iRecord, conColCenter), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If

                        .Fields("touroku_user_id").Value = lstrTantousyaId
                        .Fields("touroku_ymd_hms").Value = lstrNow
                    Else
                        ' 更新処理
                        'センター区分
                        If (strTenCd <> HON_TEN_CD) And _
                                pTChkJigyouRS.Fields("center_jyogai_kbn").Value = "1" Then
                            '商品選定WF(変更除外店舗)テーブルの｢センター除外区分｣は｢除外｣であれば、NULLにセット
                                .Fields("center_kbn").Value = dbNull.value
                        Else
                            If arryData(iRecord)(enmTenColNo.CenterKbn) <> "" Then
                                .Fields("center_kbn").Value = arryData(iRecord)(enmTenColNo.CenterKbn)
                                'センターコード：センター区分が0の場合はNULL、1の場合は'0151'、2の場合は事業所マスタから取得する
                                .Fields("center_cd").Value = centerCD
                                strCenterFlg = "1"
                            End If
                        End If
                        
'センター先付変更情報が基幹DBに既存の場合は登録区分｢2｣をセットする(承認されると先付変更情報を上書するデータとなる)
                        If InStr(1, arrTourokuKbn2(iRecord, conColCenter), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If

                        .Fields("kousin_user_id").Value = lstrTantousyaId
                        .Fields("kousin_ymd_hms").Value = lstrNow
                    End If
                End With
                strCenterAdded = "1"
           End If
           
'本部("0000")の場合は一回でDoループを抜ける
            If (strTenCd = HON_TEN_CD) Then Exit Do
            pTChkJigyouRS.MoveNext
            
        Loop
            
        '売価、仕入、センターコードRSのいずれかで追加処理があった場合
        If (strBaikaAdded = "1" Or strCenterAdded = "1" Or strSiireAdded = "1") Then


            'JANコードと紐付く主キーを取得する
            Dim strKey_cd

'            With vk1RS
            With pChkvk1JANCodeRS
                If Not .BOF Then
                    .MoveFirst
                    .Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
                    strKey_cd = .Fields("key_cd")
                End If
            End With

            
            'WF管理項目を最新情報に更新する
            With wfRS
                If Not .BOF Then
                    .MoveFirst
                    .Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
                End If
                'WF管理項目に無い場合は、新規追加
                If .EOF Then
                    .AddNew
                    .Fields("kikaku_syurui_cd").Value = lstrKikakuSyurui
                    .Fields("nen_ren_no").Value = lstrKikakuNenrenban
                    .Fields("eda_no").Value = lstrKikakuEdaban
                    .Fields("jan_cd").Value = arryData(iRecord)(enmTenColNo.janCD)
                    .Fields("jan_umu_kbn").Value = 1
                    .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                    .Fields("jyogai_kbn").Value = 0
                    .Fields("touroku_tantousya").Value = lstrTantousyaId
                    .Fields("keikakusuu").Value = 0
                    .Fields("kikaku_no").Value = lstrKikakuSyurui & lstrKikakuNenrenban & "-" & lstrKikakuEdaban
                    .Fields("kikaku_kahi_kbn").Value = 0
                    .Fields("kikaku_syouninsya_comment").Value = dbNull.value
                    .Fields("hinsyu_kahi_kbn").Value = 0
                    .Fields("hinsyu_syouninsya_comment").Value = dbNull.value
                    .Fields("naibu_cd").Value = strKey_cd
                    .Fields("hyouji_siyou_touroku_kbn").Value = 1
                    .Fields("sekisai_size_touroku_kbn").Value = 0
                    .Fields("pop_touroku_kbn").Value = 0
                    .Fields("syanai_kihon_touroku_kbn").Value = 0
                    .Fields("kowake_touroku_kbn").Value = 0
                    .Fields("jigyoubetu_touroku_kbn").Value = 0
                    .Fields("hanbai_touroku_kbn").Value = 0
                    .Fields("siire_touroku_kbn").Value = 0
                    .Fields("hinsyu_taikei_touroku_kbn").Value = 0
                    .Fields("uriba_touroku_kbn").Value = 0
                    If strBaikaFlg = "1" Then
                        .Fields("baika_henkou_touroku_kbn").Value = strBaikaAdded
                    Else
                        .Fields("baika_henkou_touroku_kbn").Value = 0
                    End If

                    If strSiireFlg = "1" Then
                        .Fields("siire_henkou_touroku_kbn").Value = strSiireAdded
                    Else
                        .Fields("siire_henkou_touroku_kbn").Value = 0
                    End If
                    
                    If strCenterFlg = "1" Then
                        .Fields("center_cd_henkou_touroku_kbn").Value = strCenterAdded
                    Else
                        .Fields("center_cd_henkou_touroku_kbn").Value = 0
                    End If
                    
                    .Fields("kowake_henkou_touroku_kbn").Value = 0
                    .Fields("hanki_henkou_touroku_kbn").Value = 0
                    .Fields("touroku_user_id").Value = lstrTantousyaId
                    .Fields("kousin_user_id").Value = dbNull.value
                    .Fields("touroku_ymd_hms").Value = lstrNow
                    .Fields("kousin_ymd_hms").Value = dbNull.value
        
                'WF管理項目にある場合は、更新
                Else
                    .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                    .Fields("touroku_tantousya").Value = lstrTantousyaId
                    .Fields("kikaku_no").Value = lstrKikakuSyurui & lstrKikakuNenrenban & "-" & lstrKikakuEdaban
                    .Fields("naibu_cd").Value = strKey_cd
                    .Fields("hyouji_siyou_touroku_kbn").Value = 1

                    If strBaikaFlg = "1" Then
                        If strBaikaAdded = "1" Then
                            '売価変更情報登録区分
                            .Fields("baika_henkou_touroku_kbn").Value = strBaikaAdded
                        End If
                        
                    End If
                    If strSiireFlg = "1" Then
                        If strSiireAdded = "1" Then
                            '仕入変更情報登録区分
                            .Fields("siire_henkou_touroku_kbn").Value = strSiireAdded
                        End If
                    End If
                    If strCenterFlg = "1" Then
                        If strCenterAdded = "1" Then
                            'ｾﾝﾀｰｺｰﾄﾞ変更情報登録区分
                            .Fields("center_cd_henkou_touroku_kbn").Value = strCenterAdded
                        End If
                    End If

                    .Fields("kousin_user_id").Value = lstrTantousyaId
                    .Fields("kousin_ymd_hms").Value = lstrNow
                End If
                .Update
            End With
        End If
        
        '各RSのフィルター解除
        baikaRS.Filter = adFilterNone
        siireRS.Filter = adFilterNone
        centerRS.Filter = adFilterNone
        pTChkkikakuRS.Filter = adFilterNone
        
'VMST一般基本テーブルより商品情報を取得し、VSWF表示仕様テーブルにレコードを追加する
        With hyoujiSiyouRS
            If Not .BOF Then
                .MoveFirst
                .Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
            End If
            
            'VSWF表示仕様テーブルにレコードが存在しない場合はVMST一般基本テーブルより商品情報を取得しVSWF表示仕様テーブルにレコードを追加する
            If .EOF Then
                With ippanKihonRS
                    If Not .BOF Then
                        .MoveFirst
                        .Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
                    End If
                    If .EOF Then
                        GoTo DefectErrorHandler
                    Else
                        'VSWF表示仕様テーブルRSを更新する
                        With hyoujiSiyouRS
                            .AddNew
                            .Fields("kikaku_syurui_cd") = lstrKikakuSyurui
                            .Fields("nen_ren_no") = lstrKikakuNenrenban
                            .Fields("eda_no") = lstrKikakuEdaban
                            .Fields("jan_cd") = arryData(iRecord)(enmTenColNo.janCD)
                            .Fields("hyoujiyou_syouhin_mei_kana") = ippanKihonRS("hyoujiyou_syouhin_mei_kana")
                            .Fields("hyoujiyou_syouhin_mei_kanji") = ippanKihonRS("hyoujiyou_syouhin_mei_kanji")
                            .Fields("receipt_mei_kana") = ippanKihonRS("receipt_mei_kana")
                            .Fields("receipt_mei_kanji") = ippanKihonRS("receipt_mei_kanji")
                            .Fields("syouhin_mei_kanji") = ippanKihonRS("syouhin_mei_kanji")
                            .Fields("maker_mei_kanji") = ippanKihonRS("maker_mei_kanji")
                            .Fields("kikaku_size_kanji") = ippanKihonRS("kikaku_size_kanji")
                            .Fields("sina_no_kata_no_kanji") = ippanKihonRS("sina_no_kata_no_kanji")
                            .Fields("iro_gara_kanji") = ippanKihonRS("iro_gara_kanji")
                            .Fields("taste_kanji") = ippanKihonRS("taste_kanji")
                            .Fields("seisanti_mei_kanji") = ippanKihonRS("seisanti_mei_kanji")
                            .Fields("tokutyou") = ippanKihonRS("tokutyou")
                            .Fields("syoumi_siyoukigen_um_kbn") = ippanKihonRS("syoumi_siyoukigen_um_kbn")
                            .Fields("eco_syouhin_kbn") = ippanKihonRS("eco_syouhin_kbn")
                            .Fields("green_tekigou_syouhin_kbn") = ippanKihonRS("green_tekigou_syouhin_kbn")
                            .Fields("gpn_db_keisai_syouhin_kbn") = ippanKihonRS("gpn_db_keisai_syouhin_kbn")
                            .Fields("touroku_user_id") = lstrTantousyaId
                            .Fields("touroku_ymd_hms") = lstrNow
                            .Update
                        End With
                    End If
                End With
            End If
        End With
    Next iRecord

    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(RSのバッチ更新0)"
    'RSのバッチ更新
    Call objCDBAccess.UpdateDCRecordset("VSWF", baikaRS)
    Call objCDBAccess.UpdateDCRecordset("VSWF", siireRS)
    Call objCDBAccess.UpdateDCRecordset("VSWF", centerRS)
    Call objCDBAccess.UpdateDCRecordset("VSWF", wfRS)
    Call objCDBAccess.UpdateDCRecordset("VSWF", hyoujiSiyouRS)
    
    'GetObjectContext.SetComplete
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(オブジェクト開放)"
    
    
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If

    If Not baikaRS Is Nothing Then
            baikaRS = Nothing
    End If

    If Not siireRS Is Nothing Then
            siireRS = Nothing
    End If

    If Not centerRS Is Nothing Then
            centerRS = Nothing
    End If

    If Not wfRS Is Nothing Then
            wfRS = Nothing
    End If

    If Not hyoujiSiyouRS Is Nothing Then
            hyoujiSiyouRS = Nothing
    End If

    If Not ippanKihonRS Is Nothing Then
            ippanKihonRS = Nothing
    End If
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(登録時利用したのでここで開放する)"
'登録時利用したのでここで開放する(作成はFncTenTourokuで行っている)
    If Not pTChkJigyouRS Is Nothing Then
        pTChkJigyouRS.Close
            pTChkJigyouRS = Nothing
    End If
    If Not pTChkkikakuRS Is Nothing Then
        pTChkkikakuRS.Close
            pTChkkikakuRS = Nothing
    End If
    
'登録時利用したのでここで開放する(作成はFncTenTourokuで行っている)
    If Not pChkvk1JANCodeRS Is Nothing Then
        pChkvk1JANCodeRS.Close
            pChkvk1JANCodeRS = Nothing
    End If
    
    If Not centerCDRS Is Nothing Then
        centerCDRS.Close
            centerCDRS = Nothing
    End If

    Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If

    If Not baikaRS Is Nothing Then
            baikaRS = Nothing
    End If

    If Not siireRS Is Nothing Then
            siireRS = Nothing
    End If

    If Not centerRS Is Nothing Then
            centerRS = Nothing
    End If

    If Not wfRS Is Nothing Then
            wfRS = Nothing
    End If
    
    If Not hyoujiSiyouRS Is Nothing Then
            hyoujiSiyouRS = Nothing
    End If

    If Not ippanKihonRS Is Nothing Then
            ippanKihonRS = Nothing
    End If

'登録時利用したのでここで開放する(作成自体はFncTenTourokuで行っている)
    If Not pTChkJigyouRS Is Nothing Then
       pTChkJigyouRS.Close
            pTChkJigyouRS = Nothing
    End If
    
    If Not pTChkkikakuRS Is Nothing Then
        pTChkkikakuRS.Close
            pTChkkikakuRS = Nothing
    End If

'登録時利用したのでここで開放する(作成はFncTenTourokuで行っている)
    If Not pChkvk1JANCodeRS Is Nothing Then
        pChkvk1JANCodeRS.Close
            pChkvk1JANCodeRS = Nothing
    End If
    
    If Not centerCDRS Is Nothing Then
        centerCDRS.Close
            centerCDRS = Nothing
    End If
    
    'If Not GetObjectContext Is Nothing Then
        'GetObjectContext.SetAbort
        'End If

    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
            'Err.Number, Err.Source, Err.Description
End Sub


'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckParameter_CorrectForm
'*【処理概要】
'*?    CSVファイルのパラメータに対して、DBに依存しない形式チェックを行います。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckParameter_CorrectForm; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncCheckParameter_CorrectForm(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheckParameter_CorrectForm = False
    
    Dim blnIsOKRequired
    Dim blnIsOKNumeric
    Dim blnIsOKLimit
    Dim blnIsOKHenkoubi
    Dim blnIsOKCheckDisit
'********** 2005/05/07 北澤 **********↓
    Dim blnIsOKNotZero
'********** 2005/05/07 北澤 **********↑
    
    'a.必須項目（該当店コード、変更日、JANコード）が入っているか確認する。
    blnIsOKRequired = FncCheck_RequiredParameter(arryCSVRecord(), intNowLine)

    'b.全ての項目において、半角数字であるか確認する。
    blnIsOKNumeric = FncCheck_IsParameterNumeric(arryCSVRecord(), intNowLine)
    
    'c.桁数が制限を超えていないか確認する。
    blnIsOKLimit = FncCheck_Limit(arryCSVRecord(), intNowLine)

    'd.変更日の日付が有効であるか確認する。
    blnIsOKHenkoubi = FncCheck_Henkoubi(arryCSVRecord(enmTenColNo.HenkouDate), intNowLine)
    
    'e.JANコードに正しいチェックデジットの値がセットされているか、
'********** 2005/04/26 黒沢 変更 **********↓
    blnIsOKCheckDisit = FncCheck_CheckDigit(arryCSVRecord(enmTenColNo.janCD), intNowLine, 1)
'    blnIsOKCheckDisit = FncCheck_CheckDigit(arryCSVRecord(enmTenColNo.JanCd), intNowLine)
'********** 2005/04/26 黒沢 変更 **********↑
    
'********** 2005/05/07 北澤 **********↓
'金額類が正の値かをチェックする
'********** 2005/06/16 北澤 **********↓
'売価変更理由は、店=0000かつ売価<>""の場合は入力必須、それ以外の場合は入力不可 のチェックを追加
'********** 2005/06/16 北澤 **********↑
    blnIsOKNotZero = FncCheck_NotZero(arryCSVRecord(), intNowLine, "TEN")
'********** 2005/05/07 北澤 **********↑
    
'********** 2005/06/16 北澤 **********↓
'    If blnIsOKRequired And blnIsOKNumeric And blnIsOKLimit And blnIsOKHenkoubi And blnIsOKCheckDisit Then
    If blnIsOKRequired And blnIsOKNumeric And blnIsOKLimit And _
       blnIsOKHenkoubi And blnIsOKCheckDisit And blnIsOKNotZero Then
'********** 2005/06/16 北澤 **********↑
        FncCheckParameter_CorrectForm = True
    End If
    
End Function
'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckParameter_CorrectForm_Modosi
'*【処理概要】
'*?    CSVファイルのパラメータに対して、DBに依存しない形式チェックを行います。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckParameter_CorrectForm_Modosi; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2009/06/06; Supattra T.(Tostem Thai); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncCheckParameter_CorrectForm_Modosi(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheckParameter_CorrectForm_Modosi = False
    
    Dim blnIsOKRequired
    Dim blnIsOKNumeric
    Dim blnIsOKLimit
    Dim blnIsOKHenkoubi
    Dim blnIsOKCheckDisit
    Dim blnIsOKNotZero
    
    'a.必須項目（該当店コード、変更日、JANコード）が入っているか確認する。
    blnIsOKRequired = FncCheck_RequiredParameter_Modosi(arryCSVRecord(), intNowLine)

    'b.全ての項目において、半角数字であるか確認する。
    blnIsOKNumeric = FncCheck_IsParameterNumeric_Modosi(arryCSVRecord(), intNowLine)
    
    'c.桁数が制限を超えていないか確認する。
    blnIsOKLimit = FncCheck_Limit_Modosi(arryCSVRecord(), intNowLine)

    'd.変更日の日付が有効であるか確認する。
    blnIsOKHenkoubi = FncCheck_Henkoubi(arryCSVRecord(enmTenColNo.HenkouDate), intNowLine)
    
    'e.JANコードに正しいチェックデジットの値がセットされているか、
    blnIsOKCheckDisit = FncCheck_CheckDigit(arryCSVRecord(enmTenColNo.janCD), intNowLine, 1)
    
    '売価変更理由は、店=0000かつ売価<>""の場合は入力必須、それ以外の場合は入力不可 のチェックを追加
    blnIsOKNotZero = FncCheck_NotZero(arryCSVRecord(), intNowLine, "TEN")
    
    If blnIsOKRequired And blnIsOKNumeric And blnIsOKLimit And _
       blnIsOKHenkoubi And blnIsOKCheckDisit And blnIsOKNotZero Then
        FncCheckParameter_CorrectForm_Modosi = True
    End If
    
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckParameter_PDC_Modosi
'*【処理概要】
'*?    CSVファイルのパラメータに対して、「変更日」、「JANコード」が同一の行単位でチェックする。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckParameter_PDC_Modosi; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2012/08/10; 楊双(大連); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncCheckParameter_PDC_Modosi(ByRef arryCSVRecord() As String, intNowLine As Integer, _
                                        ByRef arryCheckRecord As Object, ByRef flgFirst As Boolean) As Boolean

    FncCheckParameter_PDC_Modosi = True
    
    '以下は「変更日」、「JANコード」が同一の行単位でチェックする。
    Dim aryLineInfo As Object
    Dim intLoopCnt As Integer
    
    If arryCSVRecord(enmModColNo.CenterKbn) = "2" Or arryCSVRecord(enmModColNo.TenCd) = "0000" Then
        '配列の有無判定
        If Not flgFirst Then
            '配列領域定義
            ReDim arryCheckRecord(0)
            'ワーク変数
            ReDim aryLineInfo(4)
            aryLineInfo(0) = arryCSVRecord(enmModColNo.HenkouDate)
            aryLineInfo(1) = arryCSVRecord(enmModColNo.janCD)
            aryLineInfo(2) = arryCSVRecord(enmModColNo.Genka)
            aryLineInfo(3) = arryCSVRecord(enmModColNo.Hacchuutani)
            aryLineInfo(4) = "'" & arryCSVRecord(enmModColNo.TenCd) & "'"
            arryCheckRecord(0) = aryLineInfo
            flgFirst = True
          Else
    
            For intLoopCnt = 0 To UBound(arryCheckRecord, 1)
                If arryCheckRecord(intLoopCnt)(0) = arryCSVRecord(enmModColNo.HenkouDate) _
                And arryCheckRecord(intLoopCnt)(1) = arryCSVRecord(enmModColNo.janCD) Then
                
                    If arryCheckRecord(intLoopCnt)(2) <> arryCSVRecord(enmModColNo.Genka) Then
                        If arryCheckRecord(intLoopCnt)(2) = "" Then
                            arryCheckRecord(intLoopCnt)(2) = arryCSVRecord(enmModColNo.Genka)
                        End If
                       'エラーの場合
                        FncSetErrorPDC intNowLine, enmModColNo.Genka, PARAM_GENKA, arryCSVRecord(enmModColNo.HenkouDate) & "," & arryCSVRecord(enmModColNo.janCD)
                        FncCheckParameter_PDC_Modosi = False
                    End If
                    If arryCheckRecord(intLoopCnt)(3) <> arryCSVRecord(enmModColNo.Hacchuutani) Then
                        If arryCheckRecord(intLoopCnt)(3) = "" Then
                            arryCheckRecord(intLoopCnt)(3) = arryCSVRecord(enmModColNo.Hacchuutani)
                        End If
                       'エラーの場合
                        FncSetErrorPDC intNowLine, enmModColNo.Hacchuutani, PARAM_HATTYUU_TANI, arryCSVRecord(enmModColNo.HenkouDate) & "," & arryCSVRecord(enmModColNo.janCD)
                        FncCheckParameter_PDC_Modosi = False
                    End If
                    
                    'If arryCheckRecord(intLoopCnt)(4) <> "0000" Then
                        arryCheckRecord(intLoopCnt)(4) = arryCheckRecord(intLoopCnt)(4) & ",'" & arryCSVRecord(enmModColNo.TenCd) & "'"
                    'End If
                                        
                    Exit Function
                    
                End If
            Next
            
            ReDim Preserve arryCheckRecord(UBound(arryCheckRecord) + 1)
            'ワーク変数
            ReDim aryLineInfo(4)
            aryLineInfo(0) = arryCSVRecord(enmModColNo.HenkouDate)
            aryLineInfo(1) = arryCSVRecord(enmModColNo.janCD)
            aryLineInfo(2) = arryCSVRecord(enmModColNo.Genka)
            aryLineInfo(3) = arryCSVRecord(enmModColNo.Hacchuutani)
            aryLineInfo(4) = "'" & arryCSVRecord(enmModColNo.TenCd) & "'"
    
            arryCheckRecord(UBound(arryCheckRecord)) = aryLineInfo
            
        End If

    End If

    FncCheckParameter_PDC_Modosi = True
    
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_RequiredParameter
'*【処理概要】
'*?    CSVデータの１行に対し、必須項目が空白でないかどうかチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_RequiredParameter; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'

Private Function FncCheck_RequiredParameter(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_RequiredParameter = True
    
    If arryCSVRecord(enmTenColNo.TenCd) = "" Then       '該当店コード
        FncSetError (intNowLine, enmTenColNo.TenCd, PARAM_EMPTY
        FncCheck_RequiredParameter = False
    End If
    If arryCSVRecord(enmTenColNo.HenkouDate) = "" Then  '変更日
        FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_EMPTY
        FncCheck_RequiredParameter = False
    End If
    If arryCSVRecord(enmTenColNo.janCD) = "" Then       'JANコード
        FncSetError (intNowLine, enmTenColNo.janCD, PARAM_EMPTY
        FncCheck_RequiredParameter = False
    End If
    

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_RequiredParameter_Modosi
'*【処理概要】
'*?    CSVデータの１行に対し、必須項目が空白でないかどうかチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_RequiredParameter_Modosi; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*M    2011/02/10; 営業システム/和田;         センターコード追加;
'*M    2015/09/15; 流通システムG/和田;        TCシステム構築;
'*< Comment End >**************************************************'

Private Function FncCheck_RequiredParameter_Modosi(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_RequiredParameter_Modosi = True
    
    If arryCSVRecord(enmModColNo.TenCd) = "" Then       '該当店コード
        FncSetError (intNowLine, enmModColNo.TenCd, PARAM_EMPTY
        FncCheck_RequiredParameter_Modosi = False
    End If
    If arryCSVRecord(enmModColNo.HenkouDate) = "" Then  '変更日
        FncSetError (intNowLine, enmModColNo.HenkouDate, PARAM_EMPTY
        FncCheck_RequiredParameter_Modosi = False
    End If
    If arryCSVRecord(enmModColNo.janCD) = "" Then       'JANコード
        FncSetError (intNowLine, enmModColNo.janCD, PARAM_EMPTY
        FncCheck_RequiredParameter_Modosi = False
    End If
    '2015/09/15 wadak7 upd
    'If arryCSVRecord(enmModColNo.CenterKbn) = "1" Then       'センター区分が1：備蓄センターの場合、必須入力
    'センター区分が1：備蓄センター・3：TCセンターの場合、必須入力
    If arryCSVRecord(enmModColNo.CenterKbn) = "1" Or arryCSVRecord(enmModColNo.CenterKbn) = "3" Then
        If arryCSVRecord(enmModColNo.centerCD) = "" Then       'センターコード
            FncSetError (intNowLine, enmModColNo.centerCD, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
    Else    'センター区分が1：備蓄センター・3：TCセンター以外の場合、入力できない
        If arryCSVRecord(enmModColNo.centerCD) <> "" Then       'センターコード
            FncSetError (intNowLine, enmModColNo.centerCD, PARAM_KINSI
            FncCheck_RequiredParameter_Modosi = False
        End If
    End If
    
    '********** 2012/08/03 楊双（大連）入力チェックの追加 **********↓
    If arryCSVRecord(enmModColNo.CenterKbn) = "2" Then       '「センター区分」が"2"：PDCの場合
        If arryCSVRecord(enmModColNo.Hacchuutani) = "" Then          '「発注単位」が未入力の場合
            FncSetError (intNowLine, enmModColNo.Hacchuutani, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
        If arryCSVRecord(enmModColNo.Genka) = "" Then       '「原価」が未入力の場合
            FncSetError (intNowLine, enmModColNo.Genka, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
    End If
    '********** 2012/08/03 楊双（大連）入力チェックの追加 **********↑
    
    'letKikakuSyurui
    If pstrKikakuSyurui = "012" Then
        If arryCSVRecord(enmModColNo.MoGenka) = "" Then  '戻し原価
            FncSetError (intNowLine, enmModColNo.MoGenka, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
        If arryCSVRecord(enmModColNo.MoBaika) = "" Then  '戻し売価
            FncSetError (intNowLine, enmModColNo.MoBaika, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
        If arryCSVRecord(enmModColNo.HattyuuTani) = "" Then  '戻し発注単位
            FncSetError (intNowLine, enmModColNo.HattyuuTani, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
        If arryCSVRecord(enmModColNo.PopSize) = "" Then  'POPサイズ
            FncSetError (intNowLine, enmModColNo.PopSize, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
    ElseIf pstrKikakuSyurui = "015" Then
         If arryCSVRecord(enmModColNo.MoGenka) = "" Then  '戻し原価
            FncSetError (intNowLine, enmModColNo.MoGenka, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
        If arryCSVRecord(enmModColNo.HattyuuTani) = "" Then  '戻し発注単位
            FncSetError (intNowLine, enmModColNo.HattyuuTani, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
    End If

End Function
'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_IsParameterNumeric
'*【処理概要】
'*?    CSVデータの１行に対し、全フィールドが半角数値かどうかチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_IsParameterNumeric; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterNumeric(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_IsParameterNumeric = True
    
    If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.TenCd)) = False Then     '該当店コード
        FncSetError (intNowLine, enmTenColNo.TenCd, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.HenkouDate)) = False Then '変更日
        FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.janCD)) = False Then     'JANコード
        FncSetError (intNowLine, enmTenColNo.janCD, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric = False
    End If
    
    If arryCSVRecord(enmTenColNo.Genka) <> "" Then
        If IsNumeric(arryCSVRecord(enmTenColNo.Genka)) = False Then   '原価
'        If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.Genka)) = False Then '原価
            FncSetError (intNowLine, enmTenColNo.Genka, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric = False
        End If
    End If
    
    If arryCSVRecord(enmTenColNo.Baika) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.Baika)) = False Then   '売価
            FncSetError (intNowLine, enmTenColNo.Baika, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric = False
        End If
    End If
    
    If arryCSVRecord(enmTenColNo.BaikaRiyuu) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.BaikaRiyuu)) = False Then  '売価変更理由
            FncSetError (intNowLine, enmTenColNo.BaikaRiyuu, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric = False
        End If
    End If
    
    If arryCSVRecord(enmTenColNo.Torihikisaki) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.Torihikisaki)) = False Then '取引先コード
            FncSetError (intNowLine, enmTenColNo.Torihikisaki, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric = False
        End If
    End If
    
    If arryCSVRecord(enmTenColNo.Hacchuutani) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.Hacchuutani)) = False Then '発注単位
            FncSetError (intNowLine, enmTenColNo.Hacchuutani, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric = False
        End If
    End If

    If arryCSVRecord(enmTenColNo.CenterKbn) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.CenterKbn)) = False Then 'センター区分
            FncSetError (intNowLine, enmTenColNo.CenterKbn, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric = False
        End If
    End If

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_IsParameterNumeric_Modosi
'*【処理概要】
'*?    CSVデータの１行に対し、全フィールドが半角数値かどうかチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_IsParameterNumeric_Modosi; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2009/06/06; Supattra T.(Tostem Thai); 新規作成;
'*M    2011/02/10; 営業システム/和田;         センターコード追加;
'*M    2015/09/15; 流通システムG/和田;        TCシステム構築;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterNumeric_Modosi(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_IsParameterNumeric_Modosi = True
    
    If FncChkHankakuNumber(arryCSVRecord(enmModColNo.TenCd)) = False Then     '該当店コード
        FncSetError (intNowLine, enmModColNo.TenCd, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Modosi = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmModColNo.HenkouDate)) = False Then '変更日
        FncSetError (intNowLine, enmModColNo.HenkouDate, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Modosi = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmModColNo.janCD)) = False Then     'JANコード
        FncSetError (intNowLine, enmModColNo.janCD, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Modosi = False
    End If
    
    If arryCSVRecord(enmModColNo.Genka) <> "" Then
        If IsNumeric(arryCSVRecord(enmModColNo.Genka)) = False Then   '原価
            FncSetError (intNowLine, enmModColNo.Genka, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Modosi = False
        End If
    End If
    
    If arryCSVRecord(enmModColNo.Baika) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmModColNo.Baika)) = False Then   '売価
            FncSetError (intNowLine, enmModColNo.Baika, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Modosi = False
        End If
    End If
    
    If arryCSVRecord(enmModColNo.BaikaRiyuu) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmModColNo.BaikaRiyuu)) = False Then  '売価変更理由
            FncSetError (intNowLine, enmModColNo.BaikaRiyuu, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Modosi = False
        End If
    End If
    
    If arryCSVRecord(enmModColNo.Torihikisaki) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmModColNo.Torihikisaki)) = False Then '取引先コード
            FncSetError (intNowLine, enmModColNo.Torihikisaki, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Modosi = False
        End If
    End If
    
    If arryCSVRecord(enmModColNo.Hacchuutani) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmModColNo.Hacchuutani)) = False Then '発注単位
            FncSetError (intNowLine, enmModColNo.Hacchuutani, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Modosi = False
        End If
    End If

    If arryCSVRecord(enmModColNo.CenterKbn) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmModColNo.CenterKbn)) = False Then 'センター区分
            FncSetError (intNowLine, enmModColNo.CenterKbn, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Modosi = False
        End If
    End If
    
    If arryCSVRecord(enmModColNo.centerCD) <> "" Then
        '2015/09/15 wadak7 upd
        'If arryCSVRecord(enmModColNo.CenterKbn) = "1" Then      'センター区分が1：備蓄センターの場合、チェックする。
        'センター区分が1：備蓄センター・3：TCセンターの場合、チェックする。
        If arryCSVRecord(enmModColNo.CenterKbn) = "1" Or arryCSVRecord(enmModColNo.CenterKbn) = "3" Then
            If FncChkHankakuNumber(arryCSVRecord(enmModColNo.centerCD)) = False Then 'センターコード
                FncSetError (intNowLine, enmModColNo.centerCD, PARAM_NOTDISIT)
                FncCheck_IsParameterNumeric_Modosi = False
            End If
        End If
    End If
    
    If pstrKikakuSyurui = "012" Or pstrKikakuSyurui = "015" Then
    
        If arryCSVRecord(enmModColNo.MoGenka) <> "" Then
            If IsNumeric(arryCSVRecord(enmModColNo.MoGenka)) = False Then   '戻し原価
                FncSetError (intNowLine, enmModColNo.MoGenka, PARAM_NOTDISIT)
                FncCheck_IsParameterNumeric_Modosi = False
            End If
        End If
        
        If arryCSVRecord(enmModColNo.MoBaika) <> "" Then
            If FncChkHankakuNumber(arryCSVRecord(enmModColNo.MoBaika)) = False Then   '戻し売価
                FncSetError (intNowLine, enmModColNo.MoBaika, PARAM_NOTDISIT)
                FncCheck_IsParameterNumeric_Modosi = False
            End If
        End If
        
        If arryCSVRecord(enmModColNo.HattyuuTani) <> "" Then
            If FncChkHankakuNumber(arryCSVRecord(enmModColNo.HattyuuTani)) = False Then '戻し発注単位
                FncSetError (intNowLine, enmModColNo.HattyuuTani, PARAM_NOTDISIT)
                FncCheck_IsParameterNumeric_Modosi = False
            End If
        End If
        
        
    '2009.06.19 Del by Supattra T.(Tostem Thai)Start.
'    ElseIf pstrKikakuSyurui = "015" Then
'
'        If arryCSVRecord(enmModColNo.MoGenka) <> "" Then
'            If IsNumeric(arryCSVRecord(enmModColNo.MoGenka)) = False Then   '戻し原価
'                FncSetError (intNowLine, enmModColNo.MoGenka, PARAM_NOTDISIT)
'                FncCheck_IsParameterNumeric_Modosi = False
'            End If
'        End If
'
'        If arryCSVRecord(enmModColNo.HattyuuTani) <> "" Then
'            If FncChkHankakuNumber(arryCSVRecord(enmModColNo.HattyuuTani)) = False Then '戻し発注単位
'                FncSetError (intNowLine, enmModColNo.HattyuuTani, PARAM_NOTDISIT)
'                FncCheck_IsParameterNumeric_Modosi = False
'            End If
'        End If
     '2009.06.19 Del by Supattra T.(Tostem Thai)End.
    End If

End Function


'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncChkHankakuNumber
'*【処理概要】
'*?    指定された文字列が半角数字列かどうかをチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strValue ; String チェックしたい文字列
'*【戻り値】
'*    < Name Type Comment >
'*R    FncChkHankakuNumber; Boolean;半角数字列:True,半角数値以外が混じっている:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncChkHankakuNumber(ByVal strValue As String) As Boolean
        
    Dim lngCount As Long

    FncChkHankakuNumber = False

    For lngCount = 1 To Len(strValue)
        If Not Mid(strValue, lngCount, 1) Like "[0-9]" Then
            Exit Function
        End If
    Next
    
    FncChkHankakuNumber = True

End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncChkHankakuNarrow
'*【処理概要】
'*?    指定された文字列が半角文字列かどうかをチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strValue ; String チェックしたい文字列
'*【戻り値】
'*    < Name Type Comment >
'*R    FncChkHankakuNarrow; Boolean;半角文字列:True,半角文字列以外が混じっている:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/08/31; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncChkHankakuNarrow(ByVal strValue As String) As Boolean
        
    Dim lngCount As Long
    Dim LetterCount As Long          '文字列の文字数
    Dim ByteCount As Long            '文字列のバイト数
    
    FncChkHankakuNarrow = False

    ''半角チェック
    LetterCount = Len(strValue)
    ByteCount = LenB(StrConv(strValue, vbFromUnicode))

    If LetterCount = ByteCount Then
    Else
        Exit Function
    End If
    
    FncChkHankakuNarrow = True

End Function
'******************************************************************'
'*【プロシージャ名】
'*!    FncChkHankakuWide
'*【処理概要】
'*?    指定された文字列が全角文字列かどうかをチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strValue ; String チェックしたい文字列
'*【戻り値】
'*    < Name Type Comment >
'*R    FncChkHankakuWide; Boolean;半角文字列:True,半角文字列以外が混じっている:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/08/31; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncChkHankakuWide(ByVal strValue As String) As Boolean
        
    Dim lngCount As Long
    Dim LetterCount As Long          '文字列の文字数
    Dim ByteCount As Long            '文字列のバイト数
    
    FncChkHankakuWide = False

    ''全角チェック
    LetterCount = Len(strValue)
    ByteCount = LenB(StrConv(strValue, vbFromUnicode))

    If LetterCount <> ByteCount / 2 Then
        Exit Function
    End If
    
    FncChkHankakuWide = True

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_Limit
'*【処理概要】
'*?    CSVデータの１行に対し、フィールドが桁数制限を越えていないかチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_Limit; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_Limit(ByRef arryCSVRecord() As String, intNowLine As Integer)

    FncCheck_Limit = True
    
    '(a)該当店コード   ここではチェックしない（マスタ登録チェックで代用できる為）。
'    If FncCheck_TenCode(arryCSVRecord(enmTenColNo.TenCd)) = False Then
'        FncSetError (intNowLine, enmTenColNo.TenCd, PARAM_SHORTDISIT)
'        FncCheck_Limit = False
'    End If
    '(b)JANコード
    If FncCheck_JANCode(arryCSVRecord(enmTenColNo.janCD)) = False Then
        FncSetError (intNowLine, enmTenColNo.janCD, PARAM_SHORTDISIT)
        FncCheck_Limit = False
    End If
    '(c)原価
    If arryCSVRecord(enmTenColNo.Genka) <> "" Then
        If FncCheck_Genka(arryCSVRecord(enmTenColNo.Genka)) = False Then
            FncSetError (intNowLine, enmTenColNo.Genka, PARAM_SHORTDISIT)
            FncCheck_Limit = False
        End If
    End If
    '(d)売価
    If arryCSVRecord(enmTenColNo.Baika) <> "" Then
        If FncCheck_Baika(arryCSVRecord(enmTenColNo.Baika)) = False Then
            FncSetError (intNowLine, enmTenColNo.Baika, PARAM_SHORTDISIT)
            FncCheck_Limit = False
        End If
    End If
    '(e)発注単位
    If arryCSVRecord(enmTenColNo.Hacchuutani) <> "" Then
        If FncCheck_Hattyu_Tani(arryCSVRecord(enmTenColNo.Hacchuutani)) = False Then
            FncSetError (intNowLine, enmTenColNo.Hacchuutani, PARAM_SHORTDISIT)
            FncCheck_Limit = False
        End If
    End If
    
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_Limit_Modosi
'*【処理概要】
'*?    CSVデータの１行に対し、フィールドが桁数制限を越えていないかチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_Limit_Modosi; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2009/06/06; Supattra T.(Tostem Thai); 新規作成;
'*M    2011/02/10; 営業システム/和田;         センターコード追加;
'*M    2015/09/15; 流通システムG/和田;        TCシステム構築;
'*< Comment End >**************************************************'
Private Function FncCheck_Limit_Modosi(ByRef arryCSVRecord() As String, intNowLine As Integer)

    FncCheck_Limit_Modosi = True
    
    '(b)JANコード
    If FncCheck_JANCode(arryCSVRecord(enmModColNo.janCD)) = False Then
        FncSetError (intNowLine, enmModColNo.janCD, PARAM_SHORTDISIT)
        FncCheck_Limit_Modosi = False
    End If
    '(c)原価
    If arryCSVRecord(enmModColNo.Genka) <> "" Then
        If FncCheck_Genka(arryCSVRecord(enmModColNo.Genka)) = False Then
            FncSetError (intNowLine, enmModColNo.Genka, PARAM_SHORTDISIT)
            FncCheck_Limit_Modosi = False
        End If
    End If
    '(d)売価
    If arryCSVRecord(enmModColNo.Baika) <> "" Then
        If FncCheck_Baika(arryCSVRecord(enmModColNo.Baika)) = False Then
            FncSetError (intNowLine, enmModColNo.Baika, PARAM_SHORTDISIT)
            FncCheck_Limit_Modosi = False
        End If
    End If
    '(e)発注単位
    If arryCSVRecord(enmModColNo.Hacchuutani) <> "" Then
        If FncCheck_Hattyu_Tani(arryCSVRecord(enmModColNo.Hacchuutani)) = False Then
            FncSetError (intNowLine, enmModColNo.Hacchuutani, PARAM_SHORTDISIT)
            FncCheck_Limit_Modosi = False
        End If
    End If
    'センターコード
    If arryCSVRecord(enmModColNo.centerCD) <> "" Then
        '2015/09/15 wadak7 upd
        'If arryCSVRecord(enmModColNo.CenterKbn) = "1" Then      'センター区分が1：備蓄センターの場合、チェックする。
        'センター区分が1：備蓄センター・3：TCセンターの場合、チェックする。
        If arryCSVRecord(enmModColNo.CenterKbn) = "1" Or arryCSVRecord(enmModColNo.CenterKbn) = "3" Then
            If FncCheck_CenterCd(arryCSVRecord(enmModColNo.centerCD)) = False Then
                FncSetError (intNowLine, enmModColNo.centerCD, PARAM_SHORTDISIT)
                FncCheck_Limit_Modosi = False
            End If
        End If
    End If
    
    If Trim(pstrKikakuSyurui) = "012" Or Trim(pstrKikakuSyurui) = "015" Then
    
         '戻し原価
        If arryCSVRecord(enmModColNo.MoGenka) <> "" Then
            If FncCheck_Genka(arryCSVRecord(enmModColNo.MoGenka)) = False Then
                FncSetError (intNowLine, enmModColNo.MoGenka, PARAM_SHORTDISIT)
                FncCheck_Limit_Modosi = False
            End If
        End If
        
        '戻し売価
        If arryCSVRecord(enmModColNo.MoBaika) <> "" Then
            If FncCheck_Baika(arryCSVRecord(enmModColNo.MoBaika)) = False Then
                FncSetError (intNowLine, enmModColNo.MoBaika, PARAM_SHORTDISIT)
                FncCheck_Limit_Modosi = False
            End If
        End If
        
        '戻し発注単位
        If arryCSVRecord(enmModColNo.HattyuuTani) <> "" Then
            If FncCheck_Hattyu_Tani(arryCSVRecord(enmModColNo.HattyuuTani)) = False Then
                FncSetError (intNowLine, enmModColNo.HattyuuTani, PARAM_SHORTDISIT)
                FncCheck_Limit_Modosi = False
            End If
        End If
        'POPサイズ
        If arryCSVRecord(enmModColNo.PopSize) <> "" Then
            If LenB(StrConv(arryCSVRecord(enmModColNo.PopSize), vbFromUnicode)) > 20 Then
                FncSetError (intNowLine, enmModColNo.PopSize, PARAM_SHORTDISIT)
                FncCheck_Limit_Modosi = False
            End If
        End If
        
    '2009.06.19 Del by Supattra T.(Tostem Thai)Start.
'    ElseIf Trim(pstrKikakuSyurui) = "015" Then
'
'          '戻し原価
'        If arryCSVRecord(enmModColNo.MoGenka) <> "" Then
'            If FncCheck_Genka(arryCSVRecord(enmModColNo.MoGenka)) = False Then
'                FncSetError (intNowLine, enmModColNo.MoGenka, PARAM_SHORTDISIT)
'                FncCheck_Limit_Modosi = False
'            End If
'        End If
'
'         '戻し発注単位
'        If arryCSVRecord(enmModColNo.HattyuuTani) <> "" Then
'            If FncCheck_Hattyu_Tani(arryCSVRecord(enmModColNo.HattyuuTani)) = False Then
'                FncSetError (intNowLine, enmModColNo.HattyuuTani, PARAM_SHORTDISIT)
'                FncCheck_Limit_Modosi = False
'            End If
'        End If
     '2009.06.19 Del by Supattra T.(Tostem Thai)End.
        
    End If
    
    
End Function


'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_TenCode
'*【処理概要】
'*?    店コードの文字列長チェックを行います。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strTenCode ; String 店コード
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_TenCode; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
'このチェックは行わない（マスタ登録チェックで代用できる為）。
'Private Function FncCheck_TenCode(strTenCode As String) As Boolean
'    '該当店コード：整数4桁まで
'    FncCheck_TenCode = FncCheck_Length_Less(strTenCode, LENGTH_TENCODE)
'End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_Baika
'*【処理概要】
'*?    売価の文字列長チェックを行います。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strBaika ; String 売価
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_Baika; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_Baika(strBaika As String) As Boolean
    '売価：整数7桁まで
    FncCheck_Baika = FncCheck_Length_Less(strBaika, LENGTH_BAIKA)
End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncCheck_Length_Kotei
'*【処理概要】
'*?    文字列の文字列長チェックを行います。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strNumberString ; String 文字列
'*P    lngLength       ; Long   桁数
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_Length_Kotei; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/08/31; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_Length_Kotei(strNumberString As String, lngLength As Long) As Boolean
    FncCheck_Length_Kotei = False
    If Len(strNumberString) = lngLength Then
        FncCheck_Length_Kotei = True
    End If
End Function


'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_Hattyu_Tani
'*【処理概要】
'*?    発注単位の文字列長チェックを行います。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strHattyu_Tani ; String 発注単位
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_Hattyu_Tani; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_Hattyu_Tani(strHattyu_Tani As String) As Boolean
    '発注単位：整数5桁まで
    FncCheck_Hattyu_Tani = FncCheck_Length_Less(strHattyu_Tani, LENGTH_HATTYU_TANI)
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_Length_Less
'*【処理概要】
'*?    指定された文字列が、指定された桁数以下であるかどうかを調べます。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strNumberString ; String 文字列
'*P    lngLength ; Long 制限文字列長
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_Length_Less; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_Length_Less(strNumberString As String, lngLength As Long)
    FncCheck_Length_Less = False
    If Len(strNumberString) <= lngLength Then
        FncCheck_Length_Less = True
    End If
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_JANCode
'*【処理概要】
'*?    JANコードの文字列長チェックを行います。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strJanCode ; String JANコード
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_JANCode; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_JANCode(strJanCode As String) As Boolean
    FncCheck_JANCode = False
    'JANコード：8桁又は13桁
    If Len(strJanCode) = LENGTH_JANCODE1 Or Len(strJanCode) = LENGTH_JANCODE2 Then
        FncCheck_JANCode = True
    End If
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_CenterCd
'*【処理概要】
'*?    JANコードの文字列長チェックを行います。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strCenterCd ; String センターコード
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_CenterCd; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2011/02/10; 営業システム/和田; 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_CenterCd(strCenterCd As String) As Boolean
    FncCheck_CenterCd = False
    'センターコード：4桁
    If Len(strCenterCd) = LENGTH_CENTERCODE Then
        FncCheck_CenterCd = True
    End If
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_Genka
'*【処理概要】
'*?    原価の文字列長チェックを行います。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strGenka ; String 原価
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_Genka; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_Genka(strGenka As String) As Boolean
    '原価：整数部７桁、小数部２桁
    FncCheck_Genka = FncCheck_Syousuu(strGenka, LENGTH_GENKA_SEISU, LENGTH_GENKA_SYOUSUU)
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_Syousuu
'*【処理概要】
'*?    小数形式の実数に対し、指定された桁数以下になっているかどうかを調べます。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strValue ; String 小数文字列
'*P    intSeisuuKeta ; Integer 整数部最大桁数
'*P    intSyousuuKeta ; Integer 小数部最大桁数
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_Syousuu; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_Syousuu(ByVal strValue As String, _
                    ByVal intSeisuuKeta As Integer, _
                    ByVal intSyousuuKeta As Integer) As Boolean

    Dim lngCount As Long
    Dim arrySyousuu() As String

    FncCheck_Syousuu = False
    
    If IsNumeric(strValue) = False Then
        ' 数値ではない
        Exit Function
    End If
    
    For lngCount = 1 To Len(strValue)
        If Not Mid(strValue, lngCount, 1) Like "[0-9.]" Then
            ' 数値以外の文字が入っている
            Exit Function
        End If
            
        '整数が入力された場合
        If InStr(strValue, ".") = 0 Then
            '桁あふれチェック
            If Len(strValue) > intSeisuuKeta Then
              
                Exit Function
            End If
        '小数の場合
        Else
            arrySyousuu = Split(strValue, ".")
            '桁あふれチェック
            If Len(arrySyousuu(0)) > intSeisuuKeta Then
                Exit Function
            End If
            
            '桁あふれチェック
            If Len(arrySyousuu(1)) > intSyousuuKeta Then
                '小数部桁あふれ
                Exit Function
            End If
                    
        End If
    
    Next
    
    FncCheck_Syousuu = True

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_Henkoubi
'*【処理概要】
'*?    CSVデータの変更日に対しチェックを行います。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strHenkoubi ; String 変更日を示す文字列
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_Henkoubi; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_Henkoubi(strHenkoubi As String, intNowLine As Integer) As Boolean

    FncCheck_Henkoubi = True
    ' 桁数が８桁である事と、変更日の日付が有効であるかを確認する。
    If Len(strHenkoubi) <> 8 Or FncCheckDate(strHenkoubi) = False Then
        FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_WRONGDATE)
        FncCheck_Henkoubi = False
'********** 2005/05/18 土田 追加 **********↓
    Else
        If pstrMinDate > strHenkoubi Then
            pstrMinDate = strHenkoubi
        ElseIf pstrMaxDate < strHenkoubi Then
            pstrMaxDate = strHenkoubi
        End If
'********** 2005/05/18 土田 追加 **********↑
    End If

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckDate
'*【処理概要】
'*?    日付文字列８桁が正しい日付であることをチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strDate ; YYYYMMDDで示される８桁の日付文字列
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckDate; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*M    2005/04/27; 黒沢(USC); 日付チェックをCSV取込に合わせるように変更。;
'*M                          （19000101 ～ 90001231）;
'*< Comment End >**************************************************'
Private Function FncCheckDate(strDate As String) As Boolean
    
    Dim dtChk    As Date
    FncCheckDate = False
    
    '範囲チェック 19000101 ～ 90001231 とする
    If StrComp(strDate, 90001231) > 0 Then
        FncCheckDate = False
        Exit Function
    End If
    
    If StrComp(strDate, 19000101) < 0 Then
        FncCheckDate = False
        Exit Function
    End If
    
    '空白が文字間にある場合は、エラー
    If strDate <> Replace(strDate, " ", "") Then
        FncCheckDate = False
        Exit Function
    End If
    
    Dim strCheckDate
    strCheckDate = Format$(strDate, "@@@@/@@/@@")
    FncCheckDate = IsDate(strCheckDate)
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheck_CheckDigit
'*【処理概要】
'*?    CSVデータのJANコードに対しチェックディジットを調べます。
'*?    チェックディジットはウェイト 3-1 のモジュラス10によって
'*?    計算され、調べるJANコードの末尾に追加されているものとします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strJanCode ; String JANコードを示す文字列
'*P    intNowLine; Integer; CSVファイルの現在の行
'*P    intFieldType; Integer; 1:店別情報のJANコード
'*P    　　　　　　　　　　　　2:発注中止日情報のJANコード
'*P    　　　　　　　　　　　　3:発注中止日情報の代替商品コード
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_CheckDigit; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*M    2005/04/26; 黒沢(USC); 発注中止用でも使用できるように引数を追加;
'*< Comment End >**************************************************'
Private Function FncCheck_CheckDigit(strJanCode As String, intNowLine As Integer, intFieldType As Integer) As Boolean

    Dim varTmp As Object
    Dim strChkDigit As String
    Dim strCode As String
    Dim lngFactor As Long
    Dim lngSum As Long
    Dim lngCount As Long
    
    FncCheck_CheckDigit = False

'********** 2005/09/05 高附 追加 **********↓
'********** 2005/04/26 黒沢 追加 **********↓
'項目名の作成（エラーメッセージ用）
    Dim intField As Integer
    Select Case intFieldType
        Case 1:     intField = enmTenColNo.janCD
        Case 2:     intField = enmHacColNo.janCD
        Case 3:     intField = enmHacColNo.DaiJanCd
        Case 4:     intField = enmSyoColNo.janCD
        Case 5:     intField = enmBunColNo.janCD
        Case Else:  intField = enmTenColNo.janCD
    End Select
'********** 2005/04/26 黒沢 追加 **********↑
'********** 2005/09/05 高附 追加 **********↑

    '名称マスタチェック
    lngFactor = 3
    lngSum = 0

    varTmp = strJanCode

    '桁数チェック
    If Len(varTmp) <> LENGTH_JANCODE1 And Len(varTmp) <> LENGTH_JANCODE2 Then
'********** 2005/04/26 黒沢 変更 **********↓
        FncSetError (intNowLine, intField, PARAM_WRONGCHECKDIGIT
'        FncSetError (intNowLine, enmTenColNo.JanCd, PARAM_WRONGCHECKDIGIT
'********** 2005/04/26 黒沢 変更 **********↑
        GoTo Proc_End
    End If
            
    'JANコードを最大桁に併せて0埋め
    varTmp = Format(varTmp, String(LENGTH_JANCODE2, "0"))
            
    'チェックデジット取得
    strChkDigit = Right(varTmp, 1)
            
    'コード部取得
    strCode = Left(varTmp, Len(varTmp) - 1)
    
    If IsNumeric(strCode) Then
        For lngCount = Len(strCode) To 1 Step -1
            lngSum = lngSum + Mid(strCode, lngCount, 1) * lngFactor
            lngFactor = 4 - lngFactor
        Next
            
        'チェックデジット確認
        If strChkDigit <> ((1000 - lngSum) Mod 10) Then
    '********** 2005/04/26 黒沢 変更 **********↓
            FncSetError (intNowLine, intField, PARAM_WRONGCHECKDIGIT
    '        FncSetError (intNowLine, enmTenColNo.JanCd, PARAM_WRONGCHECKDIGIT
    '********** 2005/04/26 黒沢 変更 **********↑
            GoTo Proc_End
        End If
    End If
    
    FncCheck_CheckDigit = True

Proc_End:

End Function
'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckParameter_DataBase
'*【処理概要】
'*?    CSVファイルのパラメータに対して、DBにアクセスしてチェックを行います。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVRecord() ; String CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckParameter_DataBase; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*M    2011/02/16; 営業システム/和田; センターコード追加;
'*< Comment End >**************************************************'
Private Function FncCheckParameter_DataBase(ByRef arryCSVRecord() As String, intNowLine As Integer)

    Dim blnIsOKTenCode As Boolean
    Dim blnIsOKJANCode As Boolean
    Dim blnIsOKNaibuCode As Boolean
    Dim blnIsOKBaikaHenkouRiyuu As Boolean
    Dim blnIsOKTorihikisaki As Boolean
    Dim strIsOKCenterClass As String
    Dim blnIsOKSakituke As Boolean
    
    Dim strKeyCode As String

    FncCheckParameter_DataBase = False
    
    'a.該当店コードが事業所マスタ・WF名称マスタテーブルに登録されているか確認する。
    blnIsOKTenCode = FncCheckDB_TenCode(arryCSVRecord(enmTenColNo.TenCd), intNowLine)
    
    'b.JANコードが有効か確認する。
'********** 2005/04/26 黒沢 変更 **********↓
    blnIsOKJANCode = FncCheckDB_JANCode(arryCSVRecord(enmTenColNo.janCD), strKeyCode, intNowLine, RET_FLD_TENB)
'    blnIsOKJANCode = FncCheckDB_JANCode(arryCSVRecord(enmTenColNo.JanCd), strKeyCode, intNowLine)
'********** 2005/04/26 黒沢 変更 **********↓
    
'********** 2005/09/13 高附 変更 **********↓
    'b2.KEYコードが有効か確認する。
    blnIsOKNaibuCode = FncCheckDB_NaibuCode(strKeyCode, intNowLine, RET_FLD_TENB)
'********** 2005/09/13 高附 変更 **********↑
    
    'c.売価変更理由が有効か確認する。
    blnIsOKBaikaHenkouRiyuu = True
    If arryCSVRecord(enmTenColNo.BaikaRiyuu) <> "" Then
        blnIsOKBaikaHenkouRiyuu = FncCheckDB_BaikaHenkouRiyuu(arryCSVRecord(enmTenColNo.BaikaRiyuu), intNowLine)
    End If
    
    'd.取引先が有効か確認する。
    blnIsOKTorihikisaki = True
    If arryCSVRecord(enmTenColNo.Torihikisaki) <> "" Then
        blnIsOKTorihikisaki = FncCheckDB_Torihikisaki(arryCSVRecord(enmTenColNo.Torihikisaki), intNowLine)
    End If
    
    'e.センター区分、センターコードが有効か確認する。
    strIsOKCenterClass = "0"    '0:正常 1:センター区分異常 2:センターコード異常
    If arryCSVRecord(enmTenColNo.CenterKbn) <> "" Then
        strIsOKCenterClass = FncCheckDB_CenterClass(arryCSVRecord(enmTenColNo.CenterKbn), arryCSVRecord(enmModColNo.centerCD), intNowLine)
    End If
    
    'f.統合マスタの先付変更情報に同日の変更日情報が無いか確認する。
    blnIsOKSakituke = True
    If blnIsOKTenCode = True Then   'a.のチェックがOKの場合にチェックする。
        blnIsOKSakituke = FncCheckDB_Sakituke(strKeyCode, arryCSVRecord, intNowLine)
    End If
    
    If blnIsOKTenCode And blnIsOKJANCode And blnIsOKBaikaHenkouRiyuu And _
       blnIsOKTorihikisaki And strIsOKCenterClass = "0" And blnIsOKSakituke And blnIsOKNaibuCode Then
        FncCheckParameter_DataBase = True
    End If
 
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckDB_TenCode
'*【処理概要】
'*?    該当店コードが事業所マスタ・WF名称マスタに登録されているかどうか確認します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strTenCode ; String 店コードをあらわす文字列
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckDB_TenCode; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheckDB_TenCode(ByVal strTenCode As String, intNowLine As Integer) As Boolean
    
'********** 2005/05/11 土田 使用していないので削除 **********↓
'    Dim blnIsFindTenCode As Boolean
'    Dim blnIsFindJigyouCode As Boolean
'    Dim intLoop As Integer
'    Dim strJigyouCode As String
'********** 2005/05/11 土田 削除 **********↑

    FncCheckDB_TenCode = True
    
    If FncCheckDB_FindTenCode_TenCode(strTenCode) = False Then
        ' 指定された該当店コードは展開していない
        FncSetError (intNowLine, enmTenColNo.TenCd, PARAM_NOTEXISTDB
        FncCheckDB_TenCode = False
    End If
    
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckDB_FindTenCode_TenCode
'*【処理概要】
'*?    該当店コードがリテールDB事業所マスタテーブルに存在するかを確認します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strTenCode ; String 店コードをあらわす文字列
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckDB_FindTenCode_TenCode; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*M    2005/05/11; 土田　祐歩(USC); パフォーマンス向上の為､この関数内での切断レコード取得をやめ大域変数の切断レコードを利用する様修正;
'*< Comment End >**************************************************'
Private Function FncCheckDB_FindTenCode_TenCode(strTenCode As String) As Boolean

'********** 2005/05/11 土田 削除 **********↓
'DB接続に関する処理を削除
'    Dim strSql            As String
'    Dim objCDBAccess      As Object
'    Dim objRS      As Object
'********** 2005/05/11 土田 追加 **********↑
            
    FncCheckDB_FindTenCode_TenCode = False
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncCheckDB_FindTenCode_TenCode(" & strTenCode & ")"
 
    ' 本部(0000)が指定された場合はチェックせずに、(0000)をそのまま店コードとして返す
    If StrComp(strTenCode, HON_TEN_CD) = 0 Then
        FncCheckDB_FindTenCode_TenCode = True
'********** 2005/05/11 土田 変更 **********↓
'DB接続に関する処理を削除
'        GoTo Proc_End
        Exit Function
'********** 2005/05/11 土田 変更 **********↑
    End If
    
    
'********** 2005/05/11 土田 削除 **********↓
'この関数内での切断レコードセット取得処理削除
'    'SQL文作成（条件：判別区分=5、閉店日がNULLor一括登録日以降、KEY区分=W02、設定内容の4桁目が1、店コード）
'    strSql = "SELECT A.jigyousyo_cd "
'    strSql = strSql & "FROM v_jigyousyo AS A WITH(readcommitted) "
'    strSql = strSql & "INNER JOIN m_wf_meisyou AS B WITH(readcommitted) ON A.jigyou_cd = B.key_cd " & vbCrLf
'    strSql = strSql & "WHERE A.hanbetu_kbn = '" & HANBETU_KBN_TENPO & "' AND "
'    strSql = strSql & "(RTrim(A.heiten_date) = '' OR A.heiten_date is null OR A.heiten_date >= " & FncGetNowDate() & ") "
'    strSql = strSql & "AND Substring(B.settei_naiyou,4,1) = '1' " & vbCrLf
'    strSql = strSql & "AND B.key_kbn = '" & KEYCD_KIKAKU_DOUNYU_TAISYOU & "' " & vbCrLf
'    '全店、全事業の場合は、以下の条件は必要ない。
'    If strTenCode <> ALL_TEN_CD And strTenCode <> ALL_JIG_CD Then
'        If Len(strTenCode) = 2 Then
'            strSql = strSql & "AND A.jigyou_cd = '" & strTenCode & "'"      '各事業
'        Else
'            strSql = strSql & "AND A.jigyousyo_cd = '" & strTenCode & "'"   '各店
'        End If
'    End If
    
'    'CDBAccessインスタンス生成
'    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
'    'SQL実行
'        objRS = objCDBAccess.GetDCRecordset("VSWF", strSql)
'********** 2005/05/11 土田 削除 **********↑
    
'********** 2005/05/11 土田 追加 **********↓
'大域変数の切断レコードセットを使用しエラーチェック

    

    If strTenCode <> ALL_TEN_CD And strTenCode <> ALL_JIG_CD Then
        
        Dim strTarget As String '検索対象（事業か店か）
        If Len(strTenCode) = 2 Then
            strTarget = "jigyou_cd"      '各事業
        Else
            strTarget = "jigyousyo_cd"  '各店
        End If
        
        With pTChkJigyouRS

'********** 2005/05/17 土田 変更 **********↓
'********** 2005/05/18 土田 変更 **********↓
'            If Not .EOF Then
           .Filter = adFilterNone
            If Not .BOF Then
'********** 2005/05/18 土田 変更 **********↑
                .MoveFirst
                .Find strTarget & " = '" & strTenCode & "'"
            End If
'********** 2005/05/17 土田 変更 **********↑
        End With
    
    End If
    
    
    '店（事業）コードが無い場合はエラー
    FncCheckDB_FindTenCode_TenCode = Not pTChkJigyouRS.EOF
'********** 2005/05/11 土田 追加 **********↑

'********** 2005/05/11 土田 削除 **********↓
'DB接続に関する処理を削除
'    If objRS.EOF = True Then
'        FncCheckDB_FindTenCode_TenCode = False
'    Else
'        FncCheckDB_FindTenCode_TenCode = True
'    End If
    
    'DBクローズ
'    objRS.Close
'        objRS = Nothing
'    objCDBAccess = Nothing
    
'    'GetObjectContext.SetComplete
'********** 2005/05/11 土田 削除 **********↑
Exit Function

'********** 2005/05/11 土田 変更 **********↓
'DB接続に関する処理を削除
'Proc_End:
'    If Not objRS Is Nothing Then
'            objRS = Nothing
'    End If
'    If Not objCDBAccess Is Nothing Then
'        objCDBAccess = Nothing
'    End If
'    Exit Function

DefectErrorHandler:
'    If Not objRS Is Nothing Then
'            objRS = Nothing
'    End If
'    If Not objCDBAccess Is Nothing Then
'        objCDBAccess = Nothing
'    End If
'    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description
'********** 2005/05/11 土田 変更 **********↑

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncGetNowDate
'*【処理概要】
'*?    現在の日付を８文字の半角数値で返します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【戻り値】
'*    < Name Type Comment >
'*R    FncGetNowDate; String;YYYYMMDDで表現される８桁の日付
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncGetNowDate() As String

    Dim dtNowDate, strNowDate
    dtNowDate = Date
    strNowDate = Format(dtNowDate, "YYYYMMDD")
    
    FncGetNowDate = strNowDate

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckDB_JANCode
'*【処理概要】
'*?    JANコードが有効かをチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strJanCode ; String JANコードをあらわす文字列
'*P    strKeyCode ; String JANコードに対応する統合マスタのキーコードを格納する領域
'*P    intNowLine; Integer; CSVファイルの現在の行
'*P    intFieldType; Integer; RET_FLD_HACH:発注中止日情報
'*P    　　　　　　　　　　　　 RET_FLD_TENB:店別情報
'*【戻り値】
'*    < Name Type Comment >
'*R    strKeyCode ; String JANコードに対応する統合マスタのキーコードを格納する
'*R    FncCheckDB_JANCode; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*M    2005/04/26; 黒沢(USC); 発注中止でも使用できるように引数の追加;
'*M    2006/09/12; 和田(IC); 商品選定ＷＦ問合せ撲滅改善;
'*M    2009/05/12; 和田(IC); グループ品はエラーにする;
'*< Comment End >**************************************************'
'********** 2005/04/26 黒沢 変更 **********↓
Private Function FncCheckDB_JANCode(ByVal strJanCode As String, ByRef strKeyCode As String, intNowLine As Integer, intFieldType As Integer) As Boolean
'Private Function FncCheckDB_JANCode(ByVal strJanCode As String, ByRef strKeyCode As String, intNowLine As Integer) As Boolean
'********** 2005/04/26 黒沢 変更 **********↑
    Dim strKikakuNo As String   'JANコードが既に登録されている企画№
    
    'b.JANコードが有効か確認する。
    FncCheckDB_JANCode = True
    
'********** 2005/04/26 黒沢 追加 **********↓
'項目名の作成（エラーメッセージ用）
    Dim intField As Integer
    Select Case intFieldType
        Case RET_FLD_HACH:      intField = enmHacColNo.janCD
        Case RET_FLD_TENB:      intField = enmTenColNo.janCD
        Case RET_FLD_SYO:       intField = enmSyoColNo.janCD
        Case RET_FLD_BUN:       intField = enmBunColNo.janCD
        Case Else:              intField = enmTenColNo.janCD
    End Select
'********** 2005/04/26 黒沢 追加 **********↑
    
    '　(a)WFマスタに同一のJANコードが無い事を確認する。
    If FncCheckDB_SameJANCode(strJanCode, strKikakuNo) = True Then
'********** 2005/04/26 黒沢 変更 **********↓
        FncSetError (intNowLine, intField, PARAM_SAMECODEDB, strKikakuNo
'        FncSetError (intNowLine, enmTenColNo.JanCd, PARAM_SAMECODEDB, strJanCode
'********** 2005/04/26 黒沢 変更 **********↑
        FncCheckDB_JANCode = False
    End If
            
    '　(b)統合マスタに登録済みか確認する。
    If FncCheckDB_EntryJANCode(strJanCode, strKeyCode) = False Then
'********** 2005/04/26 黒沢 変更 **********↓
        FncSetError (intNowLine, intField, PARAM_NOTEXISTDB, strJanCode
'        FncSetError (intNowLine, enmTenColNo.JanCd, PARAM_NOTEXISTDB, strJanCode
'********** 2005/04/26 黒沢 変更 **********↑
        FncCheckDB_JANCode = False
    End If

    'グループ品の場合、エラー
    If intFieldType = RET_FLD_TENB Then
        If FncCheckDB_Group(strJanCode) = False Then
            FncSetError (intNowLine, intField, PARAM_GROUP, strJanCode
            FncCheckDB_JANCode = False
        End If
    End If

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckDB_SameJANCode
'*【処理概要】
'*?    商品選定DB（WF管理項目）テーブルに重複したJANコードが存在するかを調べます。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strJanCode ; String JANコードをあらわす文字列
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckDB_SameJANCode; Boolean;重複して存在:True,存在しない:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*M    2005/05/11; 土田　祐歩(USC); パフォーマンス向上の為､この関数内での切断レコード取得をやめ大域変数の切断レコードを利用する様修正;
'*M    2006/09/12; 和田(IC); 商品選定ＷＦ問合せ撲滅改善;
'*< Comment End >**************************************************'
Private Function FncCheckDB_SameJANCode(ByVal strJanCode As String, _
                                  ByRef strKikakuNo As String) As Boolean

    '   　DB：      商品選定DB（WF管理項目）テーブル（t_syousen_wf_kanri）
    '   　キー：        JANコード（jan_cd）                         CSVファイルから取得
    '   遷移時に取得した企画種類コード・年連番・枝番を持つJANコード以外で
    '   登録されていないか確認する。
    
'********** 2005/05/11 土田 削除 **********↓
'この関数内での切断レコードセット取得処理削除
'    Dim strSql            As String
 
'    strSql = "SELECT jan_cd FROM t_syousen_wf_kanri WITH(readcommitted) WHERE jan_cd = '" & strJanCode & "'"
'    strSql = strSql & " AND NOT("
'    strSql = strSql & "kikaku_syurui_cd = '" & pstrKikakuSyurui & "' "
'    strSql = strSql & " AND nen_ren_no = '" & pstrKikakuNenrenban & "' "
'    strSql = strSql & " AND eda_no = '" & pstrKikakuEdaban & "'"
'    strSql = strSql & ")"
     
'    FncCheckDB_SameJANCode = FncCheckDB_IsRecordExists("VSWF", strSql)
'********** 2005/05/11 土田 削除 **********↑

'********** 2005/05/11 土田 追加 **********↓
'大域変数の切断レコードセットを使用しエラーチェック
    With pChkSameJANCodeRS
'********** 2005/05/17 土田 変更 **********↓
'********** 2005/05/18 土田 変更 **********↓
'        If Not .EOF Then
        If Not .BOF Then
'********** 2005/05/18 土田 変更 **********↑
            .MoveFirst
            .Find "jan_cd = '" & strJanCode & "'"
            If Not .EOF Then
                strKikakuNo = .Fields("kikaku_no").Value
            End If
        End If
'********** 2005/05/17 土田 変更 **********↑
    End With
 
    FncCheckDB_SameJANCode = Not pChkSameJANCodeRS.EOF
'********** 2005/05/11 土田 追加 **********↑

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckDB_EntryJANCode
'*【処理概要】
'*?    JANコードが統合マスタのKEY情報1テーブルに登録されているかを調べ、
'*?    登録されていればそのキーコードを返す。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strJanCode ; String JANコードをあらわす文字列
'*P    strKeyCode ; String キーコードを格納する文字列領域
'*【戻り値】
'*    < Name Type Comment >
'*     strKeyCode   ; String　登録されていればキーコードを返す
'*R    FncCheckDB_EntryJANCode; Boolean;登録されている:True,存在しない:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*M    2005/05/11; 土田　祐歩(USC); パフォーマンス向上の為､この関数内での切断レコード取得をやめ大域変数の切断レコードを利用する様修正;
'*< Comment End >**************************************************'
Private Function FncCheckDB_EntryJANCode(ByVal strJanCode, ByRef strKeyCode As String) As Boolean
    
    '   マスタ登録されているかは、統合マスタのKEY情報1テーブルで確認する。
    '   　DB：      統合マスタ KEY情報1（v_key_1）
    '   　キー：        JANコード（jan_cd）                         CSVファイルから取得
    '   JANコードと主キーが紐付けられている事を確認する。

'********** 2005/05/11 土田 削除 **********↓
'DB接続に関する処理を削除
'    Dim strSql            As String
'    Dim objCDBAccess      As Object
'    Dim objRS             As Object
'********** 2005/05/11 土田 削除 **********↑
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncCheckDB_EntryJANCode(" & strJanCode & "," & strKeyCode & ")"
    
    FncCheckDB_EntryJANCode = False

'********** 2005/05/11 土田 削除 **********↓
'この関数内での切断レコードセット取得処理削除
'    strSql = "SELECT key_cd , jan_cd FROM v_key_1 WITH(readcommitted) WHERE jan_cd = '" & strJanCode & "'"
    
'    'CDBAccessインスタンス生成
'    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")

'    'SQL実行
'        objRS = objCDBAccess.GetDCRecordset("VSWF", strSql)
    
'    ' データが取得できなければ、テーブル内には該当レコードは存在しない
'    If objRS.EOF = True Then
'        FncCheckDB_EntryJANCode = False
'        GoTo Proc_End
'    Else
'        ' データが取得できたら、そのキーコードを取得する
'        FncCheckDB_EntryJANCode = True
'        strKeyCode = objRS.Fields("key_cd")
'    End If
    
'    objRS.Close
'        objRS = Nothing
'    objCDBAccess = Nothing
    
'    'GetObjectContext.SetComplete
'********** 2005/05/11 土田 削除 **********↑

'********** 2005/05/11 土田 追加 **********↓
'大域変数の切断レコードセットを使用しエラーチェック
    ' データが取得できたら、そのキーコードを取得する
    With pChkvk1JANCodeRS
'********** 2005/05/17 土田 変更 **********↓
'********** 2005/05/18 土田 変更 **********↓
'         If Not .EOF Then
        If Not .BOF Then
'********** 2005/05/18 土田 変更 **********↑
            .MoveFirst
            .Filter = "jan_cd = '" & strJanCode & "' "
'            .Find "jan_cd = '" & strJanCode & "'"
            If .EOF = False Then
                strKeyCode = .Fields("key_cd")
            End If
        End If
'********** 2005/05/17 土田 変更 **********↑
    End With
 
    FncCheckDB_EntryJANCode = Not pChkvk1JANCodeRS.EOF
    pChkvk1JANCodeRS.Filter = adFilterNone
'********** 2005/05/11 土田 追加 **********↑
Exit Function
    
'********** 2005/05/11 土田 変更 **********↓
'DB接続に関する処理を削除
'Proc_End:
'    If Not objRS Is Nothing Then
'            objRS = Nothing
'    End If
'    If Not objCDBAccess Is Nothing Then
'        objCDBAccess = Nothing
'    End If
'    Exit Function
DefectErrorHandler:
'    If Not objRS Is Nothing Then
'            objRS = Nothing
'    End If
'    If Not objCDBAccess Is Nothing Then
'        objCDBAccess = Nothing
'    End If
'    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
'********** 2005/05/11 土田 変更 **********↑
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckDB_Torihikisaki
'*【処理概要】
'*?    取引先が有効かを返します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strTorihikisaki ; String 取引先をあらわす文字列
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckDB_Torihikisaki; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*M    2005/05/11; 土田　祐歩(USC); パフォーマンス向上の為､この関数内での切断レコード取得をやめ大域変数の切断レコードを利用する様修正;
'*< Comment End >**************************************************'
Private Function FncCheckDB_Torihikisaki(strTorihikisaki As String, intNowLine As Integer)

    'd.取引先が有効か確認する。
    '　取引先がリテールDB取引先マスタに登録されているか確認する。
    '   　DB：      リテールDB取引先マスタ（m_torihikisaki）
    '   　キー：        取引先コード(torihikisaki_cd)                           CSVファイルから取得
    '           取消日（torikesi_date）                         日付が入っているか
    '　取消日に日付が入っている場合は、取引停止と見なし、エラーとする。　2005/04/19追加

'********** 2005/05/11 土田 削除 **********↓
'DB接続に関する処理を削除
'    Dim strSql            As String
'    Dim objCDBAccess      As Object
'    Dim objRS             As Object
'********** 2005/05/11 土田 削除 **********↑
    
    Dim strMethodCallInfo As String

    On Error GoTo DefectErrorHandler
    
    FncCheckDB_Torihikisaki = False
        
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncCheckDB_Torihikisaki(" & strTorihikisaki & "," & intNowLine & ")"

'********** 2005/05/11 土田 削除 **********↓
'この関数内での切断レコードセット取得処理削除
'    strSql = "SELECT torikesi_date FROM v_torihikisaki WITH(readcommitted) WHERE torihikisaki_cd = '" & strTorihikisaki & "'"
    
'    'CDBAccessインスタンス生成
'    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
'    'SQL実行
'        objRS = objCDBAccess.GetDCRecordset("VSWF", strSql)
'********** 2005/05/11 土田 削除 **********↑

'********** 2005/05/11 土田 追加 **********↓
'大域変数の切断レコードセットを使用しエラーチェック
    With pTChkTorihikisakiRS
'********** 2005/05/17 土田 変更 **********↓
'********** 2005/05/18 土田 変更 **********↓
'        If Not .EOF Then
        .Filter = adFilterNone
       If Not .BOF Then
'********** 2005/05/18 土田 変更 **********↑
            .MoveFirst
            .Find "torihikisaki_cd = '" & strTorihikisaki & "'"
        End If
'********** 2005/05/17 土田 変更 **********↑
    End With
'********** 2005/05/11 土田 追加 **********↑

'********** 2005/05/12 土田 変更 **********↓
    ' 取引先が存在しない場合はエラー
'    If objRS.EOF = True Then
    If pTChkTorihikisakiRS.EOF = True Then
        FncSetError (intNowLine, enmTenColNo.Torihikisaki, PARAM_NOTEXISTDB
'        GoTo Proc_End
        Exit Function
    End If

    '日付チェックは切断レコード取得時に済ませている為ここではチェックしない
'    Dim strTorikesiDate
'    strTorikesiDate = objRS.Fields("torikesi_date")
'    ' 日付が空欄でない場合はエラー
'    If Trim(strTorikesiDate) <> "" Then
'        FncSetError (intNowLine, enmTenColNo.Torihikisaki, PARAM_TORIHIKISTOP
'        GoTo Proc_End
'    End If
'********** 2005/05/12 土田 変更 **********↑
    
    FncCheckDB_Torihikisaki = True
        
'********** 2005/05/11 土田 変更 **********↓
'DB接続に関する処理を削除
'    objRS.Close
'        objRS = Nothing
'    objCDBAccess = Nothing
    
'    'GetObjectContext.SetComplete

Exit Function

'Proc_End:
'    If Not objRS Is Nothing Then
'            objRS = Nothing
'    End If
'    If Not objCDBAccess Is Nothing Then
'        objCDBAccess = Nothing
'    End If
'    Exit Function

DefectErrorHandler:
'Tsuchida start
'    If Not objRS Is Nothing Then
'            objRS = Nothing
'    End If
'    If Not objCDBAccess Is Nothing Then
'        objCDBAccess = Nothing
'    End If
'    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
'********** 2005/05/11 土田 変更 **********↑
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckDB_BaikaHenkouRiyuu
'*【処理概要】
'*?    売価変更理由が有効かを返します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strBaika_HenkouRiyuu ; String 売価変更理由をあらわす文字列
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckDB_BaikaHenkouRiyuu; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*M    2005/05/11; 土田　祐歩(USC); パフォーマンス向上の為､この関数内での切断レコード取得をやめ大域変数の切断レコードを利用する様修正;
'*< Comment End >**************************************************'
Private Function FncCheckDB_BaikaHenkouRiyuu(strBaika_HenkouRiyuu As String, intNowLine As Integer)
        
    'c.売価変更理由が有効か確認する。
    '　売価変更理由がリテールDB名称マスタに登録されているか確認する。
    '   　DB：      リテールDB名称マスタ（v_meisyou）
    '   　キー：        KEY区分（key_kbn）                          「25」
    '           KEYコード（key_cd）                         CSVファイルから取得
    
'********** 2005/05/11 土田 削除 **********↓
'この関数内での切断レコードセット取得処理削除
'    Dim strSql As String
    
'    FncCheckDB_BaikaHenkouRiyuu = True

    

'    'SQL文作成（条件：KEY区分=025、売価変更理由コード）
'    strSql = "SELECT key_cd FROM v_meisyou WITH(readcommitted) "
'    strSql = strSql & "WHERE key_kbn = '" & KEYCD_BAIKARIYUU & "' And "
'    strSql = strSql & "key_cd = '" & strBaika_HenkouRiyuu & "'"
    
'    If FncCheckDB_IsRecordExists("VSWF", strSql) = False Then
'        FncSetError (intNowLine, enmTenColNo.BaikaRiyuu, PARAM_NOTEXISTDB
'        FncCheckDB_BaikaHenkouRiyuu = False
'    End If
'********** 2005/05/11 土田 削除 **********↑
 
'********** 2005/05/11 土田 追加 **********↓
'大域変数の切断レコードセットを使用しエラーチェック
    With pTChkBaikaHenkouRS
'********** 2005/05/17 土田 変更 **********↓
'********** 2005/05/18 土田 変更 **********↓
'        If Not .EOF Then
       .Filter = adFilterNone
        If Not .BOF Then
'********** 2005/05/18 土田 変更 **********↑
            .MoveFirst
            .Find "key_cd = '" & strBaika_HenkouRiyuu & "'"
        End If
'********** 2005/05/17 土田 変更 **********↑
    End With
 
    If pTChkBaikaHenkouRS.EOF Then
        FncSetError (intNowLine, enmTenColNo.BaikaRiyuu, PARAM_NOTEXISTDB
    End If
    
    FncCheckDB_BaikaHenkouRiyuu = Not pTChkBaikaHenkouRS.EOF
'********** 2005/05/11 土田 削除 **********↑

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckDB_CenterClass
'*【処理概要】
'*?    センター区分が有効かを返します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strCenterClass ; String センター区分をあらわす文字列
'*P    strCenterCode ; String センターコードをあらわす文字列
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckDB_CenterClass; String;0:正常 1:センター区分異常 2:センターコード異常
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*M    2005/05/12; 土田　祐歩(USC); パフォーマンス向上の為､この関数内での切断レコード取得をやめ大域変数の切断レコードを利用する様修正;
'*M    2011/02/16; 営業システム/和田; センターコード追加;
'*M    2015/09/15; 流通システムG/和田; TCシステム構築;
'*< Comment End >**************************************************'
Private Function FncCheckDB_CenterClass(strCenterClass As String, _
                                  strCenterCode As String, _
                                  intNowLine As Integer) As String

    'e.センター区分が有効か確認する。
    '　センター区分がWF名称マスタに登録されているか確認する。
    '   　DB：      WF名称マスタ（m_wf_meisyou）
    '   　キー：    KEY区分（key_kbn）                      「M09」
    '           KEYコード（key_cd）                         CSVファイルから取得
    
    'センター区分が"1"の場合、センター区分とセンターコード
    'を連結させたものが、KEYコードになります。
    'つまり、
    '　・センター区分="1"の場合
    '　　KEY区分（key_kbn）：「M09」
    '　　KEYコード（key_cd）：センター区分＋センターコード
    '　・センター区分="1"以外の場合
    '　　KEY区分（key_kbn）：「M09」
    '　　KEYコード（key_cd）：センター区分

'********** 2005/05/12土田 削除 **********↓
'この関数内での切断レコードセット取得処理削除
'    Dim strSql            As String
    
'    strSql = "SELECT key_cd FROM m_wf_meisyou WITH(readcommitted) "
'    strSql = strSql & "WHERE key_kbn = '" & KEYCD_CENTER_KBN & "' "

'    If strCenterClass = "1" Then
'    '　・センター区分="1"の場合、KEYコードは"10151"でチェック
'        strSql = strSql & " AND key_cd = '" & strCenterClass & CENTERCD_KODAMA & "'"
'    Else
'    '　・センター区分="1"以外の場合
'        strSql = strSql & " AND key_cd = '" & strCenterClass & "'"
'    End If
'
'    If FncCheckDB_IsRecordExists("VSWF", strSql) = False Then
'        FncCheckDB_CenterClass = False
'        FncSetError (intNowLine, enmTenColNo.CenterKbn, PARAM_NOTEXISTDB
'    Else
'        FncCheckDB_CenterClass = True
'    End If
'********** 2005/05/12土田 削除 **********↑

'********** 2005/05/12土田 追加 **********↓
'大域変数の切断レコードセットを使用しエラーチェック

    Dim strKeyCode As String 'チェックするKEYコード
    Dim strHanteiKbn As String 'チェック判定区分をセット
    strHanteiKbn = "0"  '正常(初期値をセット)
    
    '　・センター区分="1"の場合、KEYコードは"10151"(センター区分＋センターコード)でチェック
    If strCenterClass = "1" Then
        'strKeyCode = "1" & CENTERCD_KODAMA
        strKeyCode = "1" & strCenterCode
    '2015/09/15 wadak7 ins
    'センター区分="3"の場合、KEYコードは"30653"(センター区分＋センターコード)でチェック
    ElseIf strCenterClass = "3" Then
        strKeyCode = "3" & strCenterCode
    Else
    '　・センター区分="1"以外の場合、KEYコードはセンター区分でチェック
        strKeyCode = strCenterClass
    End If
    
    With pTChkCenterRS
'********** 2005/05/17 土田 変更 **********↓
'********** 2005/05/18 土田 変更 **********↓
'        If Not .EOF Then
       .Filter = adFilterNone
        If Not .BOF Then
'********** 2005/05/18 土田 変更 **********↑
            .MoveFirst
            .Find "key_cd = '" & strKeyCode & " '"
        End If
'********** 2005/05/17 土田 変更 **********↑
    End With
 
    If pTChkCenterRS.EOF Then
        '2015/09/19 wadak7 upd
        'If strCenterClass = "1" Then
        If strCenterClass = "1" Or strCenterClass = "3" Then
            FncSetError (intNowLine, enmModColNo.centerCD, PARAM_NOTEXISTDB
            strHanteiKbn = "2"  'センターコード異常
        Else
            FncSetError (intNowLine, enmTenColNo.CenterKbn, PARAM_NOTEXISTDB
            strHanteiKbn = "1"  'センター区分異常
        End If
    End If
    
    'FncCheckDB_CenterClass = Not pTChkCenterRS.EOF
    FncCheckDB_CenterClass = strHanteiKbn

'********** 2005/05/12土田 追加 **********↑

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckDB_Sakituke
'*【処理概要】
'*?    統合マスタの先付変更情報に同日の変更日情報が無い事を確認する。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strKeyCode ; String 主キーをあらわす文字列
'*P    arrCSVRecord ; String レコードをあらわす文字列
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckDB_Sakituke; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*M    2005/05/12; 土田　祐歩(USC); パフォーマンス向上の為､大域変数の切断レコードを利用し店コードを展開する様修正;
'*M    2005/05/13; 田中　禎人(USC); パフォーマンス向上の為､変更日についても大域変数の切断レコードを利用する様修正;
'*< Comment End >**************************************************'
Private Function FncCheckDB_Sakituke(strKeyCode As String, arrCSVRecord() As String, intNowLine As Integer)

'f.統合マスタの先付変更情報に同日の変更日情報が無いか確認する。
'　統合マスタの先付変更情報（売価変更情報、仕入変更情報、店舗センターコード変更情報）
'　に同日の変更日が入っていない事を確認する。                                           2005/04/19 追加
'   　DB：      統合マスタ 売価変更情報（t_siire_henkou）
'           統合マスタ 仕入変更情報（t_baika_henkou）
'           統合マスタ 店舗センターコード変更情報（t_tenpo_hattyuu_cycle_henkou）
'   　キー：        主キー（key_cd）                            1.(2)b.(b) KEY情報1から取得
'           店コード（ten_cd）                          1.(2)a. 展開している店コード
'           変更日（henkou_date）                           CSVファイルから取得

    Dim strSQL            As String
    Dim strTenCode        As String
    Dim strTenCodeLine    As String
    
    FncCheckDB_Sakituke = True
    
    strTenCode = arrCSVRecord(enmTenColNo.TenCd)
    
'********** 2005/05/12土田 削除 **********↓
'    '展開している店コードを取得するSQL文作成
'    If strTenCode = ALL_TEN_CD Or strTenCode = ALL_JIG_CD Or Len(strTenCode) = 2 Then
'        strTenCodeLine = "SELECT A.jigyousyo_cd AS tencd "
'        strTenCodeLine = strTenCodeLine & "FROM v_jigyousyo AS A WITH(readcommitted) "
'        strTenCodeLine = strTenCodeLine & "INNER JOIN m_wf_meisyou AS B WITH(readcommitted) " & vbCrLf
'        strTenCodeLine = strTenCodeLine & "ON A.jigyou_cd = B.key_cd " & vbCrLf
'        strTenCodeLine = strTenCodeLine & "WHERE A.hanbetu_kbn = '" & HANBETU_KBN_TENPO & "' AND "
'        strTenCodeLine = strTenCodeLine & "(RTrim(A.heiten_date) = '' OR A.heiten_date is null "
'        strTenCodeLine = strTenCodeLine & "OR A.heiten_date >= " & FncGetNowDate() & ") "
'        strTenCodeLine = strTenCodeLine & "AND Substring(B.settei_naiyou,4,1) = '1' " & vbCrLf
'        strTenCodeLine = strTenCodeLine & "AND B.key_kbn = '" & KEYCD_KIKAKU_DOUNYU_TAISYOU & "' " & vbCrLf
'        If strTenCode <> ALL_TEN_CD And strTenCode <> ALL_JIG_CD Then
'            strTenCodeLine = strTenCodeLine & "AND A.jigyou_cd = " & strTenCode
'        End If
'    Else
'        strTenCodeLine = strTenCode
'    End If
'********** 2005/05/12土田 削除 **********↑

'********** 2005/05/12土田 追加 **********↓
    If strTenCode = ALL_TEN_CD Or strTenCode = ALL_JIG_CD Then
        '全店、全事業の場合
        strTenCodeLine = pTTenCodeList
        
'********** 2005/05/20 土田 追加 **********↓
    ElseIf strTenCode = HON_TEN_CD Then
        '本部の場合
        strTenCodeLine = "'" & HON_TEN_CD & "'"
'********** 2005/05/20 土田 追加 **********↑

    ElseIf Len(strTenCode) = 4 Then
        '各店
        strTenCodeLine = strTenCode
    Else
        '各事業
'        With pTChkJigyouRS
'            .MoveFirst
'            .Find "jigyou_cd = '" & strTenCode & "'"
'            If Not .EOF Then
'                strTenCodeLine = .Fields("jigyousyo_cd")
'            End If
'        End With
'
'        strTenCodeLine = ""

            With pTChkJigyouRS
    '            .MoveFirst
                .Filter = "jigyou_cd = '" & strTenCode & "'"
                If Not .EOF Then
                    Do Until .EOF
                        strTenCodeLine = strTenCodeLine & "'" & .Fields("jigyousyo_cd") & "',"
                        .MoveNext
                    Loop
                End If
                .Filter = adFilterNone
            End With
            
            strTenCodeLine = Left(strTenCodeLine, Len(strTenCodeLine) - 1)
   
    End If

'********** 2005/05/12土田 追加 **********↑
    
    ' 売価変更情報に対し問い合わせ
    If arrCSVRecord(enmTenColNo.Baika) <> "" _
        Or arrCSVRecord(enmTenColNo.BaikaRiyuu) <> "" Then
        
'********** 2005/05/13田中 変更 **********↓
'        strSql = "SELECT key_cd FROM v_baika_henkou WITH(readcommitted) "
'        strSql = strSql & "WHERE key_cd = '" & strKeyCode & "' "
'        strSql = strSql & " AND ten_cd IN(" & strTenCodeLine & ") "
'        strSql = strSql & " AND henkou_date = '" & arrCSVRecord(enmTenColNo.HenkouDate) & "'"
'
'        If FncCheckDB_IsRecordExists("VSWF", strSql) = True Then
'            FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_REPDATEBAIKA
'            FncCheckDB_Sakituke = False
'        End If

        With pTChkMstBaikaRS
'********** 2005/05/17 土田 削除 **********↓
'            .MoveFirst
'********** 2005/05/17 土田 削除 **********↑
            .Filter = "key_cd = '" & strKeyCode & "' AND henkou_date = '" & arrCSVRecord(enmTenColNo.HenkouDate) & "'"
            If Not .EOF Then
                Do Until .EOF
                    If InStr(strTenCodeLine, .Fields("ten_cd")) > 0 Then
'********** 2005/06/15 北澤 **********↓
'                        FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_REPDATEBAIKA
'                        FncCheckDB_Sakituke = False
'                        Exit Do
'売価先付変更情報が基幹DBに既存してもエラーにせず、登録区分｢2｣でデータ作成(又は更新)対象とする
                        arrTourokuKbn2(intNowLine - 1, conColBaika) _
                        = arrTourokuKbn2(intNowLine - 1, conColBaika) _
                        & "'" & .Fields("ten_cd") & "',"
'********** 2005/06/15 北澤 **********↑
                    End If
                    .MoveNext
                Loop
'********** 2005/06/15 北澤 **********↓
'登録区分=2の店コード配列の右端のカンマを除去する
                If Len(arrTourokuKbn2(intNowLine - 1, conColBaika)) > 0 Then
                    arrTourokuKbn2(intNowLine - 1, conColBaika) _
                    = Left(arrTourokuKbn2(intNowLine - 1, conColBaika), Len(arrTourokuKbn2(intNowLine - 1, conColBaika)) - 1)
                End If
'********** 2005/06/15 北澤 **********↑
            End If
            .Filter = adFilterNone
        End With
'********** 2005/05/13田中 変更 **********↑

    End If
    
    ' 仕入変更情報に対し問い合わせ

'********** 2005/05/13田中 変更 **********↓
    If arrCSVRecord(enmTenColNo.Torihikisaki) <> "" _
        Or arrCSVRecord(enmTenColNo.Hacchuutani) <> "" _
        Or arrCSVRecord(enmTenColNo.Genka) <> "" Then
'
'        strSql = "SELECT key_cd FROM v_siire_henkou WITH(readcommitted) "
'        strSql = strSql & "WHERE key_cd = '" & strKeyCode & "' "
'        strSql = strSql & " AND ten_cd IN(" & strTenCodeLine & ") "
'        strSql = strSql & " AND henkou_date = '" & arrCSVRecord(enmTenColNo.HenkouDate) & "'"
'
'        If FncCheckDB_IsRecordExists("VSWF", strSql) = True Then
'            FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_REPDATESIIRE
'            FncCheckDB_Sakituke = False
'        End If

        With pTChkMstSiireRS
'********** 2005/05/17 土田 削除 **********↓
'            .MoveFirst
'********** 2005/05/17 土田 削除 **********↑
            .Filter = "key_cd = '" & strKeyCode & "' AND henkou_date = '" & arrCSVRecord(enmTenColNo.HenkouDate) & "'"
            If Not .EOF Then
                Do Until .EOF
                    If InStr(strTenCodeLine, .Fields("ten_cd")) > 0 Then
'********** 2005/06/15 北澤 **********↓
'                        FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_REPDATESIIRE
'                        FncCheckDB_Sakituke = False
'                        Exit Do
'仕入先付変更情報が基幹DBに既存してもエラーにせず、登録区分｢2｣でデータ作成(又は更新)対象とする
                        arrTourokuKbn2(intNowLine - 1, conColSiire) _
                        = arrTourokuKbn2(intNowLine - 1, conColSiire) _
                        & "'" & .Fields("ten_cd") & "',"
'********** 2005/06/15 北澤 **********↑
                    End If
                    .MoveNext
                Loop
'********** 2005/06/15 北澤 **********↓
'登録区分=2の店コード配列の右端のカンマを除去する
                If Len(arrTourokuKbn2(intNowLine - 1, conColSiire)) > 0 Then
                    arrTourokuKbn2(intNowLine - 1, conColSiire) _
                    = Left(arrTourokuKbn2(intNowLine - 1, conColSiire), Len(arrTourokuKbn2(intNowLine - 1, conColSiire)) - 1)
                End If
'********** 2005/06/15 北澤 **********↑
            End If
            .Filter = adFilterNone
        End With
    End If
'********** 2005/05/13田中 変更 **********↑
    
    ' 店舗センターコード変更情報に対し問い合わせ

'********** 2005/05/13田中 変更 **********↓
    If arrCSVRecord(enmTenColNo.CenterKbn) <> "" Then
'        strSql = "SELECT key_cd FROM v_tenpo_center_cd_henkou WITH(readcommitted) "
'        strSql = strSql & "WHERE key_cd = '" & strKeyCode & "' "
'        strSql = strSql & " AND ten_cd IN(" & strTenCodeLine & ") "
'        strSql = strSql & " AND henkou_date = '" & arrCSVRecord(enmTenColNo.HenkouDate) & "'"
'
'        If FncCheckDB_IsRecordExists("VSWF", strSql) = True Then
'            FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_REPDATETENPO
'            FncCheckDB_Sakituke = False
'        End If


        With pTChkMstCenterRS
'********** 2005/05/17 土田 削除 **********↓
'            .MoveFirst
'********** 2005/05/17 土田 削除 **********↑
            .Filter = "key_cd = '" & strKeyCode & "' AND henkou_date = '" & arrCSVRecord(enmTenColNo.HenkouDate) & "'"
            If Not .EOF Then
                Do Until .EOF
                    If InStr(strTenCodeLine, .Fields("ten_cd")) > 0 Then
'********** 2005/06/15 北澤 **********↓
'                        FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_REPDATETENPO
'                        FncCheckDB_Sakituke = False
'                        Exit Do
'センター先付変更情報が基幹DBに既存してもエラーにせず、登録区分｢2｣でデータ作成(又は更新)対象とする
                        arrTourokuKbn2(intNowLine - 1, conColCenter) _
                        = arrTourokuKbn2(intNowLine - 1, conColCenter) _
                        & "'" & .Fields("ten_cd") & "',"
'********** 2005/06/15 北澤 **********↑
                    End If
                    .MoveNext
                Loop
'********** 2005/06/15 北澤 **********↓
'登録区分=2の店コード配列の右端のカンマを除去する
                If Len(arrTourokuKbn2(intNowLine - 1, conColCenter)) > 0 Then
                    arrTourokuKbn2(intNowLine - 1, conColCenter) _
                    = Left(arrTourokuKbn2(intNowLine - 1, conColCenter), Len(arrTourokuKbn2(intNowLine - 1, conColCenter)) - 1)
                End If
'********** 2005/06/15 北澤 **********↑
            End If
            .Filter = adFilterNone
        End With
    End If
'********** 2005/05/13田中 変更 **********↑
    
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckDB_IsRecordExists
'*【処理概要】
'*?    指定されたSQL文を実行し、テーブル内に該当するレコードが存在するかどうかを返します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strUDL_Name ; String 接続文字列名
'*P    strSQLInstruction ; String SQL文
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckDB_IsRecordExists; Boolean;レコードが存在する:True,存在しない:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheckDB_IsRecordExists(strUDL_Name As String, strSQLInstruction As String) As Boolean

    Dim objRs             As Object
    Dim objCDBAccess      As Object
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncCheckDB_IsRecordExists(" & strUDL_Name & "," & strSQLInstruction & ")"
        
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")

    'SQL実行
        objRs = objCDBAccess.GetDCRecordset(strUDL_Name, strSQLInstruction)
    
    ' データが取得できなければ、テーブル内には該当レコードは存在しない
    If objRs.EOF = True Then
        FncCheckDB_IsRecordExists = False
        GoTo Proc_End
    Else
        FncCheckDB_IsRecordExists = True
    End If
    
    objRs.Close
        objRs = Nothing
    objCDBAccess = Nothing
    
    'GetObjectContext.SetComplete
Exit Function
    
Proc_End:
    If Not objRs Is Nothing Then
            objRs = Nothing
    End If
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
    Exit Function

DefectErrorHandler:
    If Not objRs Is Nothing Then
            objRs = Nothing
    End If
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    FncTenTourokuJig

'*【処理概要】
'*?    事業コードに紐付けた店コードのデータを取得して、店コード毎に
'*?    DBに取り込む

'*【パラメータ】
'*P    strErrMsg;        string; エラーメッセージ;
'*P    intN;             int; 何番目のレコードか;
'*P    strTenCd;         string; 事業コード;

'*【戻り値】
'*R    Boolean; TRUE:成功、FALSE:失敗;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/20; 黒沢(ユニシステム株式会社); 新規作成;
'*< Comment End >**************************************************'
Private Function FncTenTourokuJig(ByRef strErrMsg, ByVal intN, ByVal strTenCd) As Boolean
    Dim sql, objCDB, intSelNum, blnRet
    Dim rsSel As ADODB.Recordset
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTourokuJig(" & CStr(intN) + 1 & "行目：【事業コード：" & strTenCd & "】)"
    
    FncTenTourokuJig = False

    '事業コードから店コードを取得する。
        objCDB = CreateObject("COMDBUtil.CDBAccess")    'CDBAccessをインスタンス化

    'SQL文作成（条件：判別区分=5、閉店日がNULLor一括登録日以降、KEY区分=W02、設定内容の４桁目が1）
    sql = "SELECT A.jigyousyo_cd AS tencd " & vbCrLf
    sql = sql & "FROM " & vbCrLf
    sql = sql & " v_jigyousyo AS A WITH(readcommitted) " & vbCrLf
    sql = sql & " INNER JOIN m_wf_meisyou AS B WITH(readcommitted) ON A.jigyou_cd = B.key_cd " & vbCrLf
    sql = sql & "WHERE " & vbCrLf
    sql = sql & " A.hanbetu_kbn = '" & HANBETU_KBN_TENPO & "' AND " & vbCrLf
    sql = sql & " (RTrim(A.heiten_date) = '' OR A.heiten_date is null OR " & vbCrLf
    sql = sql & " A.heiten_date >= '" & FncGetNowDate() & "') AND " & vbCrLf
    sql = sql & " Substring(settei_naiyou,4,1) = '1' AND " & vbCrLf
    sql = sql & " B.key_kbn = '" & KEYCD_KIKAKU_DOUNYU_TAISYOU & "' " & vbCrLf
    '全店、全事業の場合は、以下の条件は必要ない。
    If strTenCd <> ALL_TEN_CD And strTenCd <> ALL_JIG_CD Then
        sql = sql & " AND A.jigyou_cd = '" & strTenCd & "' " & vbCrLf
    End If
    
        rsSel = objCDB.GetDCRecordset("VSWF", sql)
        objCDB = Nothing
    If rsSel.EOF = True Then    '対象の店コードが無い場合はエラー
        strErrMsg = CStr(intN + 1) & "行目：【事業コード：" & strTenCd & "】の店コードは１件もありませんでした。"
        rsSel.Close
            rsSel = Nothing
        Exit Function
    End If
    
    '店ｺｰﾄﾞ毎にデータを商品選定WFのDBに取り込む（FncTenTourokuTen関数を呼ぶ）。
    Do Until rsSel.EOF
        blnRet = FncTenTourokuTen(strErrMsg, intN, rsSel.Fields("tencd"))
        If blnRet = False Then
            rsSel.Close
                rsSel = Nothing
            Exit Function
        End If
        rsSel.MoveNext
    Loop
    
    FncTenTourokuJig = True
    
    'DBクローズ
    rsSel.Close
        rsSel = Nothing
        objCDB = Nothing
    
    'GetObjectContext.SetComplete
Exit Function

DefectErrorHandler:
    If Not rsSel Is Nothing Then
            rsSel = Nothing
    End If
    If Not objCDB Is Nothing Then
            objCDB = Nothing
    End If
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    FncTenTourokuTen

'*【処理概要】
'*?    指定された店コードのデータをDBに取り込む

'*【パラメータ】
'*P    strErrMsg;        string; エラーメッセージ;
'*P    intN;             int; 何番目のレコードか;
'*P    strTenCd;         string; 店コード;

'*【戻り値】
'*R    Boolean; TRUE:成功、FALSE:失敗;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/20; 黒沢(ユニシステム株式会社); 新規作成;
'*M    2018/12/14; 李松涛(大連)              ; 一括登録時、毎回一つユーザーのみ実行できる;
'*< Comment End >**************************************************'
Private Function FncTenTourokuTen(ByRef strErrMsg, ByVal intN, ByVal strTenCd) As Boolean

    Dim sql, objCDB, intUdtNum, intInsNum, intSelNum
    Dim rsSel
    Dim arrParamSel() As Object                             ' SQLパラメーター値
    Dim arrParamIns() As Object                             ' SQLパラメーター値
    Dim arrParamUpd() As Object                             ' SQLパラメーター値
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTourokuTen(" & CStr(intN) + 1 & "行目：【店コード：" & strTenCd & "】)"
    
    FncTenTourokuTen = False

    '取得したデータを元に、DBを更新する。
        objCDB = CreateObject("COMDBUtil.CDBAccess")    'CDBAccessをインスタンス化
    
    '一括登録時、毎回一つユーザーのみ実行できる
    Call LockTableIkkatu(objCDB)

    '売価、売価変更理由のどれかがあるかチェックする。
    If paryCsvData(intN)(enmTenColNo.Baika) <> "" Or paryCsvData(intN)(enmTenColNo.BaikaRiyuu) <> "" Then
       '商品選定DB(売価変更情報)の更新
       'データの存在チェックを行う。
        Call SubSqlSelTen(intN, strTenCd, sql, arrParamSel, TBL_BAI_HEN)
            rsSel = objCDB.GetDCRecordset("VSWF", sql, arrParamSel)
        
       'データが無い場合
        If rsSel.EOF = True Then    'insert処理
            Call SubSqlInsTen(intN, strTenCd, sql, arrParamIns, TBL_BAI_HEN) 'SQL文作成
            intInsNum = objCDB.ExecuteSql("VSWF", sql, arrParamIns) 'SQL実行
            If intInsNum <= 0 Then
                strErrMsg = CStr(intN + 1) & "行目：【店コード：" & strTenCd & "】の"
                strErrMsg = strErrMsg & "売価変更情報を追加するのに失敗しました。"
                    objCDB = Nothing
                    rsSel = Nothing
                Exit Function
            End If
            pflgBaikaToukbn = True  '商品選定DB(WF管理項目)の売価変更の登録区分更新用フラグ
        'データがある場合
        Else                        'Update処理
            Call SubSqlUpdTen(intN, strTenCd, sql, arrParamUpd, TBL_BAI_HEN) 'SQL文作成
            intUdtNum = objCDB.ExecuteSql("VSWF", sql, arrParamUpd) 'SQL実行
            If intUdtNum <= 0 Then
                strErrMsg = CStr(intN + 1) & "行目：【店コード：" & strTenCd & "】の"
                strErrMsg = strErrMsg & "売価変更情報を更新するのに失敗しました。"
                    objCDB = Nothing
                rsSel.Close
                    rsSel = Nothing
                Exit Function
            End If
            rsSel.Close
        End If
            rsSel = Nothing
    End If
        
    '原価、取引先、発注単位のどれかがあるかチェックする。
    If paryCsvData(intN)(enmTenColNo.Genka) <> "" _
        Or paryCsvData(intN)(enmTenColNo.Torihikisaki) <> "" _
        Or paryCsvData(intN)(enmTenColNo.Hacchuutani) <> "" Then
        '商品選定DB(仕入変更情報)の更新
        'データの存在チェックを行う。
        Call SubSqlSelTen(intN, strTenCd, sql, arrParamSel, TBL_SII_HEN)
            rsSel = objCDB.GetDCRecordset("VSWF", sql, arrParamSel)
    
        'データが無い場合
        If rsSel.EOF = True Then   'insert処理
            Call SubSqlInsTen(intN, strTenCd, sql, arrParamIns, TBL_SII_HEN) 'SQL文作成
            intInsNum = objCDB.ExecuteSql("VSWF", sql, arrParamIns) 'SQL実行
            If intInsNum <= 0 Then
                strErrMsg = CStr(intN + 1) & "行目：【店コード：" & strTenCd & "】の"
                strErrMsg = strErrMsg & "仕入変更情報を追加するのに失敗しました。"
                    objCDB = Nothing
                    rsSel = Nothing
                Exit Function
            End If
            
            pflgSiireToukbn = True  '商品選定DB(WF管理項目)の仕入変更の登録区分更新用フラグ
           
        Else                        'Update処理
            Call SubSqlUpdTen(intN, strTenCd, sql, arrParamUpd, TBL_SII_HEN) 'SQL文作成
            intUdtNum = objCDB.ExecuteSql("VSWF", sql, arrParamUpd) 'SQL実行
            If intUdtNum <= 0 Then
                strErrMsg = CStr(intN + 1) & "行目：【店コード：" & strTenCd & "】の"
                strErrMsg = strErrMsg & "仕入変更情報を更新するのに失敗しました。"
                    objCDB = Nothing
                rsSel.Close
                    rsSel = Nothing
                Exit Function
            End If
        End If
            rsSel = Nothing
    End If
       
    'センター区分があるかチェックする。
    If paryCsvData(intN)(enmTenColNo.CenterKbn) <> "" Then
        '商品選定DB(センターコード変更情報)の更新
        'データの存在チェックを行う。
        Call SubSqlSelTen(intN, strTenCd, sql, arrParamSel, TBL_CEN_HEN)
            rsSel = objCDB.GetDCRecordset("VSWF", sql, arrParamSel)
            
        'データが無い場合
        If rsSel.EOF = True Then   'insert処理
            Call SubSqlInsTen(intN, strTenCd, sql, arrParamIns, TBL_CEN_HEN) 'SQL文作成
            intInsNum = objCDB.ExecuteSql("VSWF", sql, arrParamIns) 'SQL実行
            If intInsNum <= 0 Then
                strErrMsg = CStr(intN + 1) & "行目：【店コード：" & strTenCd & "】の"
                strErrMsg = strErrMsg & "センターコード変更情報を追加するのに失敗しました。"
                    objCDB = Nothing
                    rsSel = Nothing
                Exit Function
            End If
            
            pflgCenterToukbn = True '商品選定DB(WF管理項目)のセンターコード変更の登録区分更新用フラグ
            
        Else                        'Update処理
            Call SubSqlUpdTen(intN, strTenCd, sql, arrParamUpd, TBL_CEN_HEN) 'SQL文作成
            intUdtNum = objCDB.ExecuteSql("VSWF", sql, arrParamUpd) 'SQL実行
            If intUdtNum <= 0 Then
                strErrMsg = CStr(intN + 1) & "行目：【店コード：" & strTenCd & "】の"
                strErrMsg = strErrMsg & "センターコード変更情報を更新するのに失敗しました。"
                    objCDB = Nothing
                rsSel.Close
                    rsSel = Nothing
                Exit Function
            End If
            rsSel.Close
        End If
            rsSel = Nothing
    End If
        
    FncTenTourokuTen = True
    
    'DBクローズ
        objCDB = Nothing
    'GetObjectContext.SetComplete
Exit Function

DefectErrorHandler:
    If Not rsSel Is Nothing Then
            rsSel = Nothing
    End If
    If Not objCDB Is Nothing Then
            objCDB = Nothing
    End If
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    FncSetWFKanri

'*【処理概要】
'*?    WF管理項目を最新情報に更新する。（レコード毎）

'*【パラメータ】
'*P    strErrMsg;        string; エラーメッセージ;
'*P    intN;             int; 何番目のレコードか;
'*P    intType;          int; 発注中止or店別情報;

'*【戻り値】
'*R    Boolean; TRUE:成功、FALSE:失敗;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/20; 黒沢(ユニシステム株式会社); 新規作成;
'*M    2005/04/26; 黒沢(ユニシステム株式会社); 関数名変更（FncUpdWFKanri→FncSetWFKanri）;
'*M    2005/04/26; 黒沢(ユニシステム株式会社); 引数追加（intType）;
'*< Comment End >**************************************************'
'********** 2005/04/26 黒沢 変更 **********↓
Private Function FncSetWFKanri(ByRef strErrMsg, ByVal intN, ByVal intType) As Boolean
'Private Function FncUpdWFKanri(ByRef strErrMsg, ByVal intN) As Boolean
'********** 2005/04/26 黒沢 変更 **********↑
    Dim sql, objCDB, intUdtNum
    
'********** 2005/04/26 黒沢 変更 **********↓
    Dim strSelSql, strInsSql, strUpdSql, lngP_Cnt
    Dim arrParam() As Object
'    Dim arrParamUpd() As Object
'********** 2005/04/26 黒沢 変更 **********↑

'********** 2005/04/26 黒沢 変更 **********↓
    FncSetWFKanri = False
'    FncUpdWFKanri = False
'********** 2005/04/26 黒沢 変更 **********↑
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
'********** 2005/04/26 黒沢 変更 **********↓
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncUpdWFKanri(" & CStr(intN) + 1 & "行目：【WF管理項目登録エラー】)"
'    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncUpdWFKanri(" & CStr(intN) + 1 & "行目：【WF管理項目更新エラー】)"
'********** 2005/04/26 黒沢 変更 **********↑
    
        objCDB = CreateObject("COMDBUtil.CDBAccess")    'CDBAccessをインスタンス化
    
'********** 2005/04/26 黒沢 変更 **********↓
    lngP_Cnt = 0
    Erase arrParam
    Call SubSqlSelWFKanri(intN, strSelSql, arrParam, lngP_Cnt, intType)
    Call SubSqlInsWFKanri(intN, strInsSql, arrParam, lngP_Cnt, intType)
    Call SubSqlUpdWFKanri(intN, strUpdSql, arrParam, lngP_Cnt, intType)
    sql = "If (" & strSelSql & ") = 0 " & vbCrLf        'データが既に存在するかチェック
    sql = sql & "BEGIN " & strInsSql & " END " & vbCrLf '存在しない場合は、追加
    sql = sql & "ELSE " & vbCrLf
    sql = sql & "BEGIN " & strUpdSql & " END " & vbCrLf '存在する場合は、更新
    intUdtNum = objCDB.ExecuteSql("VSWF", sql, arrParam)
'    Call SubSqlUpdWFKanri(intN, sql, arrParamUpd)
'    intUdtNum = objCDB.ExecuteSql("VSWF", sql, arrParamUpd) 'SQL実行
'********** 2005/04/26 黒沢 変更 **********↑

    If intUdtNum <= 0 Then
'********** 2005/04/26 黒沢 変更 **********↓
        strErrMsg = CStr(intN + 1) & "行目：WF管理項目のデータを登録するのに失敗しました。"
'        strErrMsg = CStr(intN + 1) & "行目：WF管理項目を更新するのに失敗しました。"
'********** 2005/04/26 黒沢 変更 **********↑
            objCDB = Nothing
        Exit Function
    End If

'********** 2005/04/26 黒沢 変更 **********↓
    FncSetWFKanri = True
'    FncUpdWFKanri = True
'********** 2005/04/26 黒沢 変更 **********↑
    'DBクローズ
        objCDB = Nothing
    'GetObjectContext.SetComplete
Exit Function

DefectErrorHandler:
    If Not objCDB Is Nothing Then
            objCDB = Nothing
    End If
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    SubHacSqlSelJig

'*【処理概要】
'*?    SELECT文作成（該当事業コードの発注中止日データチェック用）

'*【パラメータ】
'*P    intN;            int;     何番目のレコードか;
'*P    strJigCd;        string;  事業コード;
'*P    sql;             string;  SQL文（格納用）;
'*P    arrPrm;          Object; パラメータ（格納用）;
'*P    lngP_Cnt;        Long;    パラメータ数（格納用）;

'*【戻り値】
'*R    なし;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/25; 黒沢(ユニシステム株式会社); 新規作成;
'*< Comment End >**************************************************'
Private Sub SubHacSqlSelJig(ByVal intN, ByVal strJigCD, ByRef sql, ByRef arrPrm As Object, ByRef lngP_Cnt)

    sql = "SELECT COUNT(*) FROM t_syousen_jigyoubetu_settei WITH(UPDLOCK) " & vbCrLf
    sql = sql & "WHERE" & vbCrLf
    sql = sql & " kikaku_syurui_cd = ? AND " & vbCrLf
    sql = sql & " nen_ren_no = ? AND " & vbCrLf
    sql = sql & " eda_no = ? AND " & vbCrLf
    sql = sql & " jan_cd = ? AND " & vbCrLf
    sql = sql & " jigyou_cd = ? " & vbCrLf
    
    'パラメータセット
    ReDim Preserve arrPrm(lngP_Cnt + 4) As Object
    arrPrm(lngP_Cnt + 0) = MakeParam("@KIKAKU_SYURUI", DT_Char, 3, pstrKikakuSyurui)
    arrPrm(lngP_Cnt + 1) = MakeParam("@NEN_REN_NO", DT_Char, 10, pstrKikakuNenrenban)
    arrPrm(lngP_Cnt + 2) = MakeParam("@EDA_NO", DT_Char, 3, pstrKikakuEdaban)
    arrPrm(lngP_Cnt + 3) = MakeParam("@JAN_CD", DT_VarChar, 20, paryCsvData(intN)(enmHacColNo.janCD))
    arrPrm(lngP_Cnt + 4) = MakeParam("@JIGYOU_CD", DT_Char, 2, strJigCD)
    lngP_Cnt = lngP_Cnt + 5

End Sub

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    SubHacSqlInsJig

'*【処理概要】
'*?    INSERT文作成（該当事業コードの発注中止日データが無い場合用）

'*【パラメータ】
'*P    intN;            int;     何番目のレコードか;
'*P    strJigCd;        string;  事業コード;
'*P    sql;             string;  SQL文（格納用）;
'*P    arrPrm;          Object; パラメータ（格納用）;
'*P    lngP_Cnt;        Long;    パラメータ数（格納用）;

'*【戻り値】
'*R    なし;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/25; 黒沢(ユニシステム株式会社); 新規作成;
'*< Comment End >**************************************************'
Private Sub SubHacSqlInsJig(ByVal intN, ByVal strJigCD, ByRef sql, ByRef arrParam As Object, ByRef lngP_Cnt)

'********** 2005/05/09 北澤 **********↓
    On Error GoTo Err_SubHacSqlInsJig
    Dim strSyobunHhouhou As String
    Dim strDaitaiSyouhinCd As String
'********** 2005/05/09 北澤 **********↑
    
    'SQL文作成
    sql = "INSERT INTO t_syousen_jigyoubetu_settei " & vbCrLf
    sql = sql & "(kikaku_syurui_cd, nen_ren_no, eda_no, jan_cd, jigyou_cd," & vbCrLf
    sql = sql & " hattyuu_kaisi_date, hattyuu_syuuryou_date," & vbCrLf
    sql = sql & " hattyuu_tyuusi_riyuu_kbn, syobun_houhou, daitai_syouhin_cd," & vbCrLf
    sql = sql & " dounyuu_baika, jyougen_baika, kagen_baika, " & vbCrLf
    sql = sql & " touroku_user_id, kousin_user_id, touroku_ymd_hms, kousin_ymd_hms) " & vbCrLf
    sql = sql & "SELECT TOP 1 ?,?,?,B.jan_cd,A.jigyou_cd, " & vbCrLf
    sql = sql & "A.hattyuu_kaisi_date, " & vbCrLf   '発注開始日
    sql = sql & "CONVERT(char(8),'" & paryCsvData(intN)(enmHacColNo.SyuryouDate) & "') " & vbCrLf '発注終了日
    sql = sql & ",?,?,?, " & vbCrLf             '発注中止理由区分、処分方法、代替JANコード
    sql = sql & "A.dounyuu_baika, A.jyougen_baika, A.kagen_baika, " & vbCrLf
    sql = sql & "'" & pstrTantousyaId & "', null, GETDATE(), null " & vbCrLf
    sql = sql & "FROM v_jigyou_betu AS A WITH(readcommitted) " & vbCrLf
    sql = sql & "INNER JOIN v_key_1 AS B WITH(readcommitted) ON A.key_cd = B.key_cd " & vbCrLf
    sql = sql & "WHERE B.jan_cd = ? AND " & vbCrLf
    sql = sql & " A.jigyou_cd = ? " & vbCrLf
    
    'パラメータセット
    ReDim Preserve arrParam(lngP_Cnt + 7) As Object
    arrParam(lngP_Cnt + 0) = MakeParam("@KIKAKU_SYURUI", DT_Char, 3, pstrKikakuSyurui)
    arrParam(lngP_Cnt + 1) = MakeParam("@NEN_REN_NO", DT_Char, 10, pstrKikakuNenrenban)
    arrParam(lngP_Cnt + 2) = MakeParam("@EDA_NO", DT_Char, 3, pstrKikakuEdaban)

'********** 2005/05/09 北澤 **********↓
    '発注中止区分(入力必須なのでブランクはあり得ない)
'********** 2005/05/09 北澤 **********↑
    arrParam(lngP_Cnt + 3) = MakeParam("@HACCHU_KBN", DT_Char, 2, paryCsvData(intN)(enmHacColNo.ChushiKbn))

'********** 2005/05/09 北澤 **********↓
'現行ソースコードコメント化
'    arrParam(lngP_Cnt + 4) = MakeParam("@SYOBUN_HOU", DT_Char, 1, paryCsvData(intN)(enmHacColNo.Syobun))
'    arrParam(lngP_Cnt + 5) = MakeParam("@DAITAI_CD", DT_Char, 13, paryCsvData(intN)(enmHacColNo.DaiJanCd))

'EXCELの処分方法・代替JANコードがブランクの場合は基幹DBより値を取得する
    ''EXCELの処分方法・代替JANコードがブランクの場合は基幹DBより値を取得する
    If paryCsvData(intN)(enmHacColNo.Syobun) = "" _
    Or paryCsvData(intN)(enmHacColNo.DaiJanCd) = "" Then
        
        '''変数定義
        Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess") '共通関数(SQLServer接続用)
        Dim rsAdo As ADODB.Recordset            'レコードセット
        Dim strMethodCallInfo As String         'エラー情報
        Dim strSQL As String                    '作業用文字列
        Const strUdl As String = "VSWF"         'SQLServer接続用UDLファイル名
        
        '''エラー情報を作成しておく
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.SubHacSqlInsJig(JAN:" & paryCsvData(intN)(enmHacColNo.janCD) & " 事業:" & strJigCD & ")"
        
        '''SQL文を作成する
        '''(TABLE: 事業別ビュー，主キーマスタ)
        '''(FIELDS: 処分方法，代替JANコード)
        '''(KEY:JANコード, 事業コード)
        strSQL = _
        "SELECT v_jigyou_betu.syobun_houhou, v_jigyou_betu.daitai_syouhin_cd" & _
        "  FROM v_jigyou_betu WITH(READCOMMITTED)" & _
        " INNER JOIN v_key_1 WITH(READCOMMITTED) ON v_jigyou_betu.key_cd = v_key_1.key_cd" & _
        " WHERE v_key_1.jan_cd = '" & paryCsvData(intN)(enmHacColNo.janCD) & "'" & _
        "   AND v_jigyou_betu.jigyou_cd = '" & Trim(strJigCD) & "'"
        
        '''共通関数オブジェクトを作成する
        objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
        
        '''レコードセットを取得する
            rsAdo = objCDBAccess.GetDCRecordset(strUdl, strSQL)
        
        '''処分方法・代替JANコードを変数に保管する
        If rsAdo.EOF = False Then
            If IsNull(rsAdo("syobun_houhou")) = False Then
                strSyobunHhouhou = rsAdo("syobun_houhou")
            End If
            If IsNull(rsAdo("daitai_syouhin_cd")) = False Then
                strDaitaiSyouhinCd = rsAdo("daitai_syouhin_cd")
            End If
        End If
        
        '''データベースを閉じる
            rsAdo = Nothing
        objCDBAccess = Nothing
        'GetObjectContext.SetComplete
    End If
    
    ''処分方法をセットする
        If paryCsvData(intN)(enmHacColNo.Syobun) = "" Then
            If strSyobunHhouhou <> "" Then
                arrParam(lngP_Cnt + 4) = MakeParam("@SYOBUN_HOU", DT_Char, 1, strSyobunHhouhou)
            Else
                arrParam(lngP_Cnt + 4) = MakeParam("@SYOBUN_HOU", DT_Char, 1, "0")
            End If
        Else
            arrParam(lngP_Cnt + 4) = MakeParam("@SYOBUN_HOU", DT_Char, 1, paryCsvData(intN)(enmHacColNo.Syobun))
        End If
        
        '''代替JANコードをセットする
        If paryCsvData(intN)(enmHacColNo.DaiJanCd) = "" Then
            If strDaitaiSyouhinCd <> "" Then
                arrParam(lngP_Cnt + 5) = MakeParam("@DAITAI_CD", DT_Char, 13, strDaitaiSyouhinCd)
            Else
                arrParam(lngP_Cnt + 5) = MakeParam("@DAITAI_CD", DT_Char, 13, Null)
            End If
        Else
            arrParam(lngP_Cnt + 5) = MakeParam("@DAITAI_CD", DT_Char, 13, paryCsvData(intN)(enmHacColNo.DaiJanCd))
        End If
'********** 2005/05/09 北澤 **********↑
    
    arrParam(lngP_Cnt + 6) = MakeParam("@JAN_CD", DT_VarChar, 20, paryCsvData(intN)(enmHacColNo.janCD))
    arrParam(lngP_Cnt + 7) = MakeParam("@JIGYOU_CD", DT_Char, 2, strJigCD)
    lngP_Cnt = lngP_Cnt + 8

'********** 2005/05/09 北澤 **********↓
    'エラー時の処理
Ext_SubHacSqlInsJig:
    Exit Sub
Err_SubHacSqlInsJig:
    If Not rsAdo Is Nothing Then Set rsAdo = Nothing
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    Call ErrRaise(DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description)
    Resume Ext_SubHacSqlInsJig
'********** 2005/05/09 北澤 **********↑

End Sub

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    SubHacSqlUpdJig

'*【処理概要】
'*?    UPDATE文作成（該当事業コードの発注中止日データがある場合用）

'*【パラメータ】
'*P    intN;            int;     何番目のレコードか;
'*P    strJigCd;        string;  事業コード;
'*P    sql;             string;  SQL文（格納用）;
'*P    arrParamUpd;     Object; パラメータ（格納用）;
'*P    lngParam_Cnt;    Long;    パラメータ数（格納用）;

'*【戻り値】
'*R    なし;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/25; 黒沢(ユニシステム株式会社); 新規作成;
'*< Comment End >**************************************************'
Private Sub SubHacSqlUpdJig(ByVal intN, ByVal strJigCD, ByRef sql, ByRef arrPrm As Object, ByRef lngP_Cnt)
    
    sql = "UPDATE t_syousen_jigyoubetu_settei " & vbCrLf
    sql = sql & "SET" & vbCrLf
    sql = sql & " hattyuu_syuuryou_date = " & vbCrLf                '発注終了日
    sql = sql & "CONVERT(char(8),'" & paryCsvData(intN)(enmHacColNo.SyuryouDate) & "'), " & vbCrLf '発注終了日
    
'********** 2005/05/09 北澤 **********↓
    '発注中止区分(入力必須なのでブランクはあり得ない)
'********** 2005/05/09 北澤 **********↑
    If paryCsvData(intN)(enmHacColNo.ChushiKbn) <> "" Then
        sql = sql & " hattyuu_tyuusi_riyuu_kbn = ?," & vbCrLf
        ReDim Preserve arrPrm(lngP_Cnt) As Object
        arrPrm(lngP_Cnt) = MakeParam("@HAC_RIYUU_KBN", DT_Char, 2, paryCsvData(intN)(enmHacColNo.ChushiKbn))
        lngP_Cnt = lngP_Cnt + 1
    End If

'********** 2005/05/09 北澤 **********↓
    '処分方法(EXCEL項目がブランクの場合は何もしない(基幹DBの値を生かす))
'********** 2005/05/09 北澤 **********↑
    If paryCsvData(intN)(enmHacColNo.Syobun) <> "" Then
        sql = sql & " syobun_houhou = ?," & vbCrLf
        ReDim Preserve arrPrm(lngP_Cnt) As Object
        arrPrm(lngP_Cnt) = MakeParam("@SYO_HOU", DT_Char, 1, paryCsvData(intN)(enmHacColNo.Syobun))
        lngP_Cnt = lngP_Cnt + 1
    End If

'********** 2005/05/06 北澤 **********↓
    '代替JANコード(EXCEL項目がブランクの場合は何もしない(基幹DBの値を生かす))
'********** 2005/05/09 北澤 **********↑
    If paryCsvData(intN)(enmHacColNo.DaiJanCd) <> "" Then
        sql = sql & " daitai_syouhin_cd = ?," & vbCrLf
        ReDim Preserve arrPrm(lngP_Cnt) As Object
        arrPrm(lngP_Cnt) = MakeParam("@DAITAI_CD", DT_Char, 13, paryCsvData(intN)(enmHacColNo.DaiJanCd))
        lngP_Cnt = lngP_Cnt + 1
    End If

    sql = sql & " kousin_user_id = '" & pstrTantousyaId & "'," & vbCrLf
    sql = sql & " kousin_ymd_hms = GETDATE() " & vbCrLf
    'WHERE文
    sql = sql & "WHERE " & vbCrLf
    sql = sql & " kikaku_syurui_cd = ? AND " & vbCrLf
    sql = sql & " nen_ren_no = ? AND " & vbCrLf
    sql = sql & " eda_no = ? AND " & vbCrLf
    sql = sql & " jan_cd = ? AND " & vbCrLf
    sql = sql & " jigyou_cd = ? " & vbCrLf
    
    'パラメータセット
    ReDim Preserve arrPrm(lngP_Cnt + 4) As Object
    arrPrm(lngP_Cnt + 0) = MakeParam("@KIKAKU_SYURUI", DT_Char, 3, pstrKikakuSyurui)
    arrPrm(lngP_Cnt + 1) = MakeParam("@NEN_REN_NO", DT_Char, 10, pstrKikakuNenrenban)
    arrPrm(lngP_Cnt + 2) = MakeParam("@EDA_NO", DT_Char, 3, pstrKikakuEdaban)
    arrPrm(lngP_Cnt + 3) = MakeParam("@JAN_CD", DT_Char, 13, paryCsvData(intN)(enmHacColNo.janCD))
    arrPrm(lngP_Cnt + 4) = MakeParam("@JIGYOU_CD", DT_Char, 2, strJigCD)
    lngP_Cnt = lngP_Cnt + 5

End Sub

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    SubSqlSelTen

'*【処理概要】
'*?    SELECT文作成（店別情報）

'*【パラメータ】
'*P    intN;            int;     何番目のレコードか;
'*P    strTenCd;        string;  店コード;
'*P    sql;             string;  SQL文（格納用）;
'*P    arrParamSel;     Object; パラメータ（格納用）;
'*P    strTbl;          string;  更新テーブル名;

'*【戻り値】
'*R    なし;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/20; 黒沢(ユニシステム株式会社); 新規作成;
'*< Comment End >**************************************************'
Private Sub SubSqlSelTen(ByVal intN, ByVal strTenCd, ByRef sql, ByRef arrParamSel As Object, ByVal strTbl)
    sql = "SELECT * FROM " & strTbl & " WITH(readcommitted) " & vbCrLf
    sql = sql & "WHERE " & vbCrLf
    sql = sql & " kikaku_syurui_cd = ? AND " & vbCrLf
    sql = sql & " nen_ren_no = ? AND " & vbCrLf
    sql = sql & " eda_no = ? AND " & vbCrLf
    sql = sql & " jan_cd = ? AND " & vbCrLf
    sql = sql & " ten_cd = ? AND " & vbCrLf
    sql = sql & " henkou_date = ? " & vbCrLf
    arrParamSel = Array(MakeParam("@KIKAKU_SYURUI", DT_Char, 3, pstrKikakuSyurui), _
                     MakeParam("@NEN_REN_NO", DT_Char, 10, pstrKikakuNenrenban), _
                     MakeParam("@EDA_NO", DT_Char, 3, pstrKikakuEdaban), _
                     MakeParam("@JAN_CD", DT_Char, 13, paryCsvData(intN)(enmTenColNo.janCD)), _
                     MakeParam("@TEN_CD", DT_Char, 4, strTenCd), _
                     MakeParam("@HENKOU_DATE", DT_Char, 8, paryCsvData(intN)(enmTenColNo.HenkouDate)))

End Sub

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    SubSqlInsTen

'*【処理概要】
'*?    INSERT文作成（店別情報）

'*【パラメータ】
'*P    intN;            int;     何番目のレコードか;
'*P    strTenCd;        string;  店コード;
'*P    sql;             string;  SQL文（格納用）;
'*P    arrParamIns;     Object; パラメータ（格納用）;
'*P    strTbl;          string;  更新テーブル名;

'*【戻り値】
'*R    なし;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/20; 黒沢(ユニシステム株式会社); 新規作成;
'*< Comment End >**************************************************'
Private Sub SubSqlInsTen(ByVal intN, ByVal strTenCd, ByRef sql, ByRef arrParamIns As Object, ByVal strTbl)
    'SQL文作成
    sql = "INSERT INTO " & strTbl & " WITH(UPDLOCK) " & vbCrLf
    sql = sql & "(kikaku_syurui_cd, nen_ren_no, eda_no, jan_cd, ten_cd, henkou_date, " & vbCrLf
    
    If strTbl = TBL_BAI_HEN Then
        sql = sql & "baika, baika_henkou_riyuu_cd, " & vbCrLf
    ElseIf strTbl = TBL_SII_HEN Then
        sql = sql & "torihikisaki_cd, hattyuu_tani, genka, " & vbCrLf
    ElseIf strTbl = TBL_CEN_HEN Then
        sql = sql & "center_kbn, center_cd, " & vbCrLf
    End If
    
    sql = sql & "touroku_kbn, touroku_user_id, kousin_user_id, " & vbCrLf
    sql = sql & "touroku_ymd_hms, kousin_ymd_hms) " & vbCrLf
    sql = sql & "VALUES (?,?,?,?,?,?," & vbCrLf
    
    If strTbl = TBL_BAI_HEN Then        '売価変更情報
        If paryCsvData(intN)(enmTenColNo.Baika) <> "" Then
            sql = sql & "'" & CLng(paryCsvData(intN)(enmTenColNo.Baika)) & "'," & vbCrLf
        Else
            sql = sql & "null," & vbCrLf
        End If

'********** 2005/05/06 北澤 **********↓
'現行ソースコードコメント化
'        sql = sql & "'" & paryCsvData(intN)(enmTenColNo.BaikaRiyuu) & "'," & vbCrLf
'EXCEL項目がブランクの場合はテーブルにはNULLをセットする
        If IsNull(paryCsvData(intN)(enmTenColNo.BaikaRiyuu)) _
        Or paryCsvData(intN)(enmTenColNo.BaikaRiyuu) = "" Then
            sql = sql & "Null," & vbCrLf
        Else
            sql = sql & "'" & paryCsvData(intN)(enmTenColNo.BaikaRiyuu) & "'," & vbCrLf
        End If
'********** 2005/05/06 北澤 **********↑
    
    ElseIf strTbl = TBL_SII_HEN Then    '仕入変更情報

'********** 2005/05/06 北澤 **********↓
'現行ソースコードコメント化
'        sql = sql & "'" & paryCsvData(intN)(enmTenColNo.Torihikisaki) & "'," & vbCrLf
'EXCEL項目がブランクの場合はテーブルにはNULLをセットする
        If IsNull(paryCsvData(intN)(enmTenColNo.Torihikisaki)) _
        Or paryCsvData(intN)(enmTenColNo.Torihikisaki) = "" Then
            sql = sql & "Null," & vbCrLf
        Else
            sql = sql & "'" & paryCsvData(intN)(enmTenColNo.Torihikisaki) & "'," & vbCrLf
        End If
'********** 2005/05/06 北澤 **********↑
        
        If paryCsvData(intN)(enmTenColNo.Hacchuutani) <> "" Then
            sql = sql & "'" & CLng(paryCsvData(intN)(enmTenColNo.Hacchuutani)) & "'," & vbCrLf
        Else
            sql = sql & "null," & vbCrLf
        End If
        If paryCsvData(intN)(enmTenColNo.Genka) <> "" Then
            sql = sql & "'" & CDbl(paryCsvData(intN)(enmTenColNo.Genka)) & "'," & vbCrLf
        Else
            sql = sql & "null," & vbCrLf
        End If
    ElseIf strTbl = TBL_CEN_HEN Then    'センターコード変更情報

'********** 2005/05/06 北澤 **********↓
'現行ソースコードコメント化
'        sql = sql & "'" & paryCsvData(intN)(enmTenColNo.CenterKbn) & "'," & vbCrLf
'EXCEL項目がブランクの場合はテーブルにはNULLをセットする
        If IsNull(paryCsvData(intN)(enmTenColNo.CenterKbn)) _
        Or paryCsvData(intN)(enmTenColNo.CenterKbn) = "" Then
            sql = sql & "Null," & vbCrLf
        Else
            sql = sql & "'" & paryCsvData(intN)(enmTenColNo.CenterKbn) & "'," & vbCrLf
        End If
'********** 2005/05/06 北澤 **********↑
        
        'センター区分が0の場合はNULL、1の場合は'0151'、2の場合は事業所マスタから取得
'********** 2005/05/06 北澤 **********↓
'現行ソースコードコメント化
'        sql = sql & "'" & FncGetCenterCd(paryCsvData(intN)(enmTenColNo.CenterKbn), strTenCd) & "'," & vbCrLf  'センターコード
'EXCEL項目がブランクの場合はテーブルにはNULLをセットする
        Dim strWorkString As String
        strWorkString = FncGetCenterCd(paryCsvData(intN)(enmTenColNo.CenterKbn), strTenCd)
        If strWorkString = "" Then
            sql = sql & "Null," & vbCrLf
        Else
            sql = sql & "'" & strWorkString & "'," & vbCrLf
        End If
'********** 2005/05/06 北澤 **********↑
    
    End If
    
    sql = sql & "'" & WF_KANRI_TOUROKU_KBN_SINKI & "'," & vbCrLf
    sql = sql & "'" & pstrTantousyaId & "'," & vbCrLf
    sql = sql & "null," & vbCrLf
    sql = sql & "'" & Now & "'," & vbCrLf
    sql = sql & "null) " & vbCrLf
    
    arrParamIns = Array(MakeParam("@KIKAKU_SYURUI", DT_Char, 3, pstrKikakuSyurui), _
                     MakeParam("@NEN_REN_NO", DT_Char, 10, pstrKikakuNenrenban), _
                     MakeParam("@EDA_NO", DT_Char, 3, pstrKikakuEdaban), _
                     MakeParam("@JAN_CD", DT_Char, 13, paryCsvData(intN)(enmTenColNo.janCD)), _
                     MakeParam("@TEN_CD", DT_Char, 4, strTenCd), _
                     MakeParam("@HENKOU_DATE", DT_Char, 8, paryCsvData(intN)(enmTenColNo.HenkouDate)))

End Sub

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    SubSqlUpdTen

'*【処理概要】
'*?    UPDATE文作成（店別情報）

'*【パラメータ】
'*P    intN;            int;     何番目のレコードか;
'*P    strTenCd;        string;  店コード;
'*P    sql;             string;  SQL文（格納用）;
'*P    arrParamUpd;     Object; パラメータ（格納用）;
'*P    strTbl;          string;  更新テーブル名;

'*【戻り値】
'*R    なし;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/20; 黒沢(ユニシステム株式会社); 新規作成;
'*< Comment End >**************************************************'
Private Sub SubSqlUpdTen(ByVal intN, ByVal strTenCd, ByRef sql, ByRef arrParamUpd As Object, ByVal strTbl)
    'SQL文作成
    sql = "UPDATE " & strTbl & " WITH(UPDLOCK) " & vbCrLf
    sql = sql & " SET " & vbCrLf
    
    If strTbl = TBL_BAI_HEN Then    '売価変更情報（売価、売価変更理由）
        If paryCsvData(intN)(enmTenColNo.Baika) <> "" Then
            sql = sql & "baika = '" & CLng(paryCsvData(intN)(enmTenColNo.Baika)) & "', " & vbCrLf
        End If
        If paryCsvData(intN)(enmTenColNo.BaikaRiyuu) <> "" Then
            sql = sql & "baika_henkou_riyuu_cd = '" & paryCsvData(intN)(enmTenColNo.BaikaRiyuu) & "', " & vbCrLf
        End If
    ElseIf strTbl = TBL_SII_HEN Then    '仕入変更情報（取引先、発注単位、原価）
        If paryCsvData(intN)(enmTenColNo.Torihikisaki) <> "" Then
            sql = sql & "torihikisaki_cd = '" & paryCsvData(intN)(enmTenColNo.Torihikisaki) & "', " & vbCrLf
        End If
        If paryCsvData(intN)(enmTenColNo.Hacchuutani) <> "" Then
            sql = sql & "hattyuu_tani = '" & CLng(paryCsvData(intN)(enmTenColNo.Hacchuutani)) & "', " & vbCrLf
        End If
        If paryCsvData(intN)(enmTenColNo.Genka) <> "" Then
            sql = sql & "genka = '" & CDbl(paryCsvData(intN)(enmTenColNo.Genka)) & "', " & vbCrLf
        End If
    ElseIf strTbl = TBL_CEN_HEN Then    'センターコード変更情報（センター区分、センターコード）
        If paryCsvData(intN)(enmTenColNo.CenterKbn) <> "" Then
            sql = sql & "center_kbn = '" & paryCsvData(intN)(enmTenColNo.CenterKbn) & "', " & vbCrLf
            'センター区分が0の場合はNULL、1の場合は'0151'、2の場合は事業所マスタから取得
'********** 2005/05/06 北澤 **********↓
'現行ソースコードコメント化
'            sql = sql & "center_cd = '" & FncGetCenterCd(paryCsvData(intN)(enmTenColNo.CenterKbn), strTenCd) & "', " & vbCrLf
'EXCEL項目がブランクの場合はテーブルにはNULLをセットする
            Dim strWorkString As String
            strWorkString = FncGetCenterCd(paryCsvData(intN)(enmTenColNo.CenterKbn), strTenCd)
            If strWorkString = "" Then
                sql = sql & "center_cd = dbNull.value," & vbCrLf
            Else
                sql = sql & "center_cd = '" & strWorkString & "'," & vbCrLf
            End If
'********** 2005/05/06 北澤 **********↑
        End If
    End If
        
'    sql = sql & "touroku_kbn = '" & WF_KANRI_TOUROKU_KBN_SINKI & "', " & vbCrLf    '登録区分は更新しない
    sql = sql & "kousin_user_id = '" & pstrTantousyaId & "', " & vbCrLf
    sql = sql & "kousin_ymd_hms = '" & Now & "' " & vbCrLf
    'WHERE文
    sql = sql & "WHERE " & vbCrLf
    sql = sql & " kikaku_syurui_cd = ? AND " & vbCrLf
    sql = sql & " nen_ren_no = ? AND " & vbCrLf
    sql = sql & " eda_no = ? AND " & vbCrLf
    sql = sql & " jan_cd = ? AND " & vbCrLf
    sql = sql & " ten_cd = ? AND " & vbCrLf
    sql = sql & " henkou_date = ? " & vbCrLf
    
    arrParamUpd = Array(MakeParam("@KIKAKU_SYURUI", DT_Char, 3, pstrKikakuSyurui), _
                     MakeParam("@NEN_REN_NO", DT_Char, 10, pstrKikakuNenrenban), _
                     MakeParam("@EDA_NO", DT_Char, 3, pstrKikakuEdaban), _
                     MakeParam("@JAN_CD", DT_Char, 13, paryCsvData(intN)(enmTenColNo.janCD)), _
                     MakeParam("@TEN_CD", DT_Char, 4, strTenCd), _
                     MakeParam("@HENKOU_DATE", DT_Char, 8, paryCsvData(intN)(enmTenColNo.HenkouDate)))

End Sub

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    SubSqlSelWFKanri

'*【処理概要】
'*?    SELECT文作成（WF管理項目データチェック用）

'*【パラメータ】
'*P    intN;            int;     何番目のレコードか;
'*P    sql;             string;  SQL文（格納用）;
'*P    arrPrm;          Object; パラメータ（格納用）;
'*P    lngP_Cnt;        Long;    パラメータ数（格納用）;
'*P    intType;         int;     発注中止or店別情報;

'*【戻り値】
'*R    なし;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/26; 黒沢(ユニシステム株式会社); 新規作成;
'*< Comment End >**************************************************'
Private Sub SubSqlSelWFKanri(ByVal intN, ByRef sql, ByRef arrPrm As Object, ByRef lngP_Cnt, ByVal intType)

    sql = "SELECT COUNT(*) FROM t_syousen_wf_kanri WITH(UPDLOCK) " & vbCrLf
    sql = sql & "WHERE" & vbCrLf
    sql = sql & " kikaku_syurui_cd = ? AND " & vbCrLf
    sql = sql & " nen_ren_no = ? AND " & vbCrLf
    sql = sql & " eda_no = ? AND " & vbCrLf
    sql = sql & " jan_cd = ? " & vbCrLf
    
    'パラメータセット
    ReDim Preserve arrPrm(lngP_Cnt + 3) As Object
    arrPrm(lngP_Cnt + 0) = MakeParam("@KIKAKU_SYURUI", DT_Char, 3, pstrKikakuSyurui)
    arrPrm(lngP_Cnt + 1) = MakeParam("@NEN_REN_NO", DT_Char, 10, pstrKikakuNenrenban)
    arrPrm(lngP_Cnt + 2) = MakeParam("@EDA_NO", DT_Char, 3, pstrKikakuEdaban)
    If intType = RET_FLD_HACH Then  '発注中止
        arrPrm(lngP_Cnt + 3) = MakeParam("@JAN_CD", DT_VarChar, 20, paryCsvData(intN)(enmHacColNo.janCD))
    ElseIf intType = RET_FLD_TENB Then  '店別情報
        arrPrm(lngP_Cnt + 3) = MakeParam("@JAN_CD", DT_VarChar, 20, paryCsvData(intN)(enmTenColNo.janCD))
    ElseIf intType = RET_FLD_SYO Then
        arrPrm(lngP_Cnt + 3) = MakeParam("@JAN_CD", DT_VarChar, 20, paryCsvData(intN)(enmSyoColNo.janCD))
    ElseIf intType = RET_FLD_BUN Then
        arrPrm(lngP_Cnt + 3) = MakeParam("@JAN_CD", DT_VarChar, 20, paryCsvData(intN)(enmBunColNo.janCD))
    End If
    lngP_Cnt = lngP_Cnt + 4

End Sub

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    SubSqlInsWFKanri

'*【処理概要】
'*?    INSERT文作成（WF管理項目追加用）

'*【パラメータ】
'*P    intN;            int;     何番目のレコードか;
'*P    sql;             string;  SQL文（格納用）;
'*P    arrPrm;          Object; パラメータ（格納用）;
'*P    lngP_Cnt;        Long;    パラメータ数（格納用）;
'*P    intType;         int;     発注中止or店別情報;

'*【戻り値】
'*R    なし;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/26; 黒沢(ユニシステム株式会社); 新規作成;
'*< Comment End >**************************************************'
Private Sub SubSqlInsWFKanri(ByVal intN, ByRef sql, ByRef arrPrm As Object, ByRef lngP_Cnt, ByVal intType)
    
    'SQL文作成
'********** 2005/04/27 北澤 **********↓
'現行ソースコードコメント化
'    sql = "INSERT INTO t_syousen_wf_kanri " & vbCrLf
'    sql = sql & "(kikaku_syurui_cd, nen_ren_no, eda_no, jan_cd, jan_umu_kbn," & vbCrLf
'    sql = sql & " touroku_kbn, jyogai_kbn, touroku_tantousya," & vbCrLf
'    sql = sql & " keikakusuu, kikaku_no, kikaku_kahi_kbn, kikaku_syouninsya_comment," & vbCrLf
'    sql = sql & " hinsyu_kahi_kbn, hinsyu_syouninsya_comment, naibu_cd," & vbCrLf
'    sql = sql & " hyouji_siyou_touroku_kbn, sekisai_size_touroku_kbn, pop_touroku_kbn," & vbCrLf
'    sql = sql & " syanai_kihon_touroku_kbn, kowake_touroku_kbn, jigyoubetu_touroku_kbn," & vbCrLf
'    sql = sql & " hanbai_touroku_kbn, siire_touroku_kbn, hinsyu_taikei_touroku_kbn," & vbCrLf
'    sql = sql & " uriba_touroku_kbn, baika_henkou_touroku_kbn, siire_henkou_touroku_kbn," & vbCrLf
'    sql = sql & " center_cd_henkou_touroku_kbn, kowake_henkou_touroku_kbn, hanki_henkou_touroku_kbn," & vbCrLf
'    sql = sql & " touroku_user_id, kousin_user_id, touroku_ymd_hms, kousin_ymd_hms) " & vbCrLf
'    sql = sql & "SELECT ?,?,?,jan_cd,null," & vbCrLf
'    sql = sql & " '" & WF_KANRI_TOUROKU_KBN_SYUUSEI & "','0'," & vbCrLf '登録区分=2、除外区分=0
'    sql = sql & " '" & pstrTantousyaId & "','0',?," & vbCrLf '登録担当者、計画数=0、企画No.
'    sql = sql & " null,null,null,null,key_cd,null,null,null,null,null," & vbCrLf  '内部コード
'    If pflgJigJyoToukbn = True Then '事業別情報登録区分
'        sql = sql & "'1'," & vbCrLf
'    Else
'        sql = sql & "null," & vbCrLf
'    End If
'    sql = sql & " null,null,null,null," & vbCrLf
'    If pflgBaikaToukbn = True Then '売価変更情報登録区分
'        sql = sql & "'1'," & vbCrLf
'    Else
'        sql = sql & "null," & vbCrLf
'    End If
'    If pflgSiireToukbn = True Then '仕入変更情報登録区分
'        sql = sql & "'1'," & vbCrLf
'    Else
'        sql = sql & "null," & vbCrLf
'    End If
'    If pflgCenterToukbn = True Then 'ｾﾝﾀｰｺｰﾄﾞ変更情報登録区分
'        sql = sql & "'1'," & vbCrLf
'    Else
'        sql = sql & "null," & vbCrLf
'    End If
'    sql = sql & " null,null,'" & pstrTantousyaId & "',null," & vbCrLf   '登録UserID
'    sql = sql & " '" & Now & "',null " & vbCrLf   '登録年月日
'    sql = sql & "FROM v_key_1 WITH(readcommitted) " & vbCrLf
'    sql = sql & "WHERE jan_cd = ? " & vbCrLf

'フィールドにセットする値を一部修正
    sql = "INSERT INTO t_syousen_wf_kanri"
    sql = sql & "( kikaku_syurui_cd"             '企画種類コード
    sql = sql & ", nen_ren_no"                   '年連番
    sql = sql & ", eda_no"                       '枝番
    sql = sql & ", jan_cd"                       'JANｺｰﾄﾞ
    sql = sql & ", jan_umu_kbn"                  'JAN有無区分(0:なし 1:あり)
    sql = sql & ", touroku_kbn"                  '登録区分(1:新規 2:修正 9:抹消)
    sql = sql & ", jyogai_kbn"                   '除外区分(0:紐付あり 1:除外)
    sql = sql & ", touroku_tantousya"            '登録担当者
    sql = sql & ", keikakusuu"                   '計画数
    sql = sql & ", kikaku_no"                    '企画No.(企画種類+年連番+枝番)
    sql = sql & ", kikaku_kahi_kbn"              '企画可否区分(0:未承認 1:承認 9:否認)
    sql = sql & ", kikaku_syouninsya_comment"    '企画承認者ｺﾒﾝﾄ
    sql = sql & ", hinsyu_kahi_kbn"              '品種可否区分(0:未承認 1:承認 9:否認)
    sql = sql & ", hinsyu_syouninsya_comment"    '品種承認者ｺﾒﾝﾄ
    sql = sql & ", naibu_cd"                     '内部ｺｰﾄﾞ(統合マスタの主キー)
    sql = sql & ", hyouji_siyou_touroku_kbn"     '表示/仕様情報登録区分(0:未登録 1:登録済)
    sql = sql & ", sekisai_size_touroku_kbn"     '積載/ｻｲｽﾞ情報登録区分(0:未登録 1:登録済)
    sql = sql & ", pop_touroku_kbn"              'POP情報登録区分(0:未登録 1:登録済)
    sql = sql & ", syanai_kihon_touroku_kbn"     '社内基本情報登録区分(0:未登録 1:登録済)
    sql = sql & ", kowake_touroku_kbn"           '小分情報登録区分(0:未登録 1:登録済)
    sql = sql & ", jigyoubetu_touroku_kbn"       '事業別情報登録区分(0:未登録 1:登録済)
    sql = sql & ", hanbai_touroku_kbn"           '販売情報登録区分(0:未登録 1:登録済)
    sql = sql & ", siire_touroku_kbn"            '仕入情報登録区分(0:未登録 1:登録済)
    sql = sql & ", hinsyu_taikei_touroku_kbn"    '品種体系情報登録区分(0:未登録 1:登録済)
    sql = sql & ", uriba_touroku_kbn"            '売場体系情報登録区分(0:未登録 1:登録済)
    sql = sql & ", baika_henkou_touroku_kbn"     '売価変更情報登録区分(0:未登録 1:登録済)
    sql = sql & ", siire_henkou_touroku_kbn"     '仕入変更情報登録区分(0:未登録 1:登録済)
    sql = sql & ", center_cd_henkou_touroku_kbn" 'ｾﾝﾀｰｺｰﾄﾞ変更情報登録区分(0:未登録 1:登録済)
    sql = sql & ", kowake_henkou_touroku_kbn"    '小分変更情報登録区分(0:未登録 1:登録済)
    sql = sql & ", hanki_henkou_touroku_kbn"     '半期変更情報登録区分(0:未登録 1:登録済)
    sql = sql & ", touroku_user_id"              '登録UserID
    sql = sql & ", kousin_user_id"               '更新UserID
    sql = sql & ", touroku_ymd_hms"              '登録年月日時分秒
    sql = sql & ", kousin_ymd_hms"               '更新年月日時分秒
    
    sql = sql & ") SELECT"
    sql = sql & "  ?"                                                '企画種類コード
    sql = sql & ", ?"                                                '年連番
    sql = sql & ", ?"                                                '枝番
    sql = sql & ", jan_cd"                                           'JANｺｰﾄﾞ
    sql = sql & ", '1'"                                              'JAN有無区分(0:なし 1:あり)
    sql = sql & ", '" & WF_KANRI_TOUROKU_KBN_SYUUSEI & "'"           '登録区分(1:新規 2:修正 9:抹消)
    sql = sql & ", '0'"                                              '除外区分(0:紐付あり 1:除外)
    sql = sql & ", '" & pstrTantousyaId & "'"                        '登録担当者
    sql = sql & ", '0'"                                              '計画数
    sql = sql & ", ?"                                                '企画No.(企画種類+年連番+枝番)
    sql = sql & ", '0'"                                              '企画可否区分(0:未承認 1:承認 9:否認)
    sql = sql & ", null"                                             '企画承認者ｺﾒﾝﾄ
    sql = sql & ", '0'"                                              '品種可否区分(0:未承認 1:承認 9:否認)
    sql = sql & ", null"                                             '品種承認者ｺﾒﾝﾄ
    sql = sql & ", key_cd"                                           '内部ｺｰﾄﾞ(統合マスタの主キー)
    sql = sql & ", '1'"                                              '表示/仕様情報登録区分(0:未登録 1:登録済)
    sql = sql & ", '0'"                                              '積載/ｻｲｽﾞ情報登録区分(0:未登録 1:登録済)
    sql = sql & ", '0'"                                              'POP情報登録区分(0:未登録 1:登録済)
    sql = sql & ", '0'"                                              '社内基本情報登録区分(0:未登録 1:登録済)
    sql = sql & ", '0'"                                              '小分情報登録区分(0:未登録 1:登録済)
    sql = sql & ", '" & IIf(pflgJigJyoToukbn = True, "1", "0") & "'" '事業別情報登録区分(0:未登録 1:登録済)
    sql = sql & ", '0'"                                              '販売情報登録区分(0:未登録 1:登録済)
    sql = sql & ", '0'"                                              '仕入情報登録区分(0:未登録 1:登録済)
    sql = sql & ", '0'"                                              '品種体系情報登録区分(0:未登録 1:登録済)
    sql = sql & ", '0'"                                              '売場体系情報登録区分(0:未登録 1:登録済)
    sql = sql & ", '" & IIf(pflgBaikaToukbn = True, "1", "0") & "'"  '売価変更情報登録区分(0:未登録 1:登録済)
    sql = sql & ", '" & IIf(pflgSiireToukbn = True, "1", "0") & "'"  '仕入変更情報登録区分(0:未登録 1:登録済)
    sql = sql & ", '" & IIf(pflgCenterToukbn = True, "1", "0") & "'" 'ｾﾝﾀｰｺｰﾄﾞ変更情報登録区分(0:未登録 1:登録済)
    sql = sql & ", '0'"                                              '小分変更情報登録区分(0:未登録 1:登録済)
    sql = sql & ", '0'"                                              '半期変更情報登録区分(0:未登録 1:登録済)
    sql = sql & ", '" & pstrTantousyaId & "'"                        '登録UserID
    sql = sql & ", null"                                             '更新UserID
    sql = sql & ", '" & Format(Now, "yyyy/mm/dd hh:nn:ss") & "'"     '登録年月日時分秒
    sql = sql & ", null"                                             '更新年月日時分秒
    
    sql = sql & "  FROM v_key_1 WITH(readcommitted)"
    sql = sql & " WHERE jan_cd = ?"
'********** 2005/04/27 北澤 **********↑
    
    'パラメータセット
    ReDim Preserve arrPrm(lngP_Cnt + 4) As Object
    arrPrm(lngP_Cnt + 0) = MakeParam("@KIKAKU_SYURUI", DT_Char, 3, pstrKikakuSyurui)
    arrPrm(lngP_Cnt + 1) = MakeParam("@NEN_REN_NO", DT_Char, 10, pstrKikakuNenrenban)
    arrPrm(lngP_Cnt + 2) = MakeParam("@EDA_NO", DT_Char, 3, pstrKikakuEdaban)
    arrPrm(lngP_Cnt + 3) = MakeParam("@KIKAKU_NO", DT_Char, 17, pstrKikakuSyurui & pstrKikakuNenrenban & "-" & pstrKikakuEdaban)
    If intType = RET_FLD_HACH Then
        arrPrm(lngP_Cnt + 4) = MakeParam("@JAN_CD", DT_VarChar, 20, paryCsvData(intN)(enmHacColNo.janCD))
    ElseIf intType = RET_FLD_TENB Then
        arrPrm(lngP_Cnt + 4) = MakeParam("@JAN_CD", DT_VarChar, 20, paryCsvData(intN)(enmTenColNo.janCD))
    ElseIf intType = RET_FLD_SYO Then
        arrPrm(lngP_Cnt + 4) = MakeParam("@JAN_CD", DT_VarChar, 20, paryCsvData(intN)(enmSyoColNo.janCD))
    ElseIf intType = RET_FLD_BUN Then
        arrPrm(lngP_Cnt + 4) = MakeParam("@JAN_CD", DT_VarChar, 20, paryCsvData(intN)(enmBunColNo.janCD))
    End If
    lngP_Cnt = lngP_Cnt + 5

End Sub

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    SubSqlUpdWFKanri

'*【処理概要】
'*?    UPDATE文作成（WF管理項目テーブル更新用）

'*【パラメータ】
'*P    intN;            int;     何番目のレコードか;
'*P    sql;             string;  SQL文（格納用）;
'*P    arrPrm;          Object; パラメータ（格納用）;
'*P    lngP_Cnt;        Long;    パラメータ数（格納用）;
'*P    intType;         int;     発注中止or店別情報;

'*【戻り値】
'*R    なし;

'******************************************************************'
'*【変更履歴】
'*M    2005/04/20; 黒沢(ユニシステム株式会社); 新規作成;
'*< Comment End >**************************************************'
'********** 2005/04/26 黒沢 変更 **********↓
Private Sub SubSqlUpdWFKanri(ByVal intN, ByRef sql, ByRef arrPrm As Object, ByRef lngP_Cnt, ByVal intType)
'Private Sub SubSqlUpdWFKanri(ByVal intN, ByRef sql, ByRef arrParamUpd As Object)
'********** 2005/04/26 黒沢 変更 **********↑
    'SQL文作成
    sql = "UPDATE t_syousen_wf_kanri WITH(UPDLOCK) " & vbCrLf
    sql = sql & "SET " & vbCrLf
    sql = sql & "touroku_kbn = '" & WF_KANRI_TOUROKU_KBN_SYUUSEI & "', " & vbCrLf  '登録区分は、2:修正
    sql = sql & "touroku_tantousya = '" & pstrTantousyaId & "', " & vbCrLf     '登録担当者は、UserID
    sql = sql & "kikaku_no = '" & pstrKikakuSyurui & pstrKikakuNenrenban & "-" & pstrKikakuEdaban & "', " & vbCrLf    '企画No.
    sql = sql & "naibu_cd = A.key_cd, " & vbCrLf    '主キー
    
'********** 2005/04/25 黒沢 追加 **********↓
    If pflgJigJyoToukbn = True Then     '事業別情報登録区分
        sql = sql & "jigyoubetu_touroku_kbn = '1', " & vbCrLf
    End If
'********** 2005/04/25 黒沢 追加 **********↑
    If pflgBaikaToukbn = True Then      '売価変更情報登録区分
        sql = sql & "baika_henkou_touroku_kbn = '1', " & vbCrLf
    End If
    If pflgSiireToukbn = True Then      '仕入変更情報登録区分
        sql = sql & "siire_henkou_touroku_kbn = '1', " & vbCrLf
    End If
    If pflgCenterToukbn = True Then     'センターコード変更情報登録区分
        sql = sql & "center_cd_henkou_touroku_kbn = '1', " & vbCrLf
    End If
    
'********** 2005/04/27 北澤 **********↓
'表示仕様登録区分を立てる
    sql = sql & "hyouji_siyou_touroku_kbn = '1', " & vbCrLf
'********** 2005/04/27 北澤 **********↑
    
    sql = sql & "kousin_user_id = '" & pstrTantousyaId & "', " & vbCrLf
    sql = sql & "kousin_ymd_hms = '" & Now & "' " & vbCrLf
    
    '主キー用のテーブルの宣言
    sql = sql & "FROM (SELECT key_cd FROM v_key_1 WITH(readcommitted) WHERE jan_cd = ?) AS A " & vbCrLf    '主キー
    'WHERE文
    sql = sql & "WHERE " & vbCrLf
    sql = sql & " kikaku_syurui_cd = ? AND " & vbCrLf
    sql = sql & " nen_ren_no = ? AND " & vbCrLf
    sql = sql & " eda_no = ? AND " & vbCrLf
    sql = sql & " jan_cd = ? " & vbCrLf
    
'********** 2005/04/26 黒沢 変更 **********↓
    'パラメータセット
    ReDim Preserve arrPrm(lngP_Cnt + 4) As Object
    arrPrm(lngP_Cnt + 1) = MakeParam("@KIKAKU_SYURUI", DT_Char, 3, pstrKikakuSyurui)
    arrPrm(lngP_Cnt + 2) = MakeParam("@NEN_REN_NO", DT_Char, 10, pstrKikakuNenrenban)
    arrPrm(lngP_Cnt + 3) = MakeParam("@EDA_NO", DT_Char, 3, pstrKikakuEdaban)
    If intType = RET_FLD_HACH Then
        arrPrm(lngP_Cnt + 0) = MakeParam("@JAN_CD", DT_Char, 13, paryCsvData(intN)(enmHacColNo.janCD))
        arrPrm(lngP_Cnt + 4) = MakeParam("@JAN_CD", DT_Char, 13, paryCsvData(intN)(enmHacColNo.janCD))
    ElseIf intType = RET_FLD_TENB Then
        arrPrm(lngP_Cnt + 0) = MakeParam("@JAN_CD", DT_Char, 13, paryCsvData(intN)(enmTenColNo.janCD))
        arrPrm(lngP_Cnt + 4) = MakeParam("@JAN_CD", DT_Char, 13, paryCsvData(intN)(enmTenColNo.janCD))
    ElseIf intType = RET_FLD_SYO Then
        arrPrm(lngP_Cnt + 0) = MakeParam("@JAN_CD", DT_Char, 13, paryCsvData(intN)(enmSyoColNo.janCD))
        arrPrm(lngP_Cnt + 4) = MakeParam("@JAN_CD", DT_Char, 13, paryCsvData(intN)(enmSyoColNo.janCD))
    ElseIf intType = RET_FLD_BUN Then
        arrPrm(lngP_Cnt + 0) = MakeParam("@JAN_CD", DT_Char, 13, paryCsvData(intN)(enmBunColNo.janCD))
        arrPrm(lngP_Cnt + 4) = MakeParam("@JAN_CD", DT_Char, 13, paryCsvData(intN)(enmBunColNo.janCD))
    End If
    lngP_Cnt = lngP_Cnt + 5
    
'    arrParamUpd = Array(MakeParam("@JAN_CD", DT_Char, 13, paryCsvData(intN)(enmTenColNo.JanCd)), _
'                     MakeParam("@KIKAKU_SYURUI", DT_Char, 3, pstrKikakuSyurui), _
'                     MakeParam("@NEN_REN_NO", DT_Char, 10, pstrKikakuNenrenban), _
'                     MakeParam("@EDA_NO", DT_Char, 3, pstrKikakuEdaban), _
'                     MakeParam("@JAN_CD", DT_Char, 13, paryCsvData(intN)(enmTenColNo.JanCd)))
'********** 2005/04/26 黒沢 変更 **********↑

End Sub

'*< UNISYSTEM >****************************************************'

'*【プロシージャ名】
'*!    FncGetCenterCd

'*【処理概要】
'*?    センター区分を元にして、センターコードを取得する。

'*【パラメータ】
'*P    strCenterKbn; string; センター区分;
'*P    strTenCd;     string; 店コード（センター区分が2の場合に使用）;

'*【戻り値】
'*R    String;　センターコード

'******************************************************************'
'*【変更履歴】
'*M    2005/04/21; 黒沢(ユニシステム株式会社); 新規作成;
'*< Comment End >**************************************************'
Private Function FncGetCenterCd(ByVal strCenterKbn, ByVal strTenCd) As String
    Dim sql, objCDB
    Dim rsSel As ADODB.Recordset

    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncGetCenterCd(センター区分：" & strCenterKbn & ",店コード：" & strTenCd & ")"
    
    Select Case strCenterKbn
    
        Case "0":       'センター区分が0の場合、NULLをセット
            FncGetCenterCd = ""
        Case "1":       'センター区分が1の場合、0151(児玉センター)をセット
            FncGetCenterCd = CENTERCD_KODAMA
        Case "2":       'センター区分が2の場合、店コードを元に事業所マスタから取得したセンターコードをセット
                objCDB = CreateObject("COMDBUtil.CDBAccess")    'CDBAccessをインスタンス化
            sql = "SELECT center_cd FROM v_jigyousyo WITH(readcommitted) " & vbCrLf
            sql = sql & "WHERE jigyousyo_cd = '" & strTenCd & "' " & vbCrLf
                rsSel = objCDB.GetDCRecordset("VSWF", sql)

            'データが無かったらNULLをセットする。
            If rsSel.EOF Then
                FncGetCenterCd = ""
            Else
                FncGetCenterCd = rsSel.Fields("center_cd")
            End If

                objCDB = Nothing
            rsSel.Close
                rsSel = Nothing

        Case Else:
            FncGetCenterCd = ""
            
    End Select
    
    'GetObjectContext.SetComplete

    
Exit Function

DefectErrorHandler:
    If Not rsSel Is Nothing Then
            rsSel = Nothing
    End If
    If Not objCDB Is Nothing Then
            objCDB = Nothing
    End If
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncSetError
'*【処理概要】
'*?    発生したエラー情報を設定します。エラー情報構造体配列の最後部に
'*?    エラー情報を挿入します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intErrorLine ; エラーが発生した行
'*P    intErrorField ; エラーが発生したフィールド
'*P    intErrorType ; 発生したエラー種別
'*P    strErrorParameter ; エラーメッセージ補助用文字列
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncSetError(intErrorLine As Integer, intErrorField As Integer, intErrorType As Integer, Optional strErrorParameter As String = "")
        
    ReDim Preserve typErrorArray(intErrorCount)
    typErrorArray(intErrorCount).intErrorLine = intErrorLine
    typErrorArray(intErrorCount).intErrorField = intErrorField
    typErrorArray(intErrorCount).intErrorType = intErrorType
    typErrorArray(intErrorCount).strErrorParameter = strErrorParameter
    
    intErrorCount = intErrorCount + 1

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncSetErrorPDC
'*【処理概要】
'*?    発生したエラー情報を設定します。エラー情報構造体配列の最後部に
'*?    エラー情報を挿入します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intErrorLine ; エラーが発生した行
'*P    intErrorField ; エラーが発生したフィールド
'*P    intErrorType ; 発生したエラー種別
'*P    strErrorParameter ; エラーメッセージ補助用文字列
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2012/08/10; 楊双(大連); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncSetErrorPDC(intErrorLine As Integer, intErrorField As Integer, intErrorType As Integer, Optional strErrorParameter As String = "")
        
    ReDim Preserve typErrorArrayPDC(intErrorCountPDC)
    typErrorArrayPDC(intErrorCountPDC).intErrorLine = intErrorLine
    typErrorArrayPDC(intErrorCountPDC).intErrorField = intErrorField
    typErrorArrayPDC(intErrorCountPDC).intErrorType = intErrorType
    typErrorArrayPDC(intErrorCountPDC).strErrorParameter = strErrorParameter
    
    intErrorCountPDC = intErrorCountPDC + 1
    
    Dim iLoopCount As Integer

    For iLoopCount = 0 To intErrorCountPDC - 2

        If intErrorField = typErrorArrayPDC(iLoopCount).intErrorField And _
        intErrorType = typErrorArrayPDC(iLoopCount).intErrorType And _
        strErrorParameter = typErrorArrayPDC(iLoopCount).strErrorParameter Then
        
            intErrorCountPDC = intErrorCountPDC - 1
            Exit For
        End If
    
    Next

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncSetErrorMsg_Hac
'*【処理概要】
'*?    発生したエラーすべてのエラーメッセージを取得します。
'*?    (発注中止日版）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【戻り値】
'*    < Name Type Comment >
'*R    FncGetErrorMsg; String;発生したエラーすべてのエラーメッセージ
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/25; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsg_Hac() As String

    Dim iLoopCount As Integer
    Dim intErrorLine As Integer
    Dim intErrorField As Integer
    Dim intErrorType As Integer
    Dim strErrorParameter As String
    
    For iLoopCount = 0 To intErrorCount - 1

        intErrorLine = typErrorArray(iLoopCount).intErrorLine
        intErrorField = typErrorArray(iLoopCount).intErrorField
        intErrorType = typErrorArray(iLoopCount).intErrorType
        strErrorParameter = typErrorArray(iLoopCount).strErrorParameter
        
        FncGetErrorMsg_Hac = FncGetErrorMsg_Hac & _
                        FncGetErrorMsgLine(intErrorLine) & _
                        "【" & FncGetErrorMsgField_Hac(intErrorField) & "】 " & _
                        FncGetErrorMsgType_Hac(intErrorField, intErrorType, strErrorParameter) & _
                        "</div>" & vbCrLf & "<div align='center' class='error'>"
    Next

End Function
'******************************************************************'
'*【プロシージャ名】
'*!    FncGetErrorMsg_Bun
'*【処理概要】
'*?    発生したエラーすべてのエラーメッセージを取得します。
'*?    (分類情報版）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【戻り値】
'*    < Name Type Comment >
'*R    FncSetErrorMsg_Bun; String;発生したエラーすべてのエラーメッセージ
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/09/01; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsg_Bun() As String

    Dim iLoopCount As Integer
    Dim intErrorLine As Integer
    Dim intErrorField As Integer
    Dim intErrorType As Integer
    Dim strErrorParameter As String
    
    For iLoopCount = 0 To intErrorCount - 1

        intErrorLine = typErrorArray(iLoopCount).intErrorLine
        intErrorField = typErrorArray(iLoopCount).intErrorField
        intErrorType = typErrorArray(iLoopCount).intErrorType
        strErrorParameter = typErrorArray(iLoopCount).strErrorParameter
        
        FncGetErrorMsg_Bun = FncGetErrorMsg_Bun & _
                        FncGetErrorMsgLine(intErrorLine) & _
                        "【" & FncGetErrorMsgField_Bun(intErrorField) & "】 " & _
                        FncGetErrorMsgType_Bun(intErrorField, intErrorType, strErrorParameter) & _
                        "</div>" & vbCrLf & "<div align='center' class='error'>"
    Next

End Function
'******************************************************************'
'*【プロシージャ名】
'*!    FncGetErrorMsg_Syo
'*【処理概要】
'*?    発生したエラーすべてのエラーメッセージを取得します。
'*?    (分類情報版）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【戻り値】
'*    < Name Type Comment >
'*R    FncGetErrorMsg_Syo; String;発生したエラーすべてのエラーメッセージ
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/09/01; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsg_Syo() As String

    Dim iLoopCount As Integer
    Dim intErrorLine As Integer
    Dim intErrorField As Integer
    Dim intErrorType As Integer
    Dim strErrorParameter As String
    
    For iLoopCount = 0 To intErrorCount - 1

        intErrorLine = typErrorArray(iLoopCount).intErrorLine
        intErrorField = typErrorArray(iLoopCount).intErrorField
        intErrorType = typErrorArray(iLoopCount).intErrorType
        strErrorParameter = typErrorArray(iLoopCount).strErrorParameter
        
        FncGetErrorMsg_Syo = FncGetErrorMsg_Syo & _
                        FncGetErrorMsgLine(intErrorLine) & _
                        "【" & FncGetErrorMsgField_Syo(intErrorField) & "】 " & _
                        FncGetErrorMsgType_Syo(intErrorField, intErrorType, strErrorParameter) & _
                        "</div>" & vbCrLf & "<div align='center' class='error'>"
    Next

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncSetError
'*【処理概要】
'*?    発生したエラーすべてのエラーメッセージを取得します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【戻り値】
'*    < Name Type Comment >
'*R    FncGetErrorMsg; String;発生したエラーすべてのエラーメッセージ
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsg() As String

    Dim iLoopCount As Integer
    Dim intErrorLine As Integer
    Dim intErrorField As Integer
    Dim intErrorType As Integer
    Dim strErrorParameter As String
    
    For iLoopCount = 0 To intErrorCount - 1

        intErrorLine = typErrorArray(iLoopCount).intErrorLine
        intErrorField = typErrorArray(iLoopCount).intErrorField
        intErrorType = typErrorArray(iLoopCount).intErrorType
        strErrorParameter = typErrorArray(iLoopCount).strErrorParameter
        
        FncGetErrorMsg = FncGetErrorMsg & _
                        FncGetErrorMsgLine(intErrorLine) & _
                        "【" & FncGetErrorMsgField(intErrorField) & "】 " & _
                        FncGetErrorMsgType(intErrorField, intErrorType, strErrorParameter) & _
                        "</div>" & vbCrLf & "<div align='center' class='error'>"
    Next

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncGetErrorMsg_Modosi
'*【処理概要】
'*?    発生したエラーすべてのエラーメッセージを取得します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【戻り値】
'*    < Name Type Comment >
'*R    FncGetErrorMsg_Modosi; String;発生したエラーすべてのエラーメッセージ
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2009/06/06; Supattra T.(Tostem Thai); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsg_Modosi() As String

    Dim iLoopCount As Integer
    Dim intErrorLine As Integer
    Dim intErrorField As Integer
    Dim intErrorType As Integer
    Dim strErrorParameter As String
    
    For iLoopCount = 0 To intErrorCount - 1

        intErrorLine = typErrorArray(iLoopCount).intErrorLine
        intErrorField = typErrorArray(iLoopCount).intErrorField
        intErrorType = typErrorArray(iLoopCount).intErrorType
        strErrorParameter = typErrorArray(iLoopCount).strErrorParameter
        
        FncGetErrorMsg_Modosi = FncGetErrorMsg_Modosi & _
                        FncGetErrorMsgLine(intErrorLine) & _
                        "【" & FncGetErrorMsgField_Modosi(intErrorField) & "】 " & _
                        FncGetErrorMsgType(intErrorField, intErrorType, strErrorParameter) & _
                        "</div>" & vbCrLf & "<div align='center' class='error'>"
    Next

End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncGetErrorMsg_PDC
'*【処理概要】
'*?    発生したエラーすべてのエラーメッセージを取得します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【戻り値】
'*    < Name Type Comment >
'*R    FncGetErrorMsg_PDC; String;発生したエラーすべてのエラーメッセージ
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2012/08/17; 楊双(大連); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsg_PDC() As String

    Dim iLoopCount As Integer
    Dim intErrorLine As Integer
    Dim intErrorField As Integer
    Dim intErrorType As Integer
    Dim strErrorParameter As String
    
    For iLoopCount = 0 To intErrorCountPDC - 1

        intErrorLine = typErrorArrayPDC(iLoopCount).intErrorLine
        intErrorField = typErrorArrayPDC(iLoopCount).intErrorField
        intErrorType = typErrorArrayPDC(iLoopCount).intErrorType
        strErrorParameter = typErrorArrayPDC(iLoopCount).strErrorParameter
        
        FncGetErrorMsg_PDC = FncGetErrorMsg_PDC & _
                        FncGetErrorMsgType(intErrorField, intErrorType, strErrorParameter) & _
                        "</div>" & vbCrLf & "<div align='center' class='error'>"
    Next

End Function

' '*< UNISYSTEM >**************************************************'
'*【プロシージャ名】
'*!    FncGetErrorMsgLine
'*【処理概要】
'*?    エラーが発生した行に対応した文字列を生成します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intErrorLine ; エラーが発生した行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncGetErrorMsgLine; String;エラーが発生した行の文字列表現
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgLine(intErrorLine As Integer) As String

    FncGetErrorMsgLine = intErrorLine & "行目："

End Function

' '*< UNISYSTEM >**************************************************'
'*【プロシージャ名】
'*!    FncGetErrorMsgField_Hac
'*【処理概要】
'*?    エラーが発生したフィールド名を取得します。
'*?    (発注中止日版）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intErrorField ; エラーが発生したフィールド
'*【戻り値】
'*    < Name Type Comment >
'*R    FncGetErrorMsgField; String;エラーが発生したフィールド名
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/25; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgField_Hac(intErrorField As Integer) As String
    Select Case intErrorField
        
        Case enmHacColNo.jigCD:         FncGetErrorMsgField_Hac = "該当事業コード"
        Case enmHacColNo.SyuryouDate:   FncGetErrorMsgField_Hac = "発注終了日"
        Case enmHacColNo.janCD:         FncGetErrorMsgField_Hac = "JANコード"
        Case enmHacColNo.Syobun:        FncGetErrorMsgField_Hac = "処分方法"
        Case enmHacColNo.ChushiKbn:     FncGetErrorMsgField_Hac = "発注中止理由区分"
        Case enmHacColNo.DaiJanCd:      FncGetErrorMsgField_Hac = "代替商品コード"
               
    End Select
End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncGetErrorMsgField_Bun
'*【処理概要】
'*?    エラーが発生したフィールド名を取得します。
'*?    (分類情報版）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intErrorField ; エラーが発生したフィールド
'*【戻り値】
'*    < Name Type Comment >
'*R    FncGetErrorMsgField_Bun; String;エラーが発生したフィールド名
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/09/01; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgField_Bun(intErrorField As Integer) As String
    Select Case intErrorField
        
        Case enmBunColNo.HenkouDate:         FncGetErrorMsgField_Bun = "変更日"
        Case enmBunColNo.janCD:              FncGetErrorMsgField_Bun = "JANコード"
        Case enmBunColNo.Bunrui1:            FncGetErrorMsgField_Bun = "第1分類コード"
        Case enmBunColNo.Bunrui2:            FncGetErrorMsgField_Bun = "第2分類コード"
        Case enmBunColNo.Bunrui3:            FncGetErrorMsgField_Bun = "第3分類コード"
    
    End Select
End Function
'******************************************************************'
'*【プロシージャ名】
'*!    FncGetErrorMsgField_Syo
'*【処理概要】
'*?    エラーが発生したフィールド名を取得します。
'*?    (基本商品版）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intErrorField ; エラーが発生したフィールド
'*【戻り値】
'*    < Name Type Comment >
'*R    FncGetErrorMsgField_Syo; String;エラーが発生したフィールド名
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/09/01; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgField_Syo(intErrorField As Integer) As String
    Select Case intErrorField
        
        Case enmSyoColNo.janCD:                         FncGetErrorMsgField_Syo = "JANコード"
        Case enmSyoColNo.SyouhinmeiKana:                FncGetErrorMsgField_Syo = "カナ商品名(TA伝票用)"
        Case enmSyoColNo.SyouhinmeiKanji:               FncGetErrorMsgField_Syo = "漢字商品名(SA/ﾌﾟﾗｲｽｶｰﾄﾞ用)"
        Case enmSyoColNo.ReceiptmeiKana:                FncGetErrorMsgField_Syo = "レシート名(ｶﾅ)"
        Case enmSyoColNo.ReceiptmeiKanji:               FncGetErrorMsgField_Syo = "レシート名(漢字)"
        Case enmSyoColNo.SyouhihinmeiHansoku:           FncGetErrorMsgField_Syo = "商品名(販促用)"
        
        Case enmSyoColNo.MakerMei:                      FncGetErrorMsgField_Syo = "メーカー名"
        Case enmSyoColNo.KikakuSize:                    FncGetErrorMsgField_Syo = "規格/サイズ"
        Case enmSyoColNo.HinbanKataban:                 FncGetErrorMsgField_Syo = "品番/型番"
        Case enmSyoColNo.IroGara:                       FncGetErrorMsgField_Syo = "色/柄"
        Case enmSyoColNo.Taste:                         FncGetErrorMsgField_Syo = "テースト"
        Case enmSyoColNo.Seisantimei:                   FncGetErrorMsgField_Syo = "生産地名"
        Case enmSyoColNo.Tokutyou:                      FncGetErrorMsgField_Syo = "特徴"
        Case enmSyoColNo.SyouhinCopy:                   FncGetErrorMsgField_Syo = "商品コピー"
        Case enmSyoColNo.TuujyouKakaku:                 FncGetErrorMsgField_Syo = "通常価格"
        Case enmSyoColNo.MakerKakaku:                   FncGetErrorMsgField_Syo = "メーカー希望小売価格"
        Case enmSyoColNo.POPLayout1:                    FncGetErrorMsgField_Syo = "POPレイアウト１"
        Case enmSyoColNo.POPLayout2:                    FncGetErrorMsgField_Syo = "POPレイアウト２"
        Case enmSyoColNo.PopSize:                       FncGetErrorMsgField_Syo = "POPサイズ"
        
        Case enmSyoColNo.AtukaiPtn:                     FncGetErrorMsgField_Syo = "扱いパターン"
        Case enmSyoColNo.HyoujyunBaika:                 FncGetErrorMsgField_Syo = "標準売価"
        Case enmSyoColNo.SikiriGenka:                   FncGetErrorMsgField_Syo = "仕切原価"
        Case enmSyoColNo.VhJyuyouCd:                    FncGetErrorMsgField_Syo = "VH需要コード"
        Case enmSyoColNo.RdJyuyouCd:                    FncGetErrorMsgField_Syo = "RD需要コード"
        Case enmSyoColNo.VcsJyuyouCd:                   FncGetErrorMsgField_Syo = "VCS需要コード"
        Case enmSyoColNo.SvhJyuyouCd:                   FncGetErrorMsgField_Syo = "SVH需要コード"
    
    End Select
End Function

' '*< UNISYSTEM >**************************************************'
'*【プロシージャ名】
'*!    FncGetErrorMsgField
'*【処理概要】
'*?    エラーが発生したフィールド名を取得します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intErrorField ; エラーが発生したフィールド
'*【戻り値】
'*    < Name Type Comment >
'*R    FncGetErrorMsgField; String;エラーが発生したフィールド名
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgField(intErrorField As Integer) As String
    Select Case intErrorField
        Case enmTenColNo.TenCd:         FncGetErrorMsgField = "該当店コード"
        Case enmTenColNo.HenkouDate:    FncGetErrorMsgField = "変更日"
        Case enmTenColNo.janCD:         FncGetErrorMsgField = "JANコード"
        Case enmTenColNo.Genka:         FncGetErrorMsgField = "原価"
        Case enmTenColNo.Baika:         FncGetErrorMsgField = "売価"
        Case enmTenColNo.BaikaRiyuu:    FncGetErrorMsgField = "売価変更理由"
        Case enmTenColNo.Torihikisaki:  FncGetErrorMsgField = "取引先"
        Case enmTenColNo.Hacchuutani:   FncGetErrorMsgField = "発注単位"
        Case enmTenColNo.CenterKbn:     FncGetErrorMsgField = "センター区分"
    End Select
End Function

' '*< UNISYSTEM >**************************************************'
'*【プロシージャ名】
'*!    FncGetErrorMsgField
'*【処理概要】
'*?    エラーが発生したフィールド名を取得します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intErrorField ; エラーが発生したフィールド
'*【戻り値】
'*    < Name Type Comment >
'*R    FncGetErrorMsgField; String;エラーが発生したフィールド名
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2009/06/06; Supattra T.(Tostem Thai); 新規作成;
'*M    2011/02/10; 営業システム/和田;         センターコード追加;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgField_Modosi(intErrorField As Integer) As String
    Select Case intErrorField
        Case enmModColNo.TenCd:         FncGetErrorMsgField_Modosi = "該当店コード"
        Case enmModColNo.HenkouDate:    FncGetErrorMsgField_Modosi = "変更日"
        Case enmModColNo.janCD:         FncGetErrorMsgField_Modosi = "JANコード"
        Case enmModColNo.Genka:         FncGetErrorMsgField_Modosi = "原価"
        Case enmModColNo.Baika:         FncGetErrorMsgField_Modosi = "売価"
        Case enmModColNo.BaikaRiyuu:    FncGetErrorMsgField_Modosi = "売価変更理由"
        Case enmModColNo.Torihikisaki:  FncGetErrorMsgField_Modosi = "取引先"
        Case enmModColNo.Hacchuutani:   FncGetErrorMsgField_Modosi = "発注単位"
        Case enmModColNo.CenterKbn:     FncGetErrorMsgField_Modosi = "センター区分"
        Case enmModColNo.centerCD:      FncGetErrorMsgField_Modosi = "センターコード"
        Case enmModColNo.MoGenka:       FncGetErrorMsgField_Modosi = "戻し原価"
        Case enmModColNo.MoBaika:       FncGetErrorMsgField_Modosi = "戻し売価"
        Case enmModColNo.HattyuuTani:   FncGetErrorMsgField_Modosi = "戻し発注単位"
        Case enmModColNo.PopSize:       FncGetErrorMsgField_Modosi = "POPサイズ"
    End Select
End Function


' '*< UNISYSTEM >**************************************************'
'*【プロシージャ名】
'*!    FncGetErrorMsgType_Hac
'*【処理概要】
'*?    発生したエラーの内容を文字列で取得します。
'*?    (発注中止日版）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intErrorField ; エラーが発生したフィールド
'*P    intErrorType ; 発生したエラーの種別
'*P    strErrorParameter ; エラーメッセージ補助用文字列
'*【戻り値】
'*    < Name Type Comment >
'*R    FncGetErrorMsgField; String;発生したエラーの内容(フィールドごとに異なることもある)
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/25; 浅野　正徳(USC); 新規作成;
'*M    2006/09/13; 和田(IC);        商品選定ＷＦ問合せ撲滅改善;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgType_Hac(intErrorField As Integer, _
                              intErrorType As Integer, _
                              Optional strErrorParameter As String = "")
     
    Select Case intErrorType
        Case PARAM_EMPTY:           FncGetErrorMsgType_Hac = "必須項目です。"
        Case PARAM_NOTDISIT:        FncGetErrorMsgType_Hac = "半角数字で入力して下さい。"
        Case PARAM_NONARROW:        FncGetErrorMsgType_Hac = "半角文字で入力して下さい。"
        Case PARAM_NOTFINDMASTER:
            Select Case intErrorField
                Case enmHacColNo.jigCD
                    FncGetErrorMsgType_Hac = "マスタに登録されていない事業コードが設定されています。"
                Case enmHacColNo.ChushiKbn
                    FncGetErrorMsgType_Hac = "マスタに登録されていない発注中止理由区分が設定されています。"
            End Select
        Case PARAM_WRONGCHECKDIGIT: FncGetErrorMsgType_Hac = "チェックデジットを確認して下さい。"
        Case PARAM_WRONGDATE:       FncGetErrorMsgType_Hac = "日付を正しく入力して下さい。YYYYMMDD"
        Case PARAM_SHORTDISIT:
            Select Case intErrorField
                Case enmHacColNo.janCD:
                    FncGetErrorMsgType_Hac = "桁数に過不足があります。8桁もしくは13桁の数字を入力して下さい。"
                Case enmHacColNo.DaiJanCd:
                    FncGetErrorMsgType_Hac = "桁数に過不足があります。8桁もしくは13桁の数字を入力して下さい。"
            End Select
        Case PARAM_NOTEXISTDB:
            Select Case intErrorField
                Case enmHacColNo.jigCD
                    FncGetErrorMsgType_Hac = "マスタに登録されていない事業コードが設定されています。"
                Case enmHacColNo.janCD
                    FncGetErrorMsgType_Hac = "JANコードは統合DBに登録されていません。"
                Case enmHacColNo.ChushiKbn
                    FncGetErrorMsgType_Hac = "マスタに登録されていない発注中止理由区分が設定されています。"
                Case enmHacColNo.DaiJanCd
                    FncGetErrorMsgType_Hac = "代替商品コードは統合DBに登録されていません。"
                Case enmHacColNo.SyuryouDate
                    FncGetErrorMsgType_Hac = "発注開始日が統合DBに登録されていません。"
            End Select
        Case PARAM_SAMECODEDB:      FncGetErrorMsgType_Hac = "JANコードは企画№" & strErrorParameter & "に登録されています。"
        Case PARAM_SAMEKEYDB:       FncGetErrorMsgType_Hac = "内部コードは企画№" & strErrorParameter & "に登録されています。"
        Case PARAM_NOTDEFINED:
            Select Case intErrorField
                Case enmHacColNo.Syobun
                    FncGetErrorMsgType_Hac = "0(未設定)、1(返品)、2(一時停止)、3(取扱終了(店振り対象))以外が設定されています。"
            End Select
        Case PARAM_EARLYCANCELDAY:   FncGetErrorMsgType_Hac = "発注開始日＜発注終了日となるように入力して下さい。"
'********** 2005/05/07 北澤 **********↓
        Case PARAM_MISENTAKU: FncGetErrorMsgType_Hac = "「00」(未選択)は指定出来ません。"
'********** 2005/05/07 北澤 **********↑
    End Select

End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncGetErrorMsgType_Bun
'*【処理概要】
'*?    発生したエラーの内容を文字列で取得します。
'*?    (分類情報版）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intErrorField ; エラーが発生したフィールド
'*P    intErrorType ; 発生したエラーの種別
'*P    strErrorParameter ; エラーメッセージ補助用文字列
'*【戻り値】
'*    < Name Type Comment >
'*R    FncGetErrorMsgType_Bun; String;発生したエラーの内容(フィールドごとに異なることもある)
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/09/01; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgType_Bun(intErrorField As Integer, _
                              intErrorType As Integer, _
                              Optional strErrorParameter As String = "")
     
    Select Case intErrorType
        Case PARAM_EMPTY:           FncGetErrorMsgType_Bun = "必須項目です。"
        Case PARAM_NOTDISIT:        FncGetErrorMsgType_Bun = "半角数字で入力して下さい。"
        Case PARAM_NONARROW:        FncGetErrorMsgType_Bun = "半角文字で入力して下さい。"
        Case PARAM_HANEISUU:        FncGetErrorMsgType_Bun = "半角英数字で入力して下さい。"
        Case PARAM_WRONGCHECKDIGIT: FncGetErrorMsgType_Bun = "チェックデジットを確認して下さい。"
        Case PARAM_WRONGDATE:       FncGetErrorMsgType_Bun = "日付を正しく入力して下さい。YYYYMMDD"
        Case PARAM_SHORTDISIT:
            Select Case intErrorField
                Case enmBunColNo.janCD
                    FncGetErrorMsgType_Bun = "桁数に過不足があります。8桁もしくは13桁の数字を入力して下さい。"
                Case enmBunColNo.Bunrui1
                    FncGetErrorMsgType_Bun = "桁数に過不足があります。6桁で入力して下さい。"
                Case enmBunColNo.Bunrui2
                    FncGetErrorMsgType_Bun = "桁数に過不足があります。6桁で入力して下さい。"
                Case enmBunColNo.Bunrui3
                    FncGetErrorMsgType_Bun = "桁数に過不足があります。6桁で入力して下さい。"
            End Select
        
        Case PARAM_NOTEXISTDB:      FncGetErrorMsgType_Bun = "JANコードは統合DBに登録されていません。"
        Case PARAM_SAMECODEDB:      FncGetErrorMsgType_Bun = "JANコードは企画№" & strErrorParameter & "に登録されています。"
        Case PARAM_SAMEKEYDB:       FncGetErrorMsgType_Bun = "内部コードは企画№" & strErrorParameter & "に登録されています。"
        Case PARAM_BUNRUIDB:        FncGetErrorMsgType_Bun = "分類情報が存在しません。"
    End Select

End Function

' '*< UNISYSTEM >**************************************************'
'*【プロシージャ名】
'*!    FncGetErrorMsgType_Syo
'*【処理概要】
'*?    発生したエラーの内容を文字列で取得します。
'*?    (商品基本情報版）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intErrorField ; エラーが発生したフィールド
'*P    intErrorType ; 発生したエラーの種別
'*P    strErrorParameter ; エラーメッセージ補助用文字列
'*【戻り値】
'*    < Name Type Comment >
'*R    FncGetErrorMsgType_Syo; String;発生したエラーの内容(フィールドごとに異なることもある)
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/25; 浅野　正徳(USC); 新規作成;
'*M    2006/09/12; 和田(IC);        商品選定ＷＦ問合せ撲滅改善;
'*M    2014/05/02; 和田(LIS);       禁止文字メッセージ追加;
'*M    2015/09/15; 流通システムG/和田;        TCシステム構築;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgType_Syo(intErrorField As Integer, _
                              intErrorType As Integer, _
                              Optional strErrorParameter As String = "")
     
    Select Case intErrorType
        Case PARAM_EMPTY:           FncGetErrorMsgType_Syo = "必須項目です。"
        Case PARAM_NOTDISIT:        FncGetErrorMsgType_Syo = "半角数字で入力して下さい。"
        Case PARAM_NONARROW:        FncGetErrorMsgType_Syo = "半角文字で入力して下さい。"
        Case PARAM_HANEISUU:        FncGetErrorMsgType_Syo = "半角英数字で入力して下さい。"
        Case PARAM_ZENKAKU:         FncGetErrorMsgType_Syo = "全角文字で入力して下さい。"
        Case PARAM_KINSIMOJI:       FncGetErrorMsgType_Syo = "禁止文字『,'""\<>』が入力されています。"
        '2014/05/02 add wadak7
        Case PARAM_KINSIMOJI_RECEIPT:       FncGetErrorMsgType_Syo = "禁止文字『①②③…ⅠⅡⅢ…㎜㎝㎞…㊤㊥㊦…』等が入力されています。"
        Case PARAM_NOTFINDMASTER:
            Select Case intErrorField
                Case enmSyoColNo.VhJyuyouCd
                    FncGetErrorMsgType_Syo = "VH需要コードがマスタに存在しません。"
                Case enmSyoColNo.RdJyuyouCd
                    FncGetErrorMsgType_Syo = "R&D需要コードがマスタに存在しません。"
                Case enmSyoColNo.VcsJyuyouCd
                    FncGetErrorMsgType_Syo = "VCS需要コードがマスタに存在しません。"
                Case enmSyoColNo.SvhJyuyouCd
                    FncGetErrorMsgType_Syo = "SVH需要コードがマスタに存在しません。"
                Case enmSyoColNo.AtukaiPtn
                    FncGetErrorMsgType_Syo = "扱いパターンコードがマスタに存在しません。"
            End Select
        Case PARAM_WRONGCHECKDIGIT: FncGetErrorMsgType_Syo = "チェックデジットを確認して下さい。"
        Case PARAM_WRONGDATE:       FncGetErrorMsgType_Syo = "日付を正しく入力して下さい。YYYYMMDD"
        Case PARAM_SHORTDISIT:
            Select Case intErrorField
                Case enmSyoColNo.janCD:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。8桁もしくは13桁の数字を入力して下さい。"
                Case enmSyoColNo.SyouhinmeiKana:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。40桁以内の半角カナ文字で入力して下さい。"
                Case enmSyoColNo.SyouhinmeiKanji:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。30桁以内の全角文字を入力して下さい。"
                Case enmSyoColNo.ReceiptmeiKana:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。14桁以内の半角カナ文字を入力して下さい。"
                Case enmSyoColNo.ReceiptmeiKanji:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。14桁以内の全角文字を入力して下さい。"
                Case enmSyoColNo.SyouhihinmeiHansoku:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。20桁以内の全角文字を入力して下さい。"
                
                Case enmSyoColNo.MakerMei:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。10桁以内の全角文字を入力して下さい。"
                Case enmSyoColNo.KikakuSize:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。20桁以内の全角文字を入力して下さい。"
                Case enmSyoColNo.HinbanKataban:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。20桁以内の全角文字を入力して下さい。"
                Case enmSyoColNo.IroGara:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。20桁以内の全角文字を入力して下さい。"
                Case enmSyoColNo.Taste:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。20桁以内の全角文字を入力して下さい。"
                Case enmSyoColNo.Seisantimei:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。20桁以内の全角文字を入力して下さい。"
                Case enmSyoColNo.Tokutyou:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。150桁以内の全角文字を入力して下さい。"
                Case enmSyoColNo.SyouhinCopy:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。50桁以内の全角文字を入力して下さい。"
                Case enmSyoColNo.TuujyouKakaku:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。整数部7桁の数字を入力して下さい。"
                Case enmSyoColNo.MakerKakaku:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。整数部7桁の数字を入力して下さい。"
                Case enmSyoColNo.POPLayout1:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。30桁以内の全角半角文字を入力して下さい。"
                Case enmSyoColNo.POPLayout2:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。30桁以内の全角半角文字を入力して下さい。"
                Case enmSyoColNo.PopSize:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。30桁以内の全角半角文字を入力して下さい。"
                
                Case enmSyoColNo.AtukaiPtn:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。3桁以内の半角文字を入力して下さい。"
                Case enmSyoColNo.HyoujyunBaika:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。整数部7桁の数字を入力して下さい。"
                Case enmSyoColNo.SikiriGenka:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。整数部が7桁以内、小数部が2桁以内で指定して下さい。"
                Case enmSyoColNo.VhJyuyouCd:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。6桁以内の半角文字を入力して下さい。"
                Case enmSyoColNo.RdJyuyouCd:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。6桁以内の半角文字を入力して下さい。"
                Case enmSyoColNo.VcsJyuyouCd:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。6桁以内の半角文字を入力して下さい。"
                Case enmSyoColNo.SvhJyuyouCd:
                    FncGetErrorMsgType_Syo = "桁数に過不足があります。6桁以内の半角文字を入力して下さい。"
            End Select
        Case PARAM_NOTEXISTDB:
            Select Case intErrorField
                Case enmSyoColNo.janCD
                    FncGetErrorMsgType_Syo = "JANコードは統合DBに登録されていません。"
            End Select
        Case PARAM_SAMECODEDB:      FncGetErrorMsgType_Syo = "JANコードは企画№" & strErrorParameter & "に登録されています。"
        Case PARAM_SAMEKEYDB:       FncGetErrorMsgType_Syo = "内部コードは企画№" & strErrorParameter & "に登録されています。"
    End Select

End Function

' '*< UNISYSTEM >**************************************************'
'*【プロシージャ名】
'*!    FncGetErrorMsgType
'*【処理概要】
'*?    発生したエラーの内容を文字列で取得します。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intErrorField ; エラーが発生したフィールド
'*P    intErrorType ; 発生したエラーの種別
'*P    strErrorParameter ; エラーメッセージ補助用文字列
'*【戻り値】
'*    < Name Type Comment >
'*R    FncGetErrorMsgField; String;発生したエラーの内容(フィールドごとに異なることもある)
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/04/19; 浅野　正徳(USC); 新規作成;
'*M    2006/09/13; 和田(IC);        商品選定ＷＦ問合せ撲滅改善;
'*M    2009/05/12; 和田(IC);        グループ品はエラーにする;
'*M    2009/06/06; Supattra T.(Tostem Thai); Add error message;
'*M    2011/02/10; 和田(IC);        センターコード追加;
'*M    2015/09/15; 和田(IC);        TCシステム構築;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgType(intErrorField As Integer, _
                              intErrorType As Integer, _
                              Optional strErrorParameter As String = "")
     
    Select Case intErrorType
        Case PARAM_EMPTY:
            Select Case intErrorField
                Case enmModColNo.centerCD
                    '2015/09/15 wadak7 upd
                    'FncGetErrorMsgType = "センター区分が1：備蓄センターの場合、必須入力項目です。"
                    FncGetErrorMsgType = "センター区分が1：備蓄センター・3：TCセンターの場合、必須入力項目です。"
                '********** 2012/08/03 楊双（大連）入力チェックの追加 **********↓
                Case enmModColNo.Genka, enmModColNo.Hacchuutani
                    FncGetErrorMsgType = "センター区分が2：PDCの場合、必須入力項目です。"
                '********** 2012/08/03 楊双（大連）入力チェックの追加 **********↑
                Case Else
                    FncGetErrorMsgType = "必須項目です。"
            End Select
        Case PARAM_NOTDISIT:        FncGetErrorMsgType = "半角数字で入力して下さい。"
        Case PARAM_NONARROW:        FncGetErrorMsgType = "半角文字で入力して下さい。"
        Case PARAM_NOTFINDMASTER:
            Select Case intErrorField
                Case enmModColNo.TenCd
                    FncGetErrorMsgType = "マスタに登録されていない店コード・事業コードが設定されています。"
                Case enmModColNo.BaikaRiyuu
                    FncGetErrorMsgType = "マスタに登録されていない売価変更理由が設定されています。"
                Case enmModColNo.Torihikisaki
                    FncGetErrorMsgType = "マスタに登録されていない取引先が設定されています。"
            End Select

        Case PARAM_WRONGCHECKDIGIT: FncGetErrorMsgType = "チェックデジットを確認して下さい。"
        Case PARAM_WRONGDATE:       FncGetErrorMsgType = "日付を正しく入力して下さい。YYYYMMDD"
        
        Case PARAM_SHORTDISIT:
            Select Case intErrorField
                Case enmModColNo.janCD
                    FncGetErrorMsgType = "桁数に過不足があります。8桁もしくは13桁の数字を入力して下さい。"
                Case enmModColNo.Genka
                    FncGetErrorMsgType = "桁数に過不足があります。整数部が7桁以内、小数部が2桁以内で指定して下さい。"
                Case enmModColNo.Baika
                    FncGetErrorMsgType = "桁数に過不足があります。整数7桁以内で指定して下さい。"
                Case enmModColNo.Hacchuutani
                    FncGetErrorMsgType = "桁数に過不足があります。整数5桁以内で指定して下さい。"
                Case enmModColNo.centerCD
                    FncGetErrorMsgType = "桁数に過不足があります。4桁の数字を入力して下さい。"
                '2009.06.06 Add by Supattra T.(Tostem Thai)Start.
                Case enmModColNo.MoGenka
                    FncGetErrorMsgType = "桁数に過不足があります。整数部が7桁以内、小数部が2桁以内で指定して下さい。"
                Case enmModColNo.MoBaika
                    FncGetErrorMsgType = "桁数に過不足があります。整数7桁以内で指定して下さい。"
                Case enmModColNo.HattyuuTani
                    FncGetErrorMsgType = "桁数に過不足があります。整数5桁以内で指定して下さい。"
                 Case enmModColNo.PopSize
                    FncGetErrorMsgType = "半角20桁又は全角10桁以内で入力して下さい。"
                '2009.06.06 Add by Supattra T.(Tostem Thai)End.
            End Select
        
        Case PARAM_NOTEXISTDB:
            Select Case intErrorField
                Case enmTenColNo.TenCd
                    FncGetErrorMsgType = "マスタに登録されていない店コード・事業コードが設定されています。"
                Case enmTenColNo.janCD
                    FncGetErrorMsgType = "JANコードは統合DBに登録されていません。"
                Case enmTenColNo.BaikaRiyuu
                    FncGetErrorMsgType = "マスタに登録されていない売価変更理由が設定されています。"
                Case enmTenColNo.Torihikisaki
'                    FncGetErrorMsgType = "マスタに登録されていない取引先が設定されています｡"
                    FncGetErrorMsgType = "マスタに登録されていないか、または取引が停止されている取引先が設定されています。"
                Case enmTenColNo.CenterKbn
                    FncGetErrorMsgType = "マスタに登録されていないセンター区分が設定されています。"
                Case enmModColNo.centerCD
                    FncGetErrorMsgType = "マスタに登録されていないセンターコードが設定されています。"
            End Select
        Case PARAM_SAMECODEDB:      FncGetErrorMsgType = "JANコードは企画№" & strErrorParameter & "に登録されています。"
        Case PARAM_SAMEKEYDB:       FncGetErrorMsgType = "内部コードは企画№" & strErrorParameter & "に登録されています。"
        Case PARAM_TORIHIKISTOP:    FncGetErrorMsgType = "取引先との取引は停止されています。"
        Case PARAM_REPDATEBAIKA:    FncGetErrorMsgType = "既に統合DB(売価変更情報)の先付変更情報に登録されている日付です。"
        Case PARAM_REPDATESIIRE:    FncGetErrorMsgType = "既に統合DB(仕入変更情報)の先付変更情報に登録されている日付です。"
        Case PARAM_REPDATETENPO:    FncGetErrorMsgType = "既に統合DB(店舗センターコード変更情報)の先付変更情報に登録されている日付です。"

'********** 2005/05/07 北澤 **********↓
        Case PARAM_ZERO: FncGetErrorMsgType = "０より大きい値を指定して下さい。"
'********** 2005/05/07 北澤 **********↑
    
'********** 2005/06/16 北澤 **********↓
        Case PARAM_BAIKARIYUU_HISSU: FncGetErrorMsgType = "店コード「0000」の場合、売価が入力されていたら、売価変更理由は入力必須です。"
        Case PARAM_BAIKARIYUU_FUKA:  FncGetErrorMsgType = "店コードが「0000」以外の場合、又は売価が未入力の場合は、売価変更理由は入力不可です。"
'********** 2005/06/16 北澤 **********↑
        Case PARAM_GROUP:    FncGetErrorMsgType = "グループ品は登録できません。"
        '2015/09/15 wadak7 upd
        'Case PARAM_KINSI:    FncGetErrorMsgType = "センター区分が1：備蓄センター以外の場合、入力できません。"
        Case PARAM_KINSI:    FncGetErrorMsgType = "センター区分が1：備蓄センター・3：TCセンター以外の場合、入力できません。"
'********** 2012/08/03 楊双（大連）入力チェックの追加 **********↓
        Case PARAM_HONBU_NAI:       FncGetErrorMsgType = "センター区分が2：PDCの場合、0000：本部のデータは必須入力です。変更日 " & Split(strErrorParameter, ",")(0) & " JANコード " & Split(strErrorParameter, ",")(1)
        Case PARAM_HATTYUU_TANI:    FncGetErrorMsgType = "センター区分が2：PDCの発注単位は全て同じ値を入力して下さい。変更日 " & Split(strErrorParameter, ",")(0) & " JANコード " & Split(strErrorParameter, ",")(1)
        Case PARAM_GENKA:           FncGetErrorMsgType = "センター区分が2：PDCの原価は全て同じ値を入力して下さい。変更日 " & Split(strErrorParameter, ",")(0) & " JANコード " & Split(strErrorParameter, ",")(1)
        Case PARAM_GENKA_HATTYUUTANI: FncGetErrorMsgType = "更新店舗以外にPDC区分になっている店舗があります。確認の上、再更新して下さい。変更日 " & Split(strErrorParameter, ",")(0) & " JANコード " & Split(strErrorParameter, ",")(1)
'********** 2012/08/03 楊双（大連）入力チェックの追加 **********↑

    End Select

End Function

'********** 2005/04/27 北澤 **********↓
Rem ========================================================================
Rem 機能説明： VMST一般基本テーブルより商品情報を取得し、VSWF表示仕様テーブルにレコードを追加する
Rem 受取引数： エラーメッセージ, CSV行カウント, 企画№, JANコード, ログインユーザーID
Rem 戻り値　： TRUE=成功, FALSE=失敗
Rem ========================================================================
Public Function Insert_t_syousen_hyouji_siyou(ByRef strErrMsg _
                                      , ByVal intN As Integer _
                                      , ByVal strKikakuSyuruiCD As String _
                                      , ByVal strNenRenNO As String _
                                      , ByVal strEdaNO As String _
                                      , ByVal strJanCd As String _
                                      , ByVal strUserId As String) As Boolean
On Error GoTo Err_Insert_t_syousen_hyouji_siyou

    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess") '共通関数(SQLServer接続用)
    Dim rsAdo1 As ADODB.Recordset           'レコードセット
    Dim rsAdo2 As ADODB.Recordset           'レコードセット
    Dim strMethodCallInfo As String         'エラー情報
    Dim strSQL As String                    '作業用文字列
    Dim arrParam As Object                 'キー情報格納用配列
    Const strUdl As String = "VSWF"         'SQLServer接続用UDLファイル名

    '戻り値の初期値をセットする
    Insert_t_syousen_hyouji_siyou = False
    
    'エラー情報を作成しておく
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.Insert_t_syousen_hyouji_siyou(企画№:" & strKikakuSyuruiCD & strNenRenNO & "-" & strEdaNO & "JANコード:" & strJanCd & ")"
    
    '共通関数オブジェクトを作成する
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'VSWF表示仕様テーブルのレコードセットを取得する
    '(TABLE: 表示仕様テーブル)
    '(FIELDS: *)
    '(KEY:企画№, JANコード)
    strSQL = _
    "SELECT *" & _
    "  FROM t_syousen_hyouji_siyou WITH(READCOMMITTED)" & _
    " WHERE kikaku_syurui_cd = ?" & _
    "   AND nen_ren_no = ?" & _
    "   AND eda_no = ?" & _
    "   AND jan_cd = ?"
    
    arrParam = Array(MakeParam("@kikaku_syurui_cd", DT_Char, 3, strKikakuSyuruiCD) _
                   , MakeParam("@nen_ren_no", DT_Char, 10, strNenRenNO) _
                   , MakeParam("@eda_no", DT_Char, 3, strEdaNO) _
                   , MakeParam("@jan_cd", DT_VarChar, 20, strJanCd))
    
        rsAdo1 = objCDBAccess.GetDCRecordset(strUdl, strSQL, arrParam)
    
    'VSWF表示仕様テーブルにレコードが既に存在した場合は処理を終了する(エラーでは無い)
    If rsAdo1.EOF = False Then
        
        ''データベースを閉じる
            rsAdo1 = Nothing
        objCDBAccess = Nothing
        'GetObjectContext.SetComplete
        
        ''戻り値をセットする(エラーでは無いのでTRUEを返す)
        Insert_t_syousen_hyouji_siyou = True
        GoTo Ext_Insert_t_syousen_hyouji_siyou
    
    'VSWF表示仕様テーブルにレコードが存在しない場合はVMST一般基本テーブルより商品情報を取得しVSWF表示仕様テーブルにレコードを追加する
    Else
   
        ''VMST一般基本テーブルのレコードセットを取得する
        ''(TABLE: 一般基本ビュー、主キーマスタ)
        ''(FIELDS: *)
        ''(KEY:JANコード)
        strSQL = _
        "SELECT v_ippan_kihon.*" & _
        "  FROM v_ippan_kihon WITH(READCOMMITTED)" & _
        " INNER JOIN v_key_1 WITH(READCOMMITTED) ON v_ippan_kihon.key_cd = v_key_1.key_cd" & _
        " WHERE v_key_1.jan_cd = ?"

        arrParam = Array(MakeParam("@v_key_1.jan_cd", DT_VarChar, 20, strJanCd))
    
            rsAdo2 = objCDBAccess.GetDCRecordset(strUdl, strSQL, arrParam)
    
        ''VMST一般基本テーブルにレコードが存在しなかったら処理を終了する(予期しないエラー)
        If rsAdo2.EOF Then
            
            strErrMsg = CStr(intN + 1) & "行目：VMST一般基本テーブルにデータが存在しない為VSWF表示仕様テーブルにレコードを追加出来ませんでした。"
            GoTo Err_Insert_t_syousen_hyouji_siyou
    
        ''VMST一般基本テーブルから取得したデータでVSWF表示仕様テーブルにレコードを追加する
        ''(TABLE: 表示仕様テーブル)
        ''(FIELDS: 企画種類コード, 年連番, 枝番, JANコード
        ''       , 表示用商品名(カナ), 表示用商品名(漢字), レシート名(カナ), レシート名(漢字)
        ''       , 商品名(漢字), メーカ名(漢字), 規格/サイズ(漢字), 品番／型番(漢字)
        ''       , 色／柄(漢字), テースト(漢字), 生産地名(漢字), 特徴, 賞味期限/使用期限有無
        ''       , エコマーク商品, グリーン購入法適合商品, GPNデータベース掲載商品
        ''       , 登録UserID, 登録年月日時分秒
        ''(KEY:企画№, JANコード)
        Else
            
            '''切断レコードセットを更新する
            With rsAdo1
                .AddNew
                .Fields("kikaku_syurui_cd") = strKikakuSyuruiCD
                .Fields("nen_ren_no") = strNenRenNO
                .Fields("eda_no") = strEdaNO
                .Fields("jan_cd") = strJanCd
                .Fields("hyoujiyou_syouhin_mei_kana") = rsAdo2("hyoujiyou_syouhin_mei_kana")
                .Fields("hyoujiyou_syouhin_mei_kanji") = rsAdo2("hyoujiyou_syouhin_mei_kanji")
                .Fields("receipt_mei_kana") = rsAdo2("receipt_mei_kana")
                .Fields("receipt_mei_kanji") = rsAdo2("receipt_mei_kanji")
                .Fields("syouhin_mei_kanji") = rsAdo2("syouhin_mei_kanji")
                .Fields("maker_mei_kanji") = rsAdo2("maker_mei_kanji")
                .Fields("kikaku_size_kanji") = rsAdo2("kikaku_size_kanji")
                .Fields("sina_no_kata_no_kanji") = rsAdo2("sina_no_kata_no_kanji")
                .Fields("iro_gara_kanji") = rsAdo2("iro_gara_kanji")
                .Fields("taste_kanji") = rsAdo2("taste_kanji")
                .Fields("seisanti_mei_kanji") = rsAdo2("seisanti_mei_kanji")
                .Fields("tokutyou") = rsAdo2("tokutyou")
                .Fields("syoumi_siyoukigen_um_kbn") = rsAdo2("syoumi_siyoukigen_um_kbn")
                .Fields("eco_syouhin_kbn") = rsAdo2("eco_syouhin_kbn")
                .Fields("green_tekigou_syouhin_kbn") = rsAdo2("green_tekigou_syouhin_kbn")
                .Fields("gpn_db_keisai_syouhin_kbn") = rsAdo2("gpn_db_keisai_syouhin_kbn")
                .Fields("touroku_user_id") = strUserId
                .Fields("touroku_ymd_hms") = Format(Now, "yyyy/mm/dd hh:nn:ss")
                .Update
            End With
            
            '''データベースを更新する
            Call objCDBAccess.UpdateDCRecordset(strUdl, rsAdo1)
            
            '''戻り値をセットする
            Insert_t_syousen_hyouji_siyou = True
        
        End If
    End If
    
    'データベースを閉じる
        rsAdo1 = Nothing
        rsAdo2 = Nothing
    objCDBAccess = Nothing
    'GetObjectContext.SetComplete

    'エラー時の処理
Ext_Insert_t_syousen_hyouji_siyou:
    Exit Function
Err_Insert_t_syousen_hyouji_siyou:
    If Not rsAdo1 Is Nothing Then Set rsAdo1 = Nothing
    If Not rsAdo2 Is Nothing Then Set rsAdo2 = Nothing
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
'********** 2005/05/20 土田 変更**********↓
'    Call ErrRaise(DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description)
    Call ErrRaise(DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。" & strErrMsg, 'Err.Number, Err.Source, Err.Description)
'********** 2005/05/20 土田 変更**********↑
    Resume Ext_Insert_t_syousen_hyouji_siyou

End Function
'********** 2005/04/27 北澤 **********↑

'********** 2005/05/07 北澤 **********↓
Rem ========================================================================
Rem 機能説明： ある１事業のレコードを作成する際、他の全ての事業のレコードも作成
Rem 受取引数： 企画№, JANコード, 作成した事業コード, ログインユーザーID
Rem 戻り値　： TRUE=成功, FALSE=失敗
Rem ========================================================================
Public Function Insert_t_syousen_jigyoubetu_settei(ByRef strErrMsg _
                                           , ByVal intN As Integer _
                                           , ByVal strKikakuSyuruiCD As String _
                                           , ByVal strNenRenNO As String _
                                           , ByVal strEdaNO As String _
                                           , ByVal strJanCd As String _
                                           , ByVal strJigyouCd As String _
                                           , ByVal strUserId As String) As Boolean
On Error GoTo Err_Insert_t_syousen_jigyoubetu_settei
    
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess") '共通関数(SQLServer接続用)
    Dim rsAdo1 As ADODB.Recordset           'レコードセット
    Dim rsAdo2 As ADODB.Recordset           'レコードセット
    Dim rsAdo3 As ADODB.Recordset           'レコードセット
    Dim strMethodCallInfo As String         'エラー情報
    Dim strSQL As String                    '作業用文字列
    Const strUdl As String = "VSWF"         'SQLServer接続用UDLファイル名
    
    '戻り値の初期値をセットする
    Insert_t_syousen_jigyoubetu_settei = False
    
    'エラー情報を作成しておく
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.Insert_t_syousen_jigyoubetu_settei(企画№:" & strKikakuSyuruiCD & strNenRenNO & "-" & strEdaNO & "JANコード:" & strJanCd & ")"
    
    '共通関数オブジェクトを作成する
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")

    'VRTL事業所マスタから全事業コードを取得する
    '(TABLE: 事業所マスタ)
    '(FIELDS: 事業コード)
    '(条件: 閉店日がNULL又は一括登録日以降，KEY区分=W02，設定内容の4桁目が1）
    strSQL = _
    "SELECT DISTINCT v_jigyousyo.jigyou_cd" & _
    "  FROM v_jigyousyo WITH(READCOMMITTED)" & _
    " INNER JOIN m_wf_meisyou WITH(READCOMMITTED)" & _
    "    ON v_jigyousyo.jigyou_cd = m_wf_meisyou.key_cd" & _
    " WHERE SUBSTRING(m_wf_meisyou.settei_naiyou,4,1) = '1'" & _
    "   AND m_wf_meisyou.key_kbn = '" & KEYCD_KIKAKU_DOUNYU_TAISYOU & "'"
    
        rsAdo1 = objCDBAccess.GetDCRecordset(strUdl, strSQL)

    '事業コードが無い場合は処理を終了する(あり得ない、エラーにはしない)
    If rsAdo1.EOF = True Then
        ''データベースを閉じる
            rsAdo1 = Nothing
        objCDBAccess = Nothing
        'GetObjectContext.SetComplete
        ''戻り値をセットする(エラーでは無いのでTRUEを返す)
        Insert_t_syousen_jigyoubetu_settei = True
        GoTo Ext_Insert_t_syousen_jigyoubetu_settei
    End If

    'CSVに入力されている１事業を除き、VSWF事業別設定テーブルにレコードを追加する
    Do While Not rsAdo1.EOF
        If Trim(rsAdo1("jigyou_cd")) <> strJigyouCd Then
            
            ''VSWF事業別設定テーブルのレコードセットを取得する
            ''(TABLE: 事業別設定テーブル)
            ''(FIELDS: *)
            ''(KEY:企画№, JANコード, 事業コード)
            strSQL = _
            "SELECT *" & _
            "  FROM t_syousen_jigyoubetu_settei" & _
            " WHERE kikaku_syurui_cd = '" & strKikakuSyuruiCD & "'" & _
            "   AND nen_ren_no = '" & strNenRenNO & "'" & _
            "   AND eda_no = '" & strEdaNO & "'" & _
            "   AND jan_cd = '" & strJanCd & "'" & _
            "   AND jigyou_cd = '" & Trim(rsAdo1("jigyou_cd")) & "'"
    
                rsAdo2 = objCDBAccess.GetDCRecordset(strUdl, strSQL)
    
            ''VSWF事業別設定テーブルにレコードが既に存在した場合は処理を終了する(エラーでは無い)
            If rsAdo2.EOF = False Then
                '''データベースを閉じる
                    rsAdo1 = Nothing
                    rsAdo2 = Nothing
                objCDBAccess = Nothing
                'GetObjectContext.SetComplete
                '''戻り値をセットする(エラーでは無いのでTRUEを返す)
                Insert_t_syousen_jigyoubetu_settei = True
                GoTo Ext_Insert_t_syousen_jigyoubetu_settei
            End If
            
            ''VMST事業別テーブルよりレコードセットを取得する
            ''(TABLE: 事業別ビュー、主キーマスタ)
            ''(FIELDS: *)
            ''(KEY:JANコード, 事業コード)
'********** 2005/05/20 土田 **********↓
'発注終了日のフィールドもコンバートをかけるよう変更
            strSQL = _
            "SELECT *" & _
            "     , CONVERT(datetime,v_jigyou_betu.hattyuu_kaisi_date) AS CONVERT_hattyuu_kaisi_date" & _
            "     , CONVERT(datetime,v_jigyou_betu.hattyuu_syuuryou_date) AS CONVERT_hattyuu_syuuryou_date" & _
            "  FROM v_jigyou_betu WITH(READCOMMITTED)" & _
            " INNER JOIN v_key_1 WITH(READCOMMITTED)" & _
            "    ON v_jigyou_betu.key_cd = v_key_1.key_cd" & _
            " WHERE v_key_1.jan_cd = '" & strJanCd & "'" & _
            "   AND v_jigyou_betu.jigyou_cd = '" & Trim(rsAdo1("jigyou_cd")) & "'"

                rsAdo3 = objCDBAccess.GetDCRecordset(strUdl, strSQL)
    
            ''VMST事業別テーブルにレコードが存在しなかったら処理を終了する(エラーにはしない)
            If rsAdo3.EOF Then
                '''データベースを閉じる
                    rsAdo1 = Nothing
                    rsAdo2 = Nothing
                    rsAdo3 = Nothing
                objCDBAccess = Nothing
                'GetObjectContext.SetComplete
                '''戻り値をセットする(エラーでは無いのでTRUEを返す)
                Insert_t_syousen_jigyoubetu_settei = True
                GoTo Ext_Insert_t_syousen_jigyoubetu_settei
            End If
            
            ''VSWF事業別設定テーブルにレコードを追加する
            ''(TABLE: 事業別設定テーブル)
            ''(FIELDS: *)
            ''(KEY:企画№, JANコード, 事業コード)
            '''切断レコードセットを更新する
            With rsAdo2
                .AddNew
                .Fields("kikaku_syurui_cd") = strKikakuSyuruiCD
                .Fields("nen_ren_no") = strNenRenNO
                .Fields("eda_no") = strEdaNO
                .Fields("jan_cd") = strJanCd
                .Fields("jigyou_cd") = Trim(rsAdo1("jigyou_cd"))
                .Fields("hattyuu_kaisi_date") = rsAdo3("CONVERT_hattyuu_kaisi_date")
'********** 2005/05/20 土田 変更**********↓
'                .Fields("hattyuu_syuuryou_date") = rsAdo3("hattyuu_syuuryou_date")
                .Fields("hattyuu_syuuryou_date") = rsAdo3("CONVERT_hattyuu_syuuryou_date")
'********** 2005/05/20 土田 変更**********↑
                .Fields("hattyuu_tyuusi_riyuu_kbn") = rsAdo3("hattyuu_tyuusi_riyuu_kbn")
                .Fields("syobun_houhou") = rsAdo3("syobun_houhou")
                .Fields("daitai_syouhin_cd") = rsAdo3("daitai_syouhin_cd")
                .Fields("dounyuu_baika") = rsAdo3("dounyuu_baika")
                .Fields("jyougen_baika") = rsAdo3("jyougen_baika")
                .Fields("kagen_baika") = rsAdo3("kagen_baika")
                .Fields("touroku_user_id") = strUserId
                .Fields("touroku_ymd_hms") = Now()
                .Update
            End With

            '''データベースを更新する
            Call objCDBAccess.UpdateDCRecordset(strUdl, rsAdo2)
                rsAdo2 = Nothing
                rsAdo3 = Nothing
        
        End If
        rsAdo1.MoveNext
    Loop

    '戻り値をセットする
    Insert_t_syousen_jigyoubetu_settei = True
    
    'データベースを閉じる
        rsAdo1 = Nothing
    objCDBAccess = Nothing
    'GetObjectContext.SetComplete

    'エラー時の処理
Ext_Insert_t_syousen_jigyoubetu_settei:
    Exit Function
Err_Insert_t_syousen_jigyoubetu_settei:
    If Not rsAdo1 Is Nothing Then Set rsAdo1 = Nothing
    If Not rsAdo2 Is Nothing Then Set rsAdo1 = Nothing
    If Not rsAdo3 Is Nothing Then Set rsAdo1 = Nothing
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    Call ErrRaise(DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description)
    Resume Ext_Insert_t_syousen_jigyoubetu_settei

End Function
'********** 2005/05/07 北澤 **********↑

'********** 2005/05/07 北澤 **********↓
Rem ========================================================================
Rem 機能説明： 金額類が正の値かをチェックする
Rem 受取引数： CSVデータ配列, CSVファイルの現在の行
Rem 戻り値　： TRUE=エラー無し, FALSE=エラー有り
Rem ========================================================================
Private Function FncCheck_NotZero(ByRef arryCSVRecord() As String _
                          , ByVal intNowLine As Integer _
                          , ByVal strTenHac As String) As Boolean

'戻り値の初期値をセットする
    FncCheck_NotZero = True
    
'------------------------------
'金額類が正数値かをチェックする
'------------------------------
    ''店別一括登録(TEN)の場合
    If strTenHac = "TEN" Then
        '''原価
        
'********** 2005/05/18 田中 変更 **********↓
'        If arryCSVRecord(enmTenColNo.Genka) <> "" Then
        If arryCSVRecord(enmTenColNo.Genka) <> "" And IsNumeric(arryCSVRecord(enmTenColNo.Genka)) Then
'********** 2005/05/18 田中 変更 **********↑
            
            If CCur(arryCSVRecord(enmTenColNo.Genka)) <= 0 Then
                FncSetError (intNowLine, enmTenColNo.Genka, PARAM_ZERO
                FncCheck_NotZero = False
            End If
        End If
        '''売価
        
'********** 2005/05/18 田中 変更 **********↓
'        If arryCSVRecord(enmTenColNo.Baika) <> "" Then
        If arryCSVRecord(enmTenColNo.Baika) <> "" And IsNumeric(arryCSVRecord(enmTenColNo.Baika)) Then
'********** 2005/05/18 田中 変更 **********↑
           
           If CCur(arryCSVRecord(enmTenColNo.Baika)) <= 0 Then
                FncSetError (intNowLine, enmTenColNo.Baika, PARAM_ZERO
                FncCheck_NotZero = False
            End If
        End If
        '''発注単位
        
'********** 2005/05/18 田中 変更 **********↓
'        If arryCSVRecord(enmTenColNo.Hacchuutani) <> "" Then
        If arryCSVRecord(enmTenColNo.Hacchuutani) <> "" And IsNumeric(arryCSVRecord(enmTenColNo.Hacchuutani)) Then
'********** 2005/05/18 田中 変更 **********↑
            
            If CCur(arryCSVRecord(enmTenColNo.Hacchuutani)) <= 0 Then
                FncSetError (intNowLine, enmTenColNo.Hacchuutani, PARAM_ZERO
                FncCheck_NotZero = False
            End If
        End If
    ''発注中止日一括登録(HAC)の場合
    Else
        '''対象項目無し
    End If

'------------------------------
'チェックを追加する場合はここに記述する
'------------------------------
'********** 2005/06/16 北澤 **********↓
'売価変更理由は、店=0000かつ売価<>""の場合は入力必須、それ以外の場合は入力不可
    ''店別一括登録(TEN)の場合のみ実行
    If strTenHac = "TEN" Then
        
        '''店=0000の場合
        If arryCSVRecord(enmTenColNo.TenCd) = "0000" Then
            ''''売価が入力されていたら、売価変更理由は入力必須
            If arryCSVRecord(enmTenColNo.Baika) <> "" Then
                If arryCSVRecord(enmTenColNo.BaikaRiyuu) = "" Then
                    FncSetError (intNowLine, enmTenColNo.BaikaRiyuu, PARAM_BAIKARIYUU_HISSU
                    FncCheck_NotZero = False
                End If
            End If
            ''''売価が未入力だったら、売価変更理由は入力不可
            If arryCSVRecord(enmTenColNo.Baika) = "" Then
                If arryCSVRecord(enmTenColNo.BaikaRiyuu) <> "" Then
                    FncSetError (intNowLine, enmTenColNo.BaikaRiyuu, PARAM_BAIKARIYUU_FUKA
                    FncCheck_NotZero = False
                End If
            End If
        End If
        
        '''店<>0000の場合
        If arryCSVRecord(enmTenColNo.TenCd) <> "0000" Then
            ''''売価変更理由は入力不可
            If arryCSVRecord(enmTenColNo.BaikaRiyuu) <> "" Then
                FncSetError (intNowLine, enmTenColNo.BaikaRiyuu, PARAM_BAIKARIYUU_FUKA
                FncCheck_NotZero = False
            End If
        End If
    
    End If
'********** 2005/06/16 北澤 **********↑

End Function
'********** 2005/05/07 北澤 **********↑
'
'*< TOSTEM >*******************************************************************'
'*
'*【プロシージャ名】
'*!     FncBunTouroku
'*
'*【処理概要】
'*?     分類情報一括変更のデータを商品選定WFのDBに取り込む
'*
'*【パラメータ】
'*      Name                   Type                Comment
'*P     strErrMsg              Object             エラーメッセージ;
'*【戻り値】
'*      Name                   Type                Comment
'*R     FncBunTouroku          Boolean;
'*
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)            EDP-No      説明
'*M 2005/08/31; 高附 (IC);             P-24884;    新規作成;
'*M 2006/09/14; 和田(IC);              P-XXXXX;    商品選定ＷＦ問合せ撲滅改善;
'*
'*< Comment End >**************************************************************
Public Function FncBunTouroku(ByRef strErrMsg) As Boolean
    
    Dim intN As Integer
    Dim blnRet As Boolean
    Dim strJanCd As String
    Dim intNowLine As Integer
    Dim strCSVRecord() As String
    Dim flgChkError As Boolean      'チェック時のエラーフラグ（TRUE:エラーあり）
    
    FncBunTouroku = False
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    
    'DB参照エラーチェック用切断レコードセットの生成
    Dim strSQL          As String
    Dim objCDBAccess    As Object
        
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")

    '一括登録時、毎回一つユーザーのみ実行できる
    Call LockTableIkkatu(objCDBAccess)
    
    Call SetIkkatuLog("1-1", "分類情報一括変更のデータを商品選定WFのDBに取り込む", "FncBunTouroku", "", "", pstrTantousyaId, "INS")
    Call SetIkkatuLog("1-1-1", "重複JANコードチェック用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")
    '重複JANコードチェック用切断レコードセット生成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(重複JANコードチェック用切断レコードセット生成)"
    strSQL = "SELECT twk.jan_cd, mk1.key_cd, twk.kikaku_no" & vbCrLf
    strSQL = strSQL & " FROM t_syousen_wf_kanri as twk WITH(readcommitted)" & vbCrLf
    strSQL = strSQL & "      INNER JOIN v_key_1 as mk1 WITH(readcommitted)" & vbCrLf
    strSQL = strSQL & "              ON(twk.jan_cd = mk1.jan_cd)" & vbCrLf
    strSQL = strSQL & " WHERE" & vbCrLf
    strSQL = strSQL & " NOT(" & vbCrLf
    strSQL = strSQL & "      kikaku_syurui_cd = '" & pstrKikakuSyurui & "' " & vbCrLf
    strSQL = strSQL & "      AND nen_ren_no = '" & pstrKikakuNenrenban & "' " & vbCrLf
    strSQL = strSQL & "      AND eda_no = '" & pstrKikakuEdaban & "'" & vbCrLf
    strSQL = strSQL & "    )" & vbCrLf
    strSQL = strSQL & " AND twk.touroku_kbn = '2'" & vbCrLf
    
        pChkSameJANCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pChkSameJANCodeRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    pChkSameJANCodeRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ
    Call SetIkkatuLog("1-1-1", "", "", "", "OK", "", "UPD")
    
    Call SetIkkatuLog("1-1-2", "CSVファイル内Janコードのリスト(Janコードの重複無し)作成", "SQL", "", "", pstrTantousyaId, "INS")
    'CSVファイル内Janコードのリスト(Janコードの重複無し)作成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(CSVファイル内Janコードのリスト(Janコードの重複無し)作成)"

    intNowLine = 0
    Do Until UBound(paryCsvData) < intNowLine
        
        strCSVRecord = paryCsvData(intNowLine)
         
        If InStr(pCsvJanCodeList, "'" & strCSVRecord(enmBunColNo.janCD) & "'") = 0 Then
            pCsvJanCodeList = pCsvJanCodeList & "'" & strCSVRecord(enmBunColNo.janCD) & "',"
        End If

        intNowLine = intNowLine + 1
    Loop
    
    If pCsvJanCodeList <> "" Then
        pCsvJanCodeList = Left(pCsvJanCodeList, Len(pCsvJanCodeList) - 1)
    Else
        pCsvJanCodeList = "''"
    End If
    intNowLine = 0
        Call SetIkkatuLog("1-1-2", "", "", "", "OK", "", "UPD")
        Call SetIkkatuLog("1-1-3", "CSVファイル内JANコード紐付け主キーのリスト(主キー重複無し)作成", "SQL", "", "", pstrTantousyaId, "INS")
    '統合マスタ登録済みJANコードチェック用切断レコードセット生成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(統合マスタ登録済みJANコードチェック用切断レコードセット生成)"

    strSQL = "SELECT jan_cd, key_cd FROM v_key_1 WITH(readcommitted)"
    strSQL = strSQL & " WHERE jan_cd IN(" & pCsvJanCodeList & ") "
    
        pChkvk1JANCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pChkvk1JANCodeRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
        Call SetIkkatuLog("1-1-3", "", "", "", "OK", "", "UPD")
        Call SetIkkatuLog("1-1-4", "CSVファイル内JANコード紐付け主キーのリスト(主キー重複無し)作成", "SQL", "", "", pstrTantousyaId, "INS")
    '統合マスタ登録済みJANコード紐付け主キーのリスト(主キー重複無し)作成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(CSVファイル内JANコード紐付け主キーのリスト(主キー重複無し)作成)"
    With pChkvk1JANCodeRS
        Do Until .EOF
            If InStr(pHKeyCDList, .Fields("key_cd")) = 0 Then
                pHKeyCDList = pHKeyCDList & "'" & .Fields("key_cd") & "',"
            End If
            .MoveNext
        Loop
    
        If pHKeyCDList <> "" Then
            pHKeyCDList = Left(pHKeyCDList, Len(pHKeyCDList) - 1)
        Else
            pHKeyCDList = "''"
        End If
    End With
        Call SetIkkatuLog("1-1-4", "", "", "", "OK", "", "UPD")
        Call SetIkkatuLog("1-1-5", "CSVファイル内第1分類コードのリスト(第1分類コードの重複無し)作成", "SQL", "", "", pstrTantousyaId, "INS")
    'CSVファイル内第1分類コードのリスト(第1分類コードの重複無し)作成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(CSVファイル内第1分類コードのリスト(第1分類コードの重複無し)作成)"

    intNowLine = 0
    Do Until UBound(paryCsvData) < intNowLine
        
        strCSVRecord = paryCsvData(intNowLine)
         
        If InStr(pCsvBunrui1CodeList, strCSVRecord(enmBunColNo.Bunrui1)) = 0 Then
            pCsvBunrui1CodeList = pCsvBunrui1CodeList & "'" & strCSVRecord(enmBunColNo.Bunrui1) & "',"
        End If

        intNowLine = intNowLine + 1
    Loop
    
    If pCsvBunrui1CodeList <> "" Then
        pCsvBunrui1CodeList = Left(pCsvBunrui1CodeList, Len(pCsvBunrui1CodeList) - 1)
    Else
        pCsvBunrui1CodeList = "''"
    End If
    intNowLine = 0
        Call SetIkkatuLog("1-1-5", "", "", "", "OK", "", "UPD")
        Call SetIkkatuLog("1-1-6", "品種表示データ登録済み分類コードチェック用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")
    '品種表示データ登録済み分類コードチェック用切断レコードセット生成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(分類マスタKEYチェック用切断レコードセット生成)"

    strSQL = "SELECT key_bnrui_cd_1, key_bnrui_cd_2, key_bnrui_cd_3, uriba_cd, old_bumon_cd"
    strSQL = strSQL & " ,old_tyuubunrui_cd, old_syoubunrui_cd, bunrui_cd_3"
    strSQL = strSQL & " FROM v_hinsyu_hyouji WITH(readcommitted)"
    strSQL = strSQL & " WHERE key_bnrui_cd_1 IN(" & pCsvBunrui1CodeList & ") "
    
        pChkBunruiCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    With pChkBunruiCodeRS
            .Fields("key_bnrui_cd_1").Properties("Optimize") = True 'オプティマイザ
            .Fields("key_bnrui_cd_2").Properties("Optimize") = True 'オプティマイザ
            .Fields("key_bnrui_cd_3").Properties("Optimize") = True 'オプティマイザ
    End With
    Call SetIkkatuLog("1-1-6", "", "", "", "OK", "", "UPD")
    Call SetIkkatuLog("1-1-7", "全ての配列データのチェックを行う。", "SQL", "", "", pstrTantousyaId, "INS")
    '全ての配列データのチェックを行う。
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(全ての配列データのチェックを行う。)"
    
    flgChkError = False
    intNowLine = 0
    strErrMsg = ""
    Do Until UBound(paryCsvData) < intNowLine
        
        strCSVRecord = paryCsvData(intNowLine)
        
        '(1)DBに依存しないチェック：FncCheckParameter_CorrectForm_Bun
        '(2)DBを参照してチェック:FncCheckParameter_DataBase_Bun
        If FncCheckParameter_CorrectForm_Bun(strCSVRecord, intNowLine + 1) = False Or _
            FncCheckParameter_DataBase_Bun(strCSVRecord, intNowLine + 1) = False Then
            flgChkError = True
        End If
        
        intNowLine = intNowLine + 1
    Loop
        Call SetIkkatuLog("1-1-7", "", "", "", "OK", "", "UPD")
        Call SetIkkatuLog("1-1-8", "DB参照エラーチェック用切断レコードセットの開放", "SQL", "", "", pstrTantousyaId, "INS")
    'DB参照エラーチェック用切断レコードセットの開放
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(DB参照エラーチェック用切断レコードセットの開放)"
        
    'チェック用のオブジェクトを開放する。
    If Not pChkSameJANCodeRS Is Nothing Then
        pChkSameJANCodeRS.Close
            pChkSameJANCodeRS = Nothing
    End If
    Call SetIkkatuLog("1-1-8", "", "", "", "OK", "", "UPD")
    
    'データに１つでも問題がある場合は、ここで終了！！！
    If flgChkError = True Then
        'エラーがあればエラーメッセージを取得する。
        strErrMsg = FncGetErrorMsg_Bun()
        
        Exit Function
    End If
    Call SetIkkatuLog("1-2", "データ更新処理開始", "SQL", "", "", pstrTantousyaId, "INS")
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(データ更新処理開始)"
    
    '全てのテーブルを更新する。
    Call subIns_SetAllTables_Bun
    Call SetIkkatuLog("1-2", "", "", "", "OK", "", "UPD")
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(データ更新処理終了)"
    
    FncBunTouroku = True
    
    'オブジェクトの開放
    If Not pChkvk1JANCodeRS Is Nothing Then
        pChkvk1JANCodeRS.Close
            pChkvk1JANCodeRS = Nothing
    End If
    
    If Not pChkBunruiCodeRS Is Nothing Then
        pChkBunruiCodeRS.Close
            pChkBunruiCodeRS = Nothing
    End If
    
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
        Call SetIkkatuLog("1-1", "", "", "", "OK", "", "UPD")
Exit Function
DefectErrorHandler:
    
    'オブジェクトの開放
    If Not pChkSameJANCodeRS Is Nothing Then
        pChkSameJANCodeRS.Close
            pChkSameJANCodeRS = Nothing
    End If
    
    If Not pChkvk1JANCodeRS Is Nothing Then
        pChkvk1JANCodeRS.Close
            pChkvk1JANCodeRS = Nothing
    End If
    
    If Not pChkBunruiCodeRS Is Nothing Then
        pChkBunruiCodeRS.Close
            pChkBunruiCodeRS = Nothing
    End If
    
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
    
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    Call SetIkkatuLog("1-1", "", "", Left(Err.Description, 400), "NG", "", "UPD")
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description

End Function


'******************************************************************'
'*【プロシージャ名】
'*!    FncCheckParameter_CorrectForm_Bun
'*【処理概要】
'*?    CSVファイルのパラメータに対して、DBに依存しない形式チェックを行います。
'*?    （分類情報一括変更用チェック）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckParameter_CorrectForm_Bun; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/08/31; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
'
Private Function FncCheckParameter_CorrectForm_Bun(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheckParameter_CorrectForm_Bun = False
    
    Dim blnIsOKRequired
    Dim blnIsOKNumeric
    Dim blnIsOKNarrow
    Dim blnIsOKLimit
    Dim blnIsOKCheckDigit
    
    'a.必須項目（変更日、JANコード、第1分類コード、第2分類コード、第3分類コード）が入っているか確認する。
    blnIsOKRequired = FncCheck_RequiredParameter_Bun(arryCSVRecord(), intNowLine)
    
    'b.変更日、JANコードにおいて、半角数字であるか確認する。
    blnIsOKNumeric = FncCheck_IsParameterNumeric_Bun(arryCSVRecord(), intNowLine)
    
    'c.第1分類コード、第2分類コード、第3分類コードにおいて、半角文字であるか確認する。
    blnIsOKNarrow = FncCheck_IsParameterNarrow_Bun(arryCSVRecord(), intNowLine)
    
    'd.桁数が制限を超えていないか確認する。
    blnIsOKLimit = FncCheck_Limit_Bun(arryCSVRecord(), intNowLine)

    'e.JANコードに正しいチェックデジットの値がセットされているかを調べる
    blnIsOKCheckDigit = FncCheck_CheckDigit(arryCSVRecord(enmBunColNo.janCD), intNowLine, 5)
    
    If blnIsOKRequired And blnIsOKNumeric And blnIsOKNarrow And _
       blnIsOKLimit And blnIsOKCheckDigit Then
       
       FncCheckParameter_CorrectForm_Bun = True
    End If
    
End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncCheck_RequiredParameter_Bun
'*【処理概要】
'*?    CSVデータの１行に対し、必須項目が空白でないかどうかチェックします。
'*?    （分類情報一括変更用チェック）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_RequiredParameter_Bun; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/08/31; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_RequiredParameter_Bun(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_RequiredParameter_Bun = True
    
    If arryCSVRecord(enmBunColNo.HenkouDate) = "" Then   '変更日
        FncSetError (intNowLine, enmBunColNo.HenkouDate, PARAM_EMPTY
        FncCheck_RequiredParameter_Bun = False
    End If
    If arryCSVRecord(enmBunColNo.janCD) = "" Then        'JANコード
        FncSetError (intNowLine, enmBunColNo.janCD, PARAM_EMPTY
        FncCheck_RequiredParameter_Bun = False
    End If
    If arryCSVRecord(enmBunColNo.Bunrui1) = "" Then      '分類コード1
        FncSetError (intNowLine, enmBunColNo.Bunrui1, PARAM_EMPTY
        FncCheck_RequiredParameter_Bun = False
    End If
    If arryCSVRecord(enmBunColNo.Bunrui2) = "" Then      '分類コード2
        FncSetError (intNowLine, enmBunColNo.Bunrui2, PARAM_EMPTY
        FncCheck_RequiredParameter_Bun = False
    End If
    If arryCSVRecord(enmBunColNo.Bunrui3) = "" Then      '分類コード3
        FncSetError (intNowLine, enmBunColNo.Bunrui3, PARAM_EMPTY
        FncCheck_RequiredParameter_Bun = False
    End If

End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncCheck_IsParameterNumeric_Bun
'*【処理概要】
'*?    CSVデータの１行に対し、対象フィールドが半角数値かどうかチェックします。
'*?    （分類情報一括変更用チェック）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_IsParameterNumeric_Bun; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/08/31; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterNumeric_Bun(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_IsParameterNumeric_Bun = True

    If FncChkHankakuNumber(arryCSVRecord(enmBunColNo.HenkouDate)) = False Then      '変更日
        FncSetError (intNowLine, enmBunColNo.HenkouDate, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Bun = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmBunColNo.janCD)) = False Then           'JANコード
        FncSetError (intNowLine, enmBunColNo.janCD, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Bun = False
    End If
    
End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncCheck_IsParameterNarrow_Bun
'*【処理概要】
'*?    CSVデータの１行に対し、対象フィールドが半角文字かどうかチェックします。
'*?    （分類情報一括変更用チェック）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_IsParameterNarrow_Bun; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/08/31; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterNarrow_Bun(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_IsParameterNarrow_Bun = True

    If fncChkHankakuEisuu(arryCSVRecord(enmBunColNo.Bunrui1)) = False Then      '第1分類コード
        FncSetError (intNowLine, enmBunColNo.Bunrui1, PARAM_HANEISUU
        FncCheck_IsParameterNarrow_Bun = False
    End If
    
    If fncChkHankakuEisuu(arryCSVRecord(enmBunColNo.Bunrui2)) = False Then      '第2分類コード
        FncSetError (intNowLine, enmBunColNo.Bunrui2, PARAM_HANEISUU
        FncCheck_IsParameterNarrow_Bun = False
    End If
    
    If fncChkHankakuEisuu(arryCSVRecord(enmBunColNo.Bunrui3)) = False Then      '第3分類コード
        FncSetError (intNowLine, enmBunColNo.Bunrui3, PARAM_HANEISUU
        FncCheck_IsParameterNarrow_Bun = False
    End If
    
End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncCheck_Limit_Bun
'*【処理概要】
'*?    CSVデータの１行に対し、フィールドが桁数制限を越えていないかチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_Limit_Bun; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/08/31; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_Limit_Bun(ByRef arryCSVRecord() As String, intNowLine As Integer)

    FncCheck_Limit_Bun = True
    
    '(a)変更日
    If FncCheck_Length_Kotei(arryCSVRecord(enmBunColNo.HenkouDate), LENGTH_HENKOU_DATE) = False _
        Or FncCheckDate(arryCSVRecord(enmBunColNo.HenkouDate)) = False Then
        
        FncSetError (intNowLine, enmBunColNo.HenkouDate, PARAM_WRONGDATE)
        FncCheck_Limit_Bun = False
    End If
    '(b)JANコード
    If FncCheck_JANCode(arryCSVRecord(enmBunColNo.janCD)) = False Then
        FncSetError (intNowLine, enmBunColNo.janCD, PARAM_SHORTDISIT)
        FncCheck_Limit_Bun = False
    End If
    '(c)第1分類コード
    If FncCheck_Length_Less(arryCSVRecord(enmBunColNo.Bunrui1), LENGTH_BUNRUI) = False Then
        FncSetError (intNowLine, enmBunColNo.Bunrui1, PARAM_SHORTDISIT)
        FncCheck_Limit_Bun = False
    End If
    '(d)第2分類コード
    If FncCheck_Length_Less(arryCSVRecord(enmBunColNo.Bunrui2), LENGTH_BUNRUI) = False Then
        FncSetError (intNowLine, enmBunColNo.Bunrui2, PARAM_SHORTDISIT)
        FncCheck_Limit_Bun = False
    End If
    '(e)第3分類コード
    If FncCheck_Length_Less(arryCSVRecord(enmBunColNo.Bunrui3), LENGTH_BUNRUI) = False Then
        FncSetError (intNowLine, enmBunColNo.Bunrui3, PARAM_SHORTDISIT)
        FncCheck_Limit_Bun = False
    End If
    
End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncCheckParameter_DataBase_Bun
'*【処理概要】
'*?    CSVファイルのパラメータに対して、DBにアクセスしてチェックを行います。
'*?    (分類情報版)
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVRecord() ; String CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckParameter_DataBase_Bun; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/08/31; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheckParameter_DataBase_Bun(ByRef arryCSVRecord() As String, intNowLine As Integer)

    Dim blnIsOKJANCode
    Dim blnIsOKNaibuCode
    Dim blnIsOKBunruiCode
    
    Dim strKeyCode As String
    
    FncCheckParameter_DataBase_Bun = "False"
    
    'a.JANコードが有効か確認する。
    blnIsOKJANCode = FncCheckDB_JANCode(arryCSVRecord(enmBunColNo.janCD), strKeyCode, intNowLine, RET_FLD_BUN)
    
    'b.KEYコードが有効か確認する。
    blnIsOKNaibuCode = FncCheckDB_NaibuCode(strKeyCode, intNowLine, RET_FLD_BUN)
    
    'c.第1分類コード、第2分類コード、第3分類コードが有効か確認する。
    blnIsOKBunruiCode = FncCheckDB_BunruiCode(arryCSVRecord(enmBunColNo.Bunrui1), arryCSVRecord(enmBunColNo.Bunrui2) _
                                            , arryCSVRecord(enmBunColNo.Bunrui3), intNowLine, RET_FLD_BUN)
    
    If blnIsOKJANCode And blnIsOKBunruiCode And blnIsOKNaibuCode Then
        FncCheckParameter_DataBase_Bun = "True"
    End If
       
End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncCheckDB_NaibuCode
'*【処理概要】
'*?    企画外商品の内部コード重複チェックをします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strBunrui1 ; String 第1分類コードをあらわす文字列
'*P    strBunrui2 ; String 第2分類コードをあらわす文字列
'*P    strBunrui3 ; String 第3分類コードをあらわす文字列
'*P    intNowLine; Integer; CSVファイルの現在の行
'*P    intFieldType; Integer; RET_FLD_BUN:bunrui情報
'*【戻り値】
'*    < Name Type Comment >
'*R    strKeyCode ; String JANコードに対応する統合マスタのキーコードを格納する
'*R    FncCheckDB_NaibuCode; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/09/02; 高附　周造(IC); 新規作成;
'*M    2006/09/12; 和田(IC); 商品選定ＷＦ問合せ撲滅改善;
'*< Comment End >**************************************************'
Private Function FncCheckDB_NaibuCode(ByVal strKeyCode As String _
                                , intNowLine As Integer, intFieldType As Integer) As Boolean
    
    Dim strKikakuNo As String   'KEYコードが既に登録されている企画№
    
    'a.分類情報が有効か確認する。
    FncCheckDB_NaibuCode = True
    
'項目名の作成（エラーメッセージ用）
    Dim intField As Integer
    Select Case intFieldType
        Case RET_FLD_HACH:      intField = enmHacColNo.janCD
        Case RET_FLD_TENB:      intField = enmTenColNo.janCD
        Case RET_FLD_SYO:       intField = enmSyoColNo.janCD
        Case RET_FLD_BUN:       intField = enmBunColNo.janCD
    End Select

    '　(a)WF管理テーブルに存在しないことを確認する。
    With pChkSameJANCodeRS
        If Not .BOF Then
            .MoveFirst
            .Filter = "key_cd = '" & strKeyCode & "'"
            If Not .EOF Then
                strKikakuNo = .Fields("kikaku_no").Value
            End If
        End If
    End With
 
    If pChkSameJANCodeRS.EOF = False Then
        pChkSameJANCodeRS.Filter = adFilterNone
        FncSetError (intNowLine, intField, PARAM_SAMEKEYDB, strKikakuNo
        FncCheckDB_NaibuCode = False
    End If
    pChkSameJANCodeRS.Filter = adFilterNone
End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncCheckDB_BunruiCode
'*【処理概要】
'*?    第1分類コード、第2分類コード、第3分類コードが有効かをチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strBunrui1 ; String 第1分類コードをあらわす文字列
'*P    strBunrui2 ; String 第2分類コードをあらわす文字列
'*P    strBunrui3 ; String 第3分類コードをあらわす文字列
'*P    intNowLine; Integer; CSVファイルの現在の行
'*P    intFieldType; Integer; RET_FLD_BUN:bunrui情報
'*【戻り値】
'*    < Name Type Comment >
'*R    strKeyCode ; String JANコードに対応する統合マスタのキーコードを格納する
'*R    FncCheckDB_BunruiCode; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/09/02; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheckDB_BunruiCode(ByVal strBunrui1 As String, ByVal strBunrui2 As String, ByVal strBunrui3 As String _
                              , intNowLine As Integer, intFieldType As Integer) As Boolean
    
    'a.分類情報が有効か確認する。
    FncCheckDB_BunruiCode = True
    
'項目名の作成（エラーメッセージ用）
    Dim intField As Integer
    Select Case intFieldType
        Case RET_FLD_BUN:       intField = enmBunColNo.Bunrui1
    End Select
    
    '　(a)品種紐付マスタに存在する事を確認する。
    With pChkBunruiCodeRS
        If Not .BOF Then
            .MoveFirst
            .Filter = "key_bnrui_cd_1 = '" & strBunrui1 & "'" _
                    & "AND key_bnrui_cd_2 = '" & strBunrui2 & "'" _
                    & "AND key_bnrui_cd_3 = '" & strBunrui3 & "'"
        End If
    End With
 
    If pChkBunruiCodeRS.EOF = True Then
       pChkBunruiCodeRS.Filter = adFilterNone
        FncSetError (intNowLine, intField, PARAM_BUNRUIDB
        FncCheckDB_BunruiCode = False
    End If
    pChkBunruiCodeRS.Filter = adFilterNone
End Function
'*< TOSTEM >*******************************************************************'
'*
'*【プロシージャ名】
'*!     fncChkKinsiMoji
'*
'*【処理概要】
'*?     半角チェック(英数のみ)
'*
'*【パラメータ】
'*      Name                   Type                Comment
'*P     strValue               String              値
'*P     strErrMsg              String              エラーメッセージ
'*【戻り値】
'*      Name                   Type                Comment
'*R     fncChkHankaku          Boolean;
'*
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)            EDP-No      説明
'*M 2005/09/20; 高附 (IC);             #-#####;    新規作成;
'*M 2006/06/15; 和田(IC);               P-30109;   ＷＦ仕様改善簡易一括機能の追加;
'*
'*< Comment End >**************************************************************
Private Function fncChkKinsiMoji(ByVal strValue As String) As Boolean
        
    Dim lngCount As Long
    Dim strCheck As String
    
    fncChkKinsiMoji = False
    
    Const PC_KINSOKU_STRING As String = "[,,,"",\,<,>,']"
    
    For lngCount = 1 To Len(strValue)
        strCheck = Mid(strValue, lngCount, 1)
        '1バイト文字の場合にチェック
        If LenB(StrConv(strCheck, vbFromUnicode)) = 1 Then
            If strCheck Like PC_KINSOKU_STRING Then
                Exit Function
            End If
        End If
    Next

    fncChkKinsiMoji = True

End Function

'*< TOSTEM >*******************************************************************'
'*
'*【プロシージャ名】
'*!     fncChkKinsiMojiReceipt
'*
'*【処理概要】
'*?     禁止文字チェック
'*
'*【パラメータ】
'*      Name                   Type                Comment
'*P     strValue               String              値
'*P     strErrMsg              String              エラーメッセージ
'*【戻り値】
'*      Name                   Type                Comment
'*R     fncChkHankaku          Boolean;
'*
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)            EDP-No      説明
'*M 2014/05/01; 和田(LIS);              #-#####;    新規作成;
'*
'*< Comment End >**************************************************************
Private Function fncChkKinsiMojiReceipt(ByVal strValue As String) As Boolean
        
    Dim lngCount As Long
    Dim strCheck As String
    
    fncChkKinsiMojiReceipt = False
    
    Const PC_KINSOKU_STRING As String = "[ПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюя─│┌┐┘└├┬┤┴┼━┃┏┓┛┗┣┳┫┻╋┠┯┨┷┿┝┰┥┸╂①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ㍉㌔㌢㍍㌘㌧㌃㌶㍑㍗㌍㌦㌣㌫㍊㌻㎜㎝㎞㎎㎏㏄㎡㍻〝〟№㏍℡㊤㊥㊦㊧㊨㈱㈲㈹㍾㍽㍼≒≡∫∮∑√⊥∠∟⊿∵∩∪]"
    
    For lngCount = 1 To Len(strValue)
        strCheck = Mid(strValue, lngCount, 1)
        If strCheck Like PC_KINSOKU_STRING Then
            Exit Function
        End If
    Next

    fncChkKinsiMojiReceipt = True

End Function

'******************************************************************'
'*【プロシージャ名】
'*!    subIns_SetAllTables_Bun
'*【処理概要】
'*?    CSVファイルのパラメータに対して、DBにアクセスしてチェックを行います。
'*?    (分類情報版)
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    subIns_SetAllTables_Bun; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/09/05; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
Private Sub subIns_SetAllTables_Bun()

    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    
    '変数
    Dim strMethodCallInfo As String
    Dim strWhereSQL As String               '条件作成用
    Dim strSQL  As String                   'RS作成用SQL
    Dim lstrNow As Object                  '当日日付格納変数
    Dim i As Long
    
    'レコードセット用変数
    Dim iRecord As Long
    Dim rstHyoujiSiyou As ADODB.Recordset   '表示/仕様情報
    Dim rstWfKanri As ADODB.Recordset       'WF管理情報
    Dim rstRiekiKanri As ADODB.Recordset    '利益管理情報
    Dim rstIppanKihon As ADODB.Recordset    '一般基本情報
    
    'WF更新用フラグ(WF更新時の効率化の為Stringで"0"又は"1"を持つ)
    Dim strRiekiKanriAdded As String        '利益管理RSが更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    Dim strHyoujiAdded As String            '表示/仕様RSが更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    
    '日付を始めに取得
    lstrNow = Format(Now(), "yyyy/mm/dd hh:nn:ss")
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun()"

    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    '---------------
    '   内部処理
    '---------------
    On Error GoTo DefectErrorHandler

    ' ************** RS作成用SQLの条件部分（企画種類コード、年連番、枝番の指定） **************
    strWhereSQL = "WHERE kikaku_syurui_cd = '" & pstrKikakuSyurui & "'" & vbCrLf
    strWhereSQL = strWhereSQL & "AND nen_ren_no ='" & pstrKikakuNenrenban & "'" & vbCrLf
    strWhereSQL = strWhereSQL & "AND eda_no ='" & pstrKikakuEdaban & "'" & vbCrLf
    
    ' ************** 表示/仕様情報RS作成 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(表示/仕様情報RS作成)"
    
    strSQL = "SELECT * FROM t_syousen_hyouji_siyou with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "

        rstHyoujiSiyou = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstHyoujiSiyou.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    
    ' ************** WF管理情報RS作成 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(WF管理情報RS作成)"
    strSQL = "SELECT * FROM t_syousen_wf_kanri with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstWfKanri = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstWfKanri.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
         
    ' ************** 利益管理情報RS作成 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(WF管理情報RS作成)"
    strSQL = "SELECT * FROM t_syousen_rieki_kanri WITH(UPDLOCK)" & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstRiekiKanri = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstRiekiKanri.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    
    ' ************** VMST一般基本情報RS作成 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(VMST一般基本情報RS作成)"
    strSQL = "SELECT vk1.jan_cd, vik.* FROM v_ippan_kihon as vik with (READCOMMITTED) " & vbCrLf
    strSQL = strSQL & " INNER JOIN v_key_1 as vk1 ON (vik.key_cd = vk1.key_cd) " & vbCrLf
    strSQL = strSQL & " WHERE vk1.jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstIppanKihon = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstIppanKihon.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    
    ' ************** 1レコードずつRSを更新 *******************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(1レコードずつRSを更新 開始)"
    
    For iRecord = LBound(paryCsvData) To UBound(paryCsvData)
    
        'WF更新用フラグの初期化
        strRiekiKanriAdded = 0
        strHyoujiAdded = 0
        
        'データ更新対象店コードの取得
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(一般基本情報の取得)"
        
        'VMST一般基本情報の取得
        With rstIppanKihon
            If Not .BOF Then
                .MoveFirst
                .Find ("jan_cd = '" & paryCsvData(iRecord)(enmBunColNo.janCD) & "'")
            End If
        End With
        
        'データ更新対象店コードの取得
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(表示/仕様情報の取得)"
        
        '表示/仕様情報
        With rstHyoujiSiyou
            '表示/仕様情報の取得
            If Not .BOF Then
                .MoveFirst
                .Find ("jan_cd = '" & paryCsvData(iRecord)(enmBunColNo.janCD) & "'")
            End If
            
            'データ更新対象店コードの取得
            'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(表示/仕様情報の更新)"
        
            If .EOF Then
                '表示/仕様情報の追加
                .AddNew
                .Fields("kikaku_syurui_cd").Value = pstrKikakuSyurui
                .Fields("nen_ren_no").Value = pstrKikakuNenrenban
                .Fields("eda_no").Value = pstrKikakuEdaban
                .Fields("jan_cd").Value = paryCsvData(iRecord)(enmBunColNo.janCD)
                .Fields("hyoujiyou_syouhin_mei_kana").Value = rstIppanKihon.Fields("hyoujiyou_syouhin_mei_kana").Value
                .Fields("hyoujiyou_syouhin_mei_kanji").Value = rstIppanKihon.Fields("hyoujiyou_syouhin_mei_kanji").Value
                .Fields("receipt_mei_kana").Value = rstIppanKihon.Fields("receipt_mei_kana").Value
                .Fields("receipt_mei_kanji").Value = rstIppanKihon.Fields("receipt_mei_kanji").Value
                .Fields("syouhin_mei_kanji").Value = rstIppanKihon.Fields("syouhin_mei_kanji").Value
                .Fields("maker_mei_kanji").Value = rstIppanKihon.Fields("maker_mei_kanji").Value
                .Fields("kikaku_size_kanji").Value = rstIppanKihon.Fields("kikaku_size_kanji").Value
                .Fields("sina_no_kata_no_kanji").Value = rstIppanKihon.Fields("sina_no_kata_no_kanji").Value
                .Fields("iro_gara_kanji").Value = rstIppanKihon.Fields("iro_gara_kanji").Value
                .Fields("taste_kanji").Value = rstIppanKihon.Fields("taste_kanji").Value
                .Fields("seisanti_mei_kanji").Value = rstIppanKihon.Fields("seisanti_mei_kanji").Value
                .Fields("tokutyou").Value = rstIppanKihon.Fields("tokutyou").Value
                .Fields("syoumi_siyoukigen_um_kbn").Value = rstIppanKihon.Fields("syoumi_siyoukigen_um_kbn").Value
                .Fields("eco_syouhin_kbn").Value = rstIppanKihon.Fields("eco_syouhin_kbn").Value
                .Fields("green_tekigou_syouhin_kbn").Value = rstIppanKihon.Fields("green_tekigou_syouhin_kbn").Value
                .Fields("gpn_db_keisai_syouhin_kbn").Value = rstIppanKihon.Fields("gpn_db_keisai_syouhin_kbn").Value
                .Fields("touroku_user_id").Value = pstrTantousyaId
                .Fields("touroku_ymd_hms").Value = lstrNow
                .Update
    
                strHyoujiAdded = 1
            End If
        End With
        
        'データ更新対象店コードの取得
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(利益管理情報の取得)"
        
        '利益管理情報
        '利益管理テーブルに引数のJANコードのデータが存在した場合は削除する。
        With rstRiekiKanri
            If Not .BOF Then
                .MoveFirst
                .Filter = "jan_cd = '" & paryCsvData(iRecord)(enmBunColNo.janCD) & "'"
                If .RecordCount <> 0 Then
                    .Delete
                End If
                .Filter = adFilterNone
            End If
        End With
        
        'データ更新対象店コードの取得
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(品種表示情報の取得)"
        
        '品種表示テーブルに引数のKEY分類コード１、KEY分類コード２、KEY分類コード３より
        '紐付く情報を取得する。
        With pChkBunruiCodeRS
            If Not .BOF Then
                .MoveFirst
                .Filter = "key_bnrui_cd_1 = '" & paryCsvData(iRecord)(enmBunColNo.Bunrui1) & "' " _
                        & "AND key_bnrui_cd_2 = '" & paryCsvData(iRecord)(enmBunColNo.Bunrui2) & "'" _
                        & "AND key_bnrui_cd_3 = '" & paryCsvData(iRecord)(enmBunColNo.Bunrui3) & "'"
            End If
        End With
        
        'データ更新対象店コードの取得
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(利益管理情報の更新)"
        
        '利益管理テーブルにデータを追加する。
        With rstRiekiKanri
            .AddNew
            .Fields("kikaku_syurui_cd") = pstrKikakuSyurui
            .Fields("nen_ren_no") = pstrKikakuNenrenban
            .Fields("eda_no") = pstrKikakuEdaban
            .Fields("jan_cd") = paryCsvData(iRecord)(enmBunColNo.janCD)
            .Fields("henkou_date") = Format$(paryCsvData(iRecord)(enmBunColNo.HenkouDate), "@@@@/@@/@@")
            .Fields("key_bnrui_cd_1") = paryCsvData(iRecord)(enmBunColNo.Bunrui1)
            .Fields("key_bnrui_cd_2") = paryCsvData(iRecord)(enmBunColNo.Bunrui2)
            .Fields("key_bnrui_cd_3") = paryCsvData(iRecord)(enmBunColNo.Bunrui3)
            .Fields("uriba_cd") = pChkBunruiCodeRS.Fields("uriba_cd").Value
            .Fields("bumon_cd") = pChkBunruiCodeRS.Fields("old_bumon_cd").Value
            .Fields("tyuubunrui_cd") = pChkBunruiCodeRS.Fields("old_tyuubunrui_cd").Value
            .Fields("syoubunrui_cd") = pChkBunruiCodeRS.Fields("old_syoubunrui_cd").Value
            .Fields("hinsyu_cd") = pChkBunruiCodeRS.Fields("bunrui_cd_3").Value
            .Fields("touroku_user_id") = pstrTantousyaId
            .Fields("touroku_ymd_hms") = Format(Now, "yyyy/mm/dd hh:nn:ss")
            .Update
        End With
        
        strRiekiKanriAdded = 1
        
        '品種表示テーブルのFilterを解除する。
        pChkBunruiCodeRS.Filter = adFilterNone


        'データ更新対象店コードの取得
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(WF管理情報の更新)"
        
        'WF管理項目を最新情報に更新する
        With rstWfKanri
            If Not .BOF Then
                .MoveFirst
                .Find ("jan_cd = '" & paryCsvData(iRecord)(enmBunColNo.janCD) & "'")
            End If
            'WF管理項目に無い場合は、新規追加
            If .EOF Then
                .AddNew
                .Fields("kikaku_syurui_cd").Value = pstrKikakuSyurui
                .Fields("nen_ren_no").Value = pstrKikakuNenrenban
                .Fields("eda_no").Value = pstrKikakuEdaban
                .Fields("jan_cd").Value = paryCsvData(iRecord)(enmBunColNo.janCD)
                .Fields("jan_umu_kbn").Value = 1
                .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                .Fields("jyogai_kbn").Value = 0
                .Fields("touroku_tantousya").Value = pstrTantousyaId
                .Fields("keikakusuu").Value = 0
                .Fields("kikaku_no").Value = pstrKikakuSyurui & pstrKikakuNenrenban & "-" & pstrKikakuEdaban
                .Fields("kikaku_kahi_kbn").Value = 0
                .Fields("kikaku_syouninsya_comment").Value = dbNull.value
                .Fields("hinsyu_kahi_kbn").Value = 0
                .Fields("hinsyu_syouninsya_comment").Value = dbNull.value
                .Fields("naibu_cd").Value = rstIppanKihon.Fields("key_cd").Value
                .Fields("hyouji_siyou_touroku_kbn").Value = strHyoujiAdded
                .Fields("sekisai_size_touroku_kbn").Value = 0
                .Fields("pop_touroku_kbn").Value = 0
                .Fields("syanai_kihon_touroku_kbn").Value = 0
                .Fields("kowake_touroku_kbn").Value = 0
                .Fields("jigyoubetu_touroku_kbn").Value = 0
                .Fields("hanbai_touroku_kbn").Value = 0
                .Fields("siire_touroku_kbn").Value = 0
                .Fields("hinsyu_taikei_touroku_kbn").Value = strRiekiKanriAdded
                .Fields("uriba_touroku_kbn").Value = 0
                .Fields("baika_henkou_touroku_kbn").Value = 0
                .Fields("siire_henkou_touroku_kbn").Value = 0
                .Fields("center_cd_henkou_touroku_kbn").Value = 0
                .Fields("kowake_henkou_touroku_kbn").Value = 0
                .Fields("hanki_henkou_touroku_kbn").Value = 0
                .Fields("touroku_user_id").Value = pstrTantousyaId
                .Fields("kousin_user_id").Value = dbNull.value
                .Fields("touroku_ymd_hms").Value = lstrNow
                .Fields("kousin_ymd_hms").Value = dbNull.value
    
            'WF管理項目にある場合は、更新
            Else
                .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                .Fields("touroku_tantousya").Value = pstrTantousyaId
                .Fields("kikaku_no").Value = pstrKikakuSyurui & pstrKikakuNenrenban & "-" & pstrKikakuEdaban
                .Fields("naibu_cd").Value = rstIppanKihon.Fields("key_cd").Value
                
                If strHyoujiAdded = 0 Then
                Else
                    .Fields("hyouji_siyou_touroku_kbn").Value = 1
                End If
                
                If strRiekiKanriAdded = 0 Then
                Else
                    .Fields("hinsyu_taikei_touroku_kbn").Value = 1
                End If
                .Fields("kousin_user_id").Value = pstrTantousyaId
                .Fields("kousin_ymd_hms").Value = lstrNow
            End If
            .Update
        End With
    Next iRecord
    
    'DBの更新を行う。
    Call objCDBAccess.UpdateDCRecordset("VSWF", rstHyoujiSiyou)
    Call objCDBAccess.UpdateDCRecordset("VSWF", rstRiekiKanri)
    Call objCDBAccess.UpdateDCRecordset("VSWF", rstWfKanri)
    
    'RSのバッチ更新
    'GetObjectContext.SetComplete
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(オブジェクト開放)"
    
    '更新で使用したオブジェクトを開放する。
    If Not rstHyoujiSiyou Is Nothing Then
            rstHyoujiSiyou = Nothing
    End If

    If Not rstWfKanri Is Nothing Then
            rstWfKanri = Nothing
    End If

    If Not rstRiekiKanri Is Nothing Then
            rstRiekiKanri = Nothing
    End If

    If Not rstIppanKihon Is Nothing Then
            rstIppanKihon = Nothing
    End If
    
    'GetObjectContext.SetComplete
    
    Exit Sub

DefectErrorHandler:
    
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If

    If Not rstHyoujiSiyou Is Nothing Then
            rstHyoujiSiyou = Nothing
    End If

    If Not rstWfKanri Is Nothing Then
            rstWfKanri = Nothing
    End If

    If Not rstRiekiKanri Is Nothing Then
            rstRiekiKanri = Nothing
    End If

    If Not rstIppanKihon Is Nothing Then
            rstIppanKihon = Nothing
    End If
    
    If Not pChkvk1JANCodeRS Is Nothing Then
            pChkvk1JANCodeRS = Nothing
    End If
    
    If Not pChkBunruiCodeRS Is Nothing Then
            pChkBunruiCodeRS = Nothing
    End If
    
    'If Not GetObjectContext Is Nothing Then
        'GetObjectContext.SetAbort
    End If

    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
            'Err.Number, Err.Source, Err.Description
End Sub
'
'*< TOSTEM >*******************************************************************'
'*
'*【プロシージャ名】
'*!     FncSyoTouroku
'*
'*【処理概要】
'*?     商品基本情報一括変更のデータを商品選定WFのDBに取り込む
'*
'*【パラメータ】
'*      Name                   Type                Comment
'*P     strErrMsg              Object             エラーメッセージ;
'*【戻り値】
'*      Name                   Type                Comment
'*R     FncSyoTouroku          Boolean;
'*
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)            EDP-No      説明
'*M 2005/09/05; 高附 (IC);             P-24884;    新規作成;
'*M 2006/09/12; 和田(IC);              P-XXXXX;    商品選定ＷＦ問合せ撲滅改善;
'*
'*< Comment End >**************************************************************
Public Function FncSyoTouroku(ByRef strErrMsg) As Boolean
    
    Dim intN As Integer
    Dim blnRet As Boolean
    Dim strJanCd As String
    Dim intNowLine As Integer
    Dim strCSVRecord() As String
    Dim flgChkError As Boolean      'チェック時のエラーフラグ（TRUE:エラーあり）
    
    FncSyoTouroku = False
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    
    'DB参照エラーチェック用切断レコードセットの生成
    Dim strSQL          As String
    Dim objCDBAccess    As Object
        
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    '一括登録時、毎回一つユーザーのみ実行できる
    Call LockTableIkkatu(objCDBAccess)
    
    Call SetIkkatuLog("1-1", "商品基本情報一括変更のデータを商品選定WFのDBに取り込む", "FncSyoTouroku", "", "", pstrTantousyaId, "INS")
    Call SetIkkatuLog("1-1-1", "重複JANコードチェック用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")
    '重複JANコードチェック用切断レコードセット生成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(重複JANコードチェック用切断レコードセット生成)"
    strSQL = "SELECT twk.jan_cd, mk1.key_cd, twk.kikaku_no " & vbCrLf
    strSQL = strSQL & " FROM t_syousen_wf_kanri as twk WITH(readcommitted)" & vbCrLf
    strSQL = strSQL & "      INNER JOIN v_key_1 as mk1 WITH(readcommitted)" & vbCrLf
    strSQL = strSQL & "              ON(twk.jan_cd = mk1.jan_cd)" & vbCrLf
    strSQL = strSQL & " WHERE" & vbCrLf
    strSQL = strSQL & " NOT(" & vbCrLf
    strSQL = strSQL & "      kikaku_syurui_cd = '" & pstrKikakuSyurui & "' " & vbCrLf
    strSQL = strSQL & "      AND nen_ren_no = '" & pstrKikakuNenrenban & "' " & vbCrLf
    strSQL = strSQL & "      AND eda_no = '" & pstrKikakuEdaban & "'" & vbCrLf
    strSQL = strSQL & "    )" & vbCrLf
    strSQL = strSQL & " AND twk.touroku_kbn = '2'" & vbCrLf
    
        pChkSameJANCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pChkSameJANCodeRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    pChkSameJANCodeRS.Fields("key_cd").Properties("Optimize") = True 'オプティマイザ

        Call SetIkkatuLog("1-1-1", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-1-2", "CSVファイル内Janコードのリスト(Janコードの重複無し)作成", "SQL", "", "", pstrTantousyaId, "INS")
    'CSVファイル内Janコードのリスト(Janコードの重複無し)作成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(CSVファイル内Janコードのリスト(Janコードの重複無し)作成)"

    intNowLine = 0
    Do Until UBound(paryCsvData) < intNowLine
        
        strCSVRecord = paryCsvData(intNowLine)
         
        If InStr(pCsvJanCodeList, "'" & strCSVRecord(enmSyoColNo.janCD) & "'") = 0 Then
            pCsvJanCodeList = pCsvJanCodeList & "'" & strCSVRecord(enmSyoColNo.janCD) & "',"
        End If

        intNowLine = intNowLine + 1
    Loop
    
    If pCsvJanCodeList <> "" Then
        pCsvJanCodeList = Left(pCsvJanCodeList, Len(pCsvJanCodeList) - 1)
    Else
        pCsvJanCodeList = "''"
    End If
    intNowLine = 0
        Call SetIkkatuLog("1-1-2", "", "", "", "OK", "", "UPD")
        Call SetIkkatuLog("1-1-3", "統合マスタ登録済みJANコードチェック用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")
    '統合マスタ登録済みJANコードチェック用切断レコードセット生成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(統合マスタ登録済みJANコードチェック用切断レコードセット生成)"

    strSQL = "SELECT jan_cd, key_cd FROM v_key_1 WITH(readcommitted)"
    strSQL = strSQL & " WHERE jan_cd IN(" & pCsvJanCodeList & ") "
    
        pChkvk1JANCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pChkvk1JANCodeRS.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
        Call SetIkkatuLog("1-1-3", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-1-4", "CSVファイル内JANコード紐付け主キーのリスト(主キー重複無し)作成", "SQL", "", "", pstrTantousyaId, "INS")
    '統合マスタ登録済みJANコード紐付け主キーのリスト(主キー重複無し)作成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(CSVファイル内JANコード紐付け主キーのリスト(主キー重複無し)作成)"
    With pChkvk1JANCodeRS
        Do Until .EOF
            If InStr(pHKeyCDList, .Fields("key_cd")) = 0 Then
                pHKeyCDList = pHKeyCDList & "'" & .Fields("key_cd") & "',"
            End If
            .MoveNext
        Loop
    
        If pHKeyCDList <> "" Then
            pHKeyCDList = Left(pHKeyCDList, Len(pHKeyCDList) - 1)
        Else
            pHKeyCDList = "''"
        End If
    End With
        Call SetIkkatuLog("1-1-4", "", "", "", "OK", "", "UPD")
    'CSVファイル内需要コードのリスト(需要コードの重複無し)作成
'    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(CSVファイル内需要コードのリスト(需要コードの重複無し)作成)"
'
'    intNowLine = 0
'    Do Until UBound(paryCsvData) < intNowLine
'
'        strCSVRecord = paryCsvData(intNowLine)
'        If strCSVRecord(enmSyoColNo.VhJyuyouCd) <> "" Or strCSVRecord(enmSyoColNo.RdJyuyouCd) <> "" _
'            Or strCSVRecord(enmSyoColNo.VcsJyuyouCd) <> "" Or strCSVRecord(enmSyoColNo.SvhJyuyouCd) <> "" Then
'
'            If InStr(pCsvJyuyouCodeList, "'" & strCSVRecord(enmSyoColNo.VhJyuyouCd) & "'") = 0 Then
'                pCsvJyuyouCodeList = pCsvJyuyouCodeList & "'" & strCSVRecord(enmSyoColNo.VhJyuyouCd) & "',"
'            End If
'            If InStr(pCsvJyuyouCodeList, "'" & strCSVRecord(enmSyoColNo.RdJyuyouCd) & "'") = 0 Then
'                pCsvJyuyouCodeList = pCsvJyuyouCodeList & "'" & strCSVRecord(enmSyoColNo.RdJyuyouCd) & "',"
'            End If
'            If InStr(pCsvJyuyouCodeList, "'" & strCSVRecord(enmSyoColNo.VcsJyuyouCd) & "'") = 0 Then
'                pCsvJyuyouCodeList = pCsvJyuyouCodeList & "'" & strCSVRecord(enmSyoColNo.VcsJyuyouCd) & "',"
'            End If
'            If InStr(pCsvJyuyouCodeList, "'" & strCSVRecord(enmSyoColNo.SvhJyuyouCd) & "'") = 0 Then
'                pCsvJyuyouCodeList = pCsvJyuyouCodeList & "'" & strCSVRecord(enmSyoColNo.SvhJyuyouCd) & "',"
'            End If
'
'        End If
'
'        intNowLine = intNowLine + 1
'    Loop
'
'    If pCsvJyuyouCodeList <> "" Then
'        pCsvJyuyouCodeList = Left(pCsvJyuyouCodeList, Len(pCsvJyuyouCodeList) - 1)
'    Else
'        pCsvJyuyouCodeList = "''"
'    End If
'    intNowLine = 0
        Call SetIkkatuLog("1-1-5", "分類マスタKEYチェック用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")
    '売場体系データ登録済み需要コードチェック用切断レコードセット生成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(分類マスタKEYチェック用切断レコードセット生成)"

    strSQL = "SELECT jigyou_cd, jyuyou_cd"
    strSQL = strSQL & " FROM v_uriba_taikei_hyouji WITH(readcommitted)"
    'strSQL = strSQL & " WHERE jyuyou_cd IN(" & pCsvJyuyouCodeList & ") "
    
        pChkJyuyouCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    With pChkJyuyouCodeRS
            .Fields("jigyou_cd").Properties("Optimize") = True 'オプティマイザ
            .Fields("jyuyou_cd").Properties("Optimize") = True 'オプティマイザ
    End With
        Call SetIkkatuLog("1-1-5", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-1-6", "扱いパターンコードチェック用切断レコードセット生成", "SQL", "", "", pstrTantousyaId, "INS")
    '扱いパターンコードチェック用切断レコードセット生成
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(扱いパターンコードチェック用切断レコードセット生成)"

    strSQL = "SELECT atukai_ptn_cd "
    strSQL = strSQL & " FROM v_atukai_ptn WITH(READCOMMITTED)"
    
        pChkAtukaiPtnCdRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    With pChkAtukaiPtnCdRS
            .Fields("atukai_ptn_cd").Properties("Optimize") = True 'オプティマイザ
    End With
        Call SetIkkatuLog("1-1-6", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-1-7", "全ての配列データのチェックを行う。", "SQL", "", "", pstrTantousyaId, "INS")
    '全ての配列データのチェックを行う。
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(全ての配列データのチェックを行う。)"
    
    flgChkError = False
    intNowLine = 0
    strErrMsg = ""
    Do Until UBound(paryCsvData) < intNowLine
        
        strCSVRecord = paryCsvData(intNowLine)
        
        '(1)DBに依存しないチェック：FncCheckParameter_CorrectForm_Syo
        '(2)DBを参照してチェック:FncCheckParameter_DataBase_Syo
        If FncCheckParameter_CorrectForm_Syo(strCSVRecord, intNowLine + 1) = False Or _
            FncCheckParameter_DataBase_Syo(strCSVRecord, intNowLine + 1) = False Then
            flgChkError = True
        End If
        
        intNowLine = intNowLine + 1
    Loop
    Call SetIkkatuLog("1-1-7", "", "", "", "OK", "", "UPD")
    
    Call SetIkkatuLog("1-1-8", "DB参照エラーチェック用切断レコードセットの開放", "SQL", "", "", pstrTantousyaId, "INS")
    'DB参照エラーチェック用切断レコードセットの開放
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(DB参照エラーチェック用切断レコードセットの開放)"
    
    'チェックで使用したオブジェクトを開放する。
    If Not pChkSameJANCodeRS Is Nothing Then
        pChkSameJANCodeRS.Close
            pChkSameJANCodeRS = Nothing
    End If
        
    'データに１つでも問題がある場合は、ここで終了！！！
    If flgChkError = True Then
        'エラーがあればエラーメッセージを取得する。
        strErrMsg = FncGetErrorMsg_Syo()
        
        Exit Function
    End If
    Call SetIkkatuLog("1-1-8", "", "", "", "OK", "", "UPD")
    
    Call SetIkkatuLog("1-2", "データ更新処理", "subIns_SetAllTables_Syo", "", "", pstrTantousyaId, "INS")
    
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(データ更新処理開始)"
    
    '全てのテーブルを更新する。
    Call subIns_SetAllTables_Syo
    Call SetIkkatuLog("1-2", "", "", "", "OK", "", "UPD")
    FncSyoTouroku = True
    
    'オブジェクトの開放
    If Not pChkvk1JANCodeRS Is Nothing Then
        pChkvk1JANCodeRS.Close
            pChkvk1JANCodeRS = Nothing
    End If
    
    If Not pChkJyuyouCodeRS Is Nothing Then
        pChkJyuyouCodeRS.Close
            pChkJyuyouCodeRS = Nothing
    End If
    
    If Not pChkAtukaiPtnCdRS Is Nothing Then
        pChkAtukaiPtnCdRS.Close
            pChkAtukaiPtnCdRS = Nothing
    End If
    
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
        Call SetIkkatuLog("1-1", "", "", "", "OK", "", "UPD")

Exit Function
DefectErrorHandler:
    
    'オブジェクトの開放
    If Not pChkSameJANCodeRS Is Nothing Then
        pChkSameJANCodeRS.Close
            pChkSameJANCodeRS = Nothing
    End If
    
    If Not pChkvk1JANCodeRS Is Nothing Then
        pChkvk1JANCodeRS.Close
            pChkvk1JANCodeRS = Nothing
    End If
    
    If Not pChkJyuyouCodeRS Is Nothing Then
        pChkJyuyouCodeRS.Close
            pChkJyuyouCodeRS = Nothing
    End If
    
    If Not pChkAtukaiPtnCdRS Is Nothing Then
        pChkAtukaiPtnCdRS.Close
            pChkAtukaiPtnCdRS = Nothing
    End If
    
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
    
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    Call SetIkkatuLog("1-1", "", "", Left(Err.Description, 400), "NG", "", "UPD")
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", 'Err.Number, Err.Source, Err.Description

End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncCheckParameter_CorrectForm_Syo
'*【処理概要】
'*?    CSVファイルのパラメータに対して、DBに依存しない形式チェックを行います。
'*?    （商品基本情報一括変更用チェック）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckParameter_CorrectForm_Syo; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/09/05; 高附　周造(IC); 新規作成;
'*M    2006/06/15; 和田(IC);       P-30109; ＷＦ仕様改善簡易一括機能の追加;
'*< Comment End >**************************************************'
'
Private Function FncCheckParameter_CorrectForm_Syo(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheckParameter_CorrectForm_Syo = False
    
    Dim blnIsOKRequired
    Dim blnIsOKNumeric
    Dim blnIsOKNarrow
    Dim blnIsOKKinsi
    Dim blnIsOKWide
    Dim blnIsOKLimit
    Dim blnIsOKCheckDigit
    
    'a.必須項目（JANコード）が入っているか確認する。
    blnIsOKRequired = FncCheck_RequiredParameter_Syo(arryCSVRecord(), intNowLine)
    
    'b.JANコード、扱いパターン、標準売価、仕切原価、通常価格、ﾒｰｶｰ希望小売価格において、半角数字であるか確認する。
    blnIsOKNumeric = FncCheck_IsParameterNumeric_Syo(arryCSVRecord(), intNowLine)
    
    'c1.カナ商品名(TA伝票用)、レシート名(ｶﾅ)、
    '  VH需要コード、R&D需要コード、VCS需要コード、SVH需要コードにおいて、半角文字であるか確認する。
    blnIsOKNarrow = FncCheck_IsParameterNarrow_Syo(arryCSVRecord(), intNowLine)
    
    'c2.カナ商品名(TA伝票用)、レシート名(ｶﾅ)、POPレイアウト１、POPレイアウト２、POPサイズにおいて、半角禁止文字であるか確認する。
    blnIsOKKinsi = FncCheck_IsParameterKinsi_Syo(arryCSVRecord(), intNowLine)
    
    'd.漢字商品名(SA/ﾌﾟﾗｲｽｶｰﾄﾞ)、レシート名(漢字)、商品名(販促用)、メーカー名、規格/サイズ、品番/型番、色/柄、
    'テースト､生産地名､特徴､商品コピーにおいて全角文字であるか確認する｡
    blnIsOKWide = FncCheck_IsParameterWide_Syo(arryCSVRecord(), intNowLine)

    'e.桁数が制限を超えていないか確認する。
    blnIsOKLimit = FncCheck_Limit_Syo(arryCSVRecord(), intNowLine)

    'f.JANコードに正しいチェックデジットの値がセットされているかを調べる
    blnIsOKCheckDigit = FncCheck_CheckDigit(arryCSVRecord(enmSyoColNo.janCD), intNowLine, 4)
    
    If blnIsOKRequired And blnIsOKNumeric And blnIsOKNarrow And blnIsOKKinsi And _
       blnIsOKWide And blnIsOKLimit And blnIsOKCheckDigit Then
       
       FncCheckParameter_CorrectForm_Syo = True
    End If
    
End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncCheck_RequiredParameter_Syo
'*【処理概要】
'*?    CSVデータの１行に対し、必須項目が空白でないかどうかチェックします。
'*?    （商品基本情報一括変更用チェック）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_RequiredParameter_Syo; Boolean;チェックOK:True,NG:False
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/09/05; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'

Private Function FncCheck_RequiredParameter_Syo(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_RequiredParameter_Syo = True
    
    If arryCSVRecord(enmSyoColNo.janCD) = "" Then        'JANコード
        FncSetError (intNowLine, enmSyoColNo.janCD, PARAM_EMPTY
        FncCheck_RequiredParameter_Syo = False
    End If

End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncCheck_IsParameterNumeric_Syo
'*【処理概要】
'*?    CSVデータの１行に対し、全フィールドが半角数値かどうかチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_IsParameterNumeric_Syo; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/09/05; 高附　周造(IC); 新規作成;
'*M    2006/06/15; 和田(IC);       P-30109 ＷＦ仕様改善簡易一括機能の追加;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterNumeric_Syo(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_IsParameterNumeric_Syo = True
    
    If FncChkHankakuNumber(arryCSVRecord(enmSyoColNo.janCD)) = False Then     'JANコード
        FncSetError (intNowLine, enmSyoColNo.janCD, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Syo = False
    End If
    
    If arryCSVRecord(enmSyoColNo.TuujyouKakaku) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmSyoColNo.TuujyouKakaku)) = False Then     '通常価格
            FncSetError (intNowLine, enmSyoColNo.TuujyouKakaku, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Syo = False
        End If
    End If
    
    If arryCSVRecord(enmSyoColNo.MakerKakaku) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmSyoColNo.MakerKakaku)) = False Then     'ﾒｰｶｰ希望小売価格
            FncSetError (intNowLine, enmSyoColNo.MakerKakaku, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Syo = False
        End If
    End If
    
    If arryCSVRecord(enmSyoColNo.AtukaiPtn) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmSyoColNo.AtukaiPtn)) = False Then     '扱いパターン
            FncSetError (intNowLine, enmSyoColNo.AtukaiPtn, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Syo = False
        End If
    End If
    
    If arryCSVRecord(enmSyoColNo.SikiriGenka) <> "" Then
        If IsNumeric(arryCSVRecord(enmSyoColNo.SikiriGenka)) = False Then   '仕切原価
            FncSetError (intNowLine, enmSyoColNo.SikiriGenka, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Syo = False
        End If
    End If
    
    If arryCSVRecord(enmSyoColNo.HyoujyunBaika) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmSyoColNo.HyoujyunBaika)) = False Then   '標準売価
            FncSetError (intNowLine, enmSyoColNo.HyoujyunBaika, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Syo = False
        End If
    End If
    
End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncCheck_IsParameterNarrow_Syo
'*【処理概要】
'*?    CSVデータの１行に対し、対象フィールドが半角文字かどうかチェックします。
'*?    （商品基本情報一括変更用チェック）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_IsParameterNarrow_Syo; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/09/05; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterNarrow_Syo(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_IsParameterNarrow_Syo = True

    If FncChkHankakuNarrow(arryCSVRecord(enmSyoColNo.SyouhinmeiKana)) = False Then      'カナ商品名(TA伝票用)
        FncSetError (intNowLine, enmSyoColNo.SyouhinmeiKana, PARAM_NONARROW
        FncCheck_IsParameterNarrow_Syo = False
    End If
    
    If FncChkHankakuNarrow(arryCSVRecord(enmSyoColNo.ReceiptmeiKana)) = False Then      'レシート名(ｶﾅ)
        FncSetError (intNowLine, enmSyoColNo.ReceiptmeiKana, PARAM_NONARROW
        FncCheck_IsParameterNarrow_Syo = False
    End If
    
    If fncChkHankakuEisuu(arryCSVRecord(enmSyoColNo.VhJyuyouCd)) = False Then      'VH需要コード
        FncSetError (intNowLine, enmSyoColNo.VhJyuyouCd, PARAM_HANEISUU
        FncCheck_IsParameterNarrow_Syo = False
    End If
    
    If fncChkHankakuEisuu(arryCSVRecord(enmSyoColNo.RdJyuyouCd)) = False Then      'R&D需要コード
        FncSetError (intNowLine, enmSyoColNo.RdJyuyouCd, PARAM_HANEISUU
        FncCheck_IsParameterNarrow_Syo = False
    End If
    
    If fncChkHankakuEisuu(arryCSVRecord(enmSyoColNo.VcsJyuyouCd)) = False Then      'VCS需要コード
        FncSetError (intNowLine, enmSyoColNo.VcsJyuyouCd, PARAM_HANEISUU
        FncCheck_IsParameterNarrow_Syo = False
    End If
    
    If fncChkHankakuEisuu(arryCSVRecord(enmSyoColNo.SvhJyuyouCd)) = False Then      'SVH需要コード
        FncSetError (intNowLine, enmSyoColNo.SvhJyuyouCd, PARAM_HANEISUU
        FncCheck_IsParameterNarrow_Syo = False
    End If
    
End Function
'******************************************************************'
'*【プロシージャ名】
'*!    FncCheck_IsParameterKinsi_Syo
'*【処理概要】
'*?    CSVデータの１行に対し、対象フィールドに半角禁止文字が含まれているかどうかチェックします。
'*?    （商品基本情報一括変更用チェック）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_IsParameterKinsi_Syo; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/09/20; 高附　周造(IC); 新規作成;
'*M    2006/06/15; 和田(IC);       P-30109 ＷＦ仕様改善簡易一括機能の追加;
'*M    2014/05/02; 和田(LIS);      P-xxxxx レシート名(漢字)禁止文字チェック追加;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterKinsi_Syo(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_IsParameterKinsi_Syo = True

    If fncChkKinsiMoji(arryCSVRecord(enmSyoColNo.SyouhinmeiKana)) = False Then      'カナ商品名(TA伝票用)
        FncSetError (intNowLine, enmSyoColNo.SyouhinmeiKana, PARAM_KINSIMOJI
        FncCheck_IsParameterKinsi_Syo = False
    End If
    
    '2014/05/02 add wadak7
    If fncChkKinsiMojiReceipt(arryCSVRecord(enmSyoColNo.ReceiptmeiKanji)) = False Then      'レシート名(漢字)
        FncSetError (intNowLine, enmSyoColNo.ReceiptmeiKanji, PARAM_KINSIMOJI_RECEIPT
        FncCheck_IsParameterKinsi_Syo = False
    End If
    
    If fncChkKinsiMoji(arryCSVRecord(enmSyoColNo.ReceiptmeiKana)) = False Then      'レシート名(ｶﾅ)
        FncSetError (intNowLine, enmSyoColNo.ReceiptmeiKana, PARAM_KINSIMOJI
        FncCheck_IsParameterKinsi_Syo = False
    End If
    
    If fncChkKinsiMoji(arryCSVRecord(enmSyoColNo.POPLayout1)) = False Then          'POPレイアウト１
        FncSetError (intNowLine, enmSyoColNo.POPLayout1, PARAM_KINSIMOJI
        FncCheck_IsParameterKinsi_Syo = False
    End If
    
    If fncChkKinsiMoji(arryCSVRecord(enmSyoColNo.POPLayout2)) = False Then          'POPレイアウト２
        FncSetError (intNowLine, enmSyoColNo.POPLayout2, PARAM_KINSIMOJI
        FncCheck_IsParameterKinsi_Syo = False
    End If
    
    If fncChkKinsiMoji(arryCSVRecord(enmSyoColNo.PopSize)) = False Then             'POPサイズ
        FncSetError (intNowLine, enmSyoColNo.PopSize, PARAM_KINSIMOJI
        FncCheck_IsParameterKinsi_Syo = False
    End If
    
End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncCheck_IsParameterWide_Syo
'*【処理概要】
'*?    CSVデータの１行に対し、フィールドが桁数制限を越えていないかチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_IsParameterWide_Syo; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/09/05; 高附　周造(IC); 新規作成;
'*M    2006/06/15; 和田(IC);       P-30109 ＷＦ仕様改善簡易一括機能の追加;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterWide_Syo(ByRef arryCSVRecord() As String, intNowLine As Integer)

    FncCheck_IsParameterWide_Syo = True
    
    '(a)漢字商品名(TA/ﾌﾟﾗｲｽｶｰﾄﾞ)
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.SyouhinmeiKanji)) = False Then
        FncSetError (intNowLine, enmSyoColNo.SyouhinmeiKanji, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    '(b)レシート名(漢字)
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.ReceiptmeiKanji)) = False Then
        FncSetError (intNowLine, enmSyoColNo.ReceiptmeiKanji, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    '(c)商品名(販促用)
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.SyouhihinmeiHansoku)) = False Then
        FncSetError (intNowLine, enmSyoColNo.SyouhihinmeiHansoku, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
    'メーカー名
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.MakerMei)) = False Then
        FncSetError (intNowLine, enmSyoColNo.MakerMei, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
    '規格/サイズ
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.KikakuSize)) = False Then
        FncSetError (intNowLine, enmSyoColNo.KikakuSize, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
    '品番/型番
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.HinbanKataban)) = False Then
        FncSetError (intNowLine, enmSyoColNo.HinbanKataban, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
    '色/柄
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.IroGara)) = False Then
        FncSetError (intNowLine, enmSyoColNo.IroGara, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
    'テースト
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.Taste)) = False Then
        FncSetError (intNowLine, enmSyoColNo.Taste, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
    '生産地名
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.Seisantimei)) = False Then
        FncSetError (intNowLine, enmSyoColNo.Seisantimei, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
    '特徴
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.Tokutyou)) = False Then
        FncSetError (intNowLine, enmSyoColNo.Tokutyou, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
    '商品コピー
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.SyouhinCopy)) = False Then
        FncSetError (intNowLine, enmSyoColNo.SyouhinCopy, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
End Function


'*< TOSTEM >*******************************************************************'
'*
'*【プロシージャ名】
'*!     fncChkHankakuEisuu
'*
'*【処理概要】
'*?     半角チェック(英数のみ) 全角文字はエラー
'*
'*【パラメータ】
'*      Name                   Type                Comment
'*P     strValue               String              値
'*【戻り値】
'*      Name                   Type                Comment
'*R     fncChkHankaku          Boolean;
'*
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)            EDP-No      説明
'*M 2005/09/20; 高附 (IC);             #-#####;    新規作成;
'*
'*< Comment End >**************************************************************
Private Function fncChkHankakuEisuu(ByVal strValue As String) As Boolean
        
    Dim lngCount As Long

    fncChkHankakuEisuu = False

    For lngCount = 1 To Len(strValue)
        If Not Mid(strValue, lngCount, 1) Like "[0-9a-zA-Z]" Then
            Exit Function
        End If
    Next

    fncChkHankakuEisuu = True

End Function
'******************************************************************'
'*【プロシージャ名】
'*!    FncCheck_Limit_Syo
'*【処理概要】
'*?    CSVデータの１行に対し、フィールドが桁数制限を越えていないかチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheck_Limit_Syo; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/08/31; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheck_Limit_Syo(ByRef arryCSVRecord() As String, intNowLine As Integer)

    FncCheck_Limit_Syo = True
    
    '(a)JANコード
    If FncCheck_JANCode(arryCSVRecord(enmSyoColNo.janCD)) = False Then
        FncSetError (intNowLine, enmSyoColNo.janCD, PARAM_SHORTDISIT)
        FncCheck_Limit_Syo = False
    End If
    '(b)カナ商品名(TA伝票用)
    If arryCSVRecord(enmSyoColNo.SyouhinmeiKana) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.SyouhinmeiKana), LENGTH_SYOKANA) = False Then
            FncSetError (intNowLine, enmSyoColNo.SyouhinmeiKana, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(c)漢字商品名(SA/ﾌﾟﾗｲｽｶｰﾄﾞ用)
    If arryCSVRecord(enmSyoColNo.SyouhinmeiKanji) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.SyouhinmeiKanji), LENGTH_SYOKANJI) = False Then
            FncSetError (intNowLine, enmSyoColNo.SyouhinmeiKanji, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(d)レシート名(ｶﾅ)
    If arryCSVRecord(enmSyoColNo.ReceiptmeiKana) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.ReceiptmeiKana), LENGTH_RESKANA) = False Then
            FncSetError (intNowLine, enmSyoColNo.ReceiptmeiKana, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(e)レシート名(漢字)
    If arryCSVRecord(enmSyoColNo.ReceiptmeiKanji) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.ReceiptmeiKanji), LENGTH_RESKANJI) = False Then
            FncSetError (intNowLine, enmSyoColNo.ReceiptmeiKanji, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(f)商品名(販促用)
    If arryCSVRecord(enmSyoColNo.SyouhihinmeiHansoku) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.SyouhihinmeiHansoku), LENGTH_SYOHANSOKU) = False Then
            FncSetError (intNowLine, enmSyoColNo.SyouhihinmeiHansoku, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
    
     'メーカー名
    If arryCSVRecord(enmSyoColNo.MakerMei) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.MakerMei), 10) = False Then
            FncSetError (intNowLine, enmSyoColNo.MakerMei, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     '規格/サイズ
    If arryCSVRecord(enmSyoColNo.KikakuSize) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.KikakuSize), 20) = False Then
            FncSetError (intNowLine, enmSyoColNo.KikakuSize, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     '品番/型番
    If arryCSVRecord(enmSyoColNo.HinbanKataban) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.HinbanKataban), 20) = False Then
            FncSetError (intNowLine, enmSyoColNo.HinbanKataban, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     '色/柄
    If arryCSVRecord(enmSyoColNo.IroGara) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.IroGara), 20) = False Then
            FncSetError (intNowLine, enmSyoColNo.IroGara, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     'テースト
    If arryCSVRecord(enmSyoColNo.Taste) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.Taste), 20) = False Then
            FncSetError (intNowLine, enmSyoColNo.Taste, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     '生産地名
    If arryCSVRecord(enmSyoColNo.Seisantimei) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.Seisantimei), 20) = False Then
            FncSetError (intNowLine, enmSyoColNo.Seisantimei, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     '特徴
    If arryCSVRecord(enmSyoColNo.Tokutyou) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.Tokutyou), 150) = False Then
            FncSetError (intNowLine, enmSyoColNo.Tokutyou, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     '商品コピー
    If arryCSVRecord(enmSyoColNo.SyouhinCopy) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.SyouhinCopy), 50) = False Then
            FncSetError (intNowLine, enmSyoColNo.SyouhinCopy, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     '通常価格
    If arryCSVRecord(enmSyoColNo.TuujyouKakaku) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.TuujyouKakaku), 7) = False Then
            FncSetError (intNowLine, enmSyoColNo.TuujyouKakaku, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     'メーカー希望小売価格
    If arryCSVRecord(enmSyoColNo.MakerKakaku) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.MakerKakaku), 7) = False Then
            FncSetError (intNowLine, enmSyoColNo.MakerKakaku, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     'POPレイアウト１
    If arryCSVRecord(enmSyoColNo.POPLayout1) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.POPLayout1), 30) = False Then
            FncSetError (intNowLine, enmSyoColNo.POPLayout1, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     'POPレイアウト２
    If arryCSVRecord(enmSyoColNo.POPLayout2) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.POPLayout2), 30) = False Then
            FncSetError (intNowLine, enmSyoColNo.POPLayout2, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     'POPサイズ
    If arryCSVRecord(enmSyoColNo.PopSize) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.PopSize), 30) = False Then
            FncSetError (intNowLine, enmSyoColNo.PopSize, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
    
    '(g)扱いパターン
    If arryCSVRecord(enmSyoColNo.AtukaiPtn) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.AtukaiPtn), LENGTH_ATUPTN) = False Then
            FncSetError (intNowLine, enmSyoColNo.AtukaiPtn, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(h)標準売価
    If arryCSVRecord(enmSyoColNo.HyoujyunBaika) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.HyoujyunBaika), LENGTH_BAIKA) = False Then
            FncSetError (intNowLine, enmSyoColNo.HyoujyunBaika, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(i)仕切原価
    If arryCSVRecord(enmSyoColNo.SikiriGenka) <> "" Then
        If FncCheck_Genka(arryCSVRecord(enmSyoColNo.SikiriGenka)) = False Then
            FncSetError (intNowLine, enmSyoColNo.SikiriGenka, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(j)VH需要コード
    If arryCSVRecord(enmSyoColNo.VhJyuyouCd) <> "" Then
        If FncCheck_Length_Kotei(arryCSVRecord(enmSyoColNo.VhJyuyouCd), LENGTH_JYUYOU) = False Then
            FncSetError (intNowLine, enmSyoColNo.VhJyuyouCd, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(k)R&D需要コード
    If arryCSVRecord(enmSyoColNo.RdJyuyouCd) <> "" Then
        If FncCheck_Length_Kotei(arryCSVRecord(enmSyoColNo.RdJyuyouCd), LENGTH_JYUYOU) = False Then
            FncSetError (intNowLine, enmSyoColNo.RdJyuyouCd, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(l)VCS需要コード
    If arryCSVRecord(enmSyoColNo.VcsJyuyouCd) <> "" Then
        If FncCheck_Length_Kotei(arryCSVRecord(enmSyoColNo.VcsJyuyouCd), LENGTH_JYUYOU) = False Then
            FncSetError (intNowLine, enmSyoColNo.VcsJyuyouCd, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(m)SVH需要コード
    If arryCSVRecord(enmSyoColNo.SvhJyuyouCd) <> "" Then
        If FncCheck_Length_Kotei(arryCSVRecord(enmSyoColNo.SvhJyuyouCd), LENGTH_JYUYOU) = False Then
            FncSetError (intNowLine, enmSyoColNo.SvhJyuyouCd, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
End Function

'******************************************************************'
'*【プロシージャ名】
'*!    FncCheckParameter_DataBase_Syo
'*【処理概要】
'*?    CSVファイルのパラメータに対して、DBにアクセスしてチェックを行います。
'*?    (分類情報版)
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckParameter_DataBase_Syo; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/08/31; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheckParameter_DataBase_Syo(ByRef arryCSVRecord() As String, intNowLine As Integer)

    Dim blnIsOKJANCode
    Dim blnIsOKNaibuCode
    Dim blnIsOKJyuyouCode
    Dim blnIsOKAtukaiPtnCd
    
    Dim strKeyCode As String
    
    FncCheckParameter_DataBase_Syo = "False"
    
    'a.JANコードが有効か確認する。
    blnIsOKJANCode = FncCheckDB_JANCode(arryCSVRecord(enmSyoColNo.janCD), strKeyCode, intNowLine, RET_FLD_SYO)
    
    'b.KEYコードが有効か確認する。
    blnIsOKNaibuCode = FncCheckDB_NaibuCode(strKeyCode, intNowLine, RET_FLD_SYO)
    
    'c.需要コードが有効か確認する。
    blnIsOKJyuyouCode = FncCheckDB_JyuyouCode(arryCSVRecord(), intNowLine, RET_FLD_SYO)
    
    'd.扱いパターンコードが有効か確認する。
    blnIsOKAtukaiPtnCd = FncCheckDB_AtukaiPtnCd(arryCSVRecord(enmSyoColNo.AtukaiPtn), intNowLine, RET_FLD_SYO)
    
    If blnIsOKJANCode And blnIsOKJyuyouCode And blnIsOKNaibuCode And blnIsOKAtukaiPtnCd Then
        FncCheckParameter_DataBase_Syo = "True"
    End If
       
End Function

'******************************************************************'
'*【プロシージャ名】
'*!    subIns_SetAllTables_Syo
'*【処理概要】
'*?    CSVファイルのパラメータに対して、DBにアクセスしてチェックを行います。
'*?    (分類情報版)
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intNowLine; Integer; CSVファイルの現在の行
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckParameter_DataBase_Syo; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/09/05; 高附　周造(IC); 新規作成;
'*M    2006/06/15; 和田(IC);       P-30109; ＷＦ仕様改善簡易一括機能の追加;
'*< Comment End >**************************************************'
Private Sub subIns_SetAllTables_Syo()

    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    
    '変数
    Dim strMethodCallInfo As String
    Dim strWhereSQL As String               '条件作成用
    Dim strSQL  As String                   'RS作成用SQL
    Dim lstrNow As Object                  '当日日付格納変数
    Dim i As Long
    
    'レコードセット用変数
    Dim iRecord As Long
    Dim rstHyoujiSiyou As ADODB.Recordset   '表示/仕様情報
    Dim rstWfKanri As ADODB.Recordset       'WF管理情報
    Dim rstHanbai As ADODB.Recordset        '販売情報
    Dim rstHanki As ADODB.Recordset         '半期変更情報
    Dim rstUriba As ADODB.Recordset         '売場情報
    Dim rstIppanKihon As ADODB.Recordset    'VMST一般基本情報
    Dim rstHankiHenkou As ADODB.Recordset   'VMST半期変更情報
    Dim rstPOP As ADODB.Recordset           'POP情報
    Dim rstSyanaiKihon As ADODB.Recordset   'VMST社内基本情報
    
    'WF更新用フラグ(WF更新時の効率化の為Stringで"0"又は"1"を持つ)
    Dim strHyoujiAdded As String            '売価変更RSが更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    Dim strHanbaiAdded As String            '仕入変更RSが更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    Dim strHankiAdded As String             'センターコード変更RSが更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    Dim strUribaAdded As String             'センターコード変更RSが更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    Dim strPOPAdded As String               'POP情報RSが更新(新規追加又は更新)されたか("0"：更新無し　"1"：更新在り)
    
    '日付を始めに取得
    lstrNow = Format(Now(), "yyyy/mm/dd hh:nn:ss")
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo()"

    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    '---------------
    '   内部処理
    '---------------
    On Error GoTo DefectErrorHandler

    ' ************** RS作成用SQLの条件部分（企画種類コード、年連番、枝番の指定） **************
    strWhereSQL = "WHERE kikaku_syurui_cd = '" & pstrKikakuSyurui & "'" & vbCrLf
    strWhereSQL = strWhereSQL & "AND nen_ren_no ='" & pstrKikakuNenrenban & "'" & vbCrLf
    strWhereSQL = strWhereSQL & "AND eda_no ='" & pstrKikakuEdaban & "'" & vbCrLf
    
    ' ************** 表示/仕様情報RS作成 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(表示/仕様情報RS作成)"
    
    strSQL = "SELECT * FROM t_syousen_hyouji_siyou with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "

        rstHyoujiSiyou = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstHyoujiSiyou.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    
    ' ************** WF管理情報RS作成 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(WF管理情報RS作成)"
    strSQL = "SELECT * FROM t_syousen_wf_kanri with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstWfKanri = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstWfKanri.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
         
    ' ************** 販売情報RS作成 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(販売情報RS作成)"
    strSQL = "SELECT * FROM t_syousen_hanbai with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstHanbai = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstHanbai.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    
    ' ************** 半期情報RS作成 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(半期情報RS作成)"
    strSQL = "SELECT * FROM t_syousen_hanki_henkou with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstHanki = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstHanki.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    
    ' ************** 売場情報RS作成 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(売場情報RS作成)"
    strSQL = "SELECT * FROM t_syousen_uriba with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstUriba = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstUriba.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
        
    ' ************** POP情報RS作成 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(POP情報RS作成)"
    strSQL = "SELECT * FROM t_syousen_pop with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstPOP = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstPOP.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
        
    ' ************** VMST一般基本情報RS作成 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(VMST一般基本情報RS作成)"
    strSQL = "SELECT vk1.jan_cd, vik.* FROM v_ippan_kihon as vik with (READCOMMITTED) " & vbCrLf
    strSQL = strSQL & " INNER JOIN v_key_1 as vk1 ON (vik.key_cd = vk1.key_cd) " & vbCrLf
    strSQL = strSQL & " WHERE vk1.jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstIppanKihon = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstIppanKihon.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    
    ' ************** VMST半期変更情報RS作成 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(VMST半期変更情報RS作成)"
    strSQL = "SELECT vk1.jan_cd, vhh.* FROM v_hanki_henkou as vhh with (READCOMMITTED) " & vbCrLf
    strSQL = strSQL & " INNER JOIN v_key_1 as vk1 with (READCOMMITTED) ON (vhh.key_cd = vk1.key_cd)" & vbCrLf
    strSQL = strSQL & " WHERE vk1.jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstHankiHenkou = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstHankiHenkou.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    
    ' ************** VMST社内基本情報RS作成 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(VMST社内基本情報RS作成)"
    strSQL = "SELECT vk1.jan_cd, vsk.* FROM v_syanai_kihon as vsk with (READCOMMITTED) " & vbCrLf
    strSQL = strSQL & " INNER JOIN v_key_1 as vk1 ON (vsk.key_cd = vk1.key_cd) " & vbCrLf
    strSQL = strSQL & " WHERE vk1.jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstSyanaiKihon = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstSyanaiKihon.Fields("jan_cd").Properties("Optimize") = True 'オプティマイザ
    
    
    ' ************** 1レコードずつRSを更新 *******************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(1レコードずつRSを更新 開始)"
    
    For iRecord = LBound(paryCsvData) To UBound(paryCsvData)
    
        'WF更新用フラグの初期化
        strHyoujiAdded = 0
        strHanbaiAdded = 0
        strHankiAdded = 0
        strUribaAdded = 0
        strPOPAdded = 0
        
        'データ更新対象店コードの取得
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(一般基本情報の取得)"
        
        'VMST一般基本情報の取得
        With rstIppanKihon
            If Not .BOF Then
                .MoveFirst
                .Find ("jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "'")
            End If
        End With
        
        '表示/仕様情報
        With rstHyoujiSiyou
             '表示/仕様情報の取得
             If Not .BOF Then
                 .MoveFirst
                 .Filter = "jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "' "
                 If .RecordCount <> 0 Then
                     .Delete
                 End If
             End If
         
             '表示/仕様情報の追加
             .AddNew
             .Fields("kikaku_syurui_cd").Value = pstrKikakuSyurui
             .Fields("nen_ren_no").Value = pstrKikakuNenrenban
             .Fields("eda_no").Value = pstrKikakuEdaban
             .Fields("jan_cd").Value = paryCsvData(iRecord)(enmSyoColNo.janCD)
             If paryCsvData(iRecord)(enmSyoColNo.SyouhinmeiKana) <> "" Then
                 .Fields("hyoujiyou_syouhin_mei_kana").Value = paryCsvData(iRecord)(enmSyoColNo.SyouhinmeiKana)
             Else
                 .Fields("hyoujiyou_syouhin_mei_kana").Value = rstIppanKihon.Fields("hyoujiyou_syouhin_mei_kana").Value
             End If
             
             If paryCsvData(iRecord)(enmSyoColNo.SyouhinmeiKanji) <> "" Then
                 .Fields("hyoujiyou_syouhin_mei_kanji").Value = paryCsvData(iRecord)(enmSyoColNo.SyouhinmeiKanji)
             Else
                 .Fields("hyoujiyou_syouhin_mei_kanji").Value = rstIppanKihon.Fields("hyoujiyou_syouhin_mei_kanji")
             End If
             
             If paryCsvData(iRecord)(enmSyoColNo.ReceiptmeiKana) <> "" Then
                 .Fields("receipt_mei_kana").Value = paryCsvData(iRecord)(enmSyoColNo.ReceiptmeiKana)
             Else
                 .Fields("receipt_mei_kana").Value = rstIppanKihon.Fields("receipt_mei_kana").Value
             End If
             
             If paryCsvData(iRecord)(enmSyoColNo.ReceiptmeiKanji) <> "" Then
                 .Fields("receipt_mei_kanji").Value = paryCsvData(iRecord)(enmSyoColNo.ReceiptmeiKanji)
             Else
                 .Fields("receipt_mei_kanji").Value = rstIppanKihon.Fields("receipt_mei_kanji").Value
             End If
             
             If paryCsvData(iRecord)(enmSyoColNo.SyouhihinmeiHansoku) <> "" Then
                 .Fields("syouhin_mei_kanji").Value = paryCsvData(iRecord)(enmSyoColNo.SyouhihinmeiHansoku)
             Else
                 .Fields("syouhin_mei_kanji").Value = rstIppanKihon.Fields("syouhin_mei_kanji").Value
             End If
             '▼メーカー名
             If paryCsvData(iRecord)(enmSyoColNo.MakerMei) <> "" Then
                 .Fields("maker_mei_kanji").Value = paryCsvData(iRecord)(enmSyoColNo.MakerMei)
             Else
                 .Fields("maker_mei_kanji").Value = rstIppanKihon.Fields("maker_mei_kanji").Value
             End If
             '▼規格/サイズ
             If paryCsvData(iRecord)(enmSyoColNo.KikakuSize) <> "" Then
                 .Fields("kikaku_size_kanji").Value = paryCsvData(iRecord)(enmSyoColNo.KikakuSize)
             Else
                 .Fields("kikaku_size_kanji").Value = rstIppanKihon.Fields("kikaku_size_kanji").Value
             End If
             '▼品番/型番
             If paryCsvData(iRecord)(enmSyoColNo.HinbanKataban) <> "" Then
                 .Fields("sina_no_kata_no_kanji").Value = paryCsvData(iRecord)(enmSyoColNo.HinbanKataban)
             Else
                 .Fields("sina_no_kata_no_kanji").Value = rstIppanKihon.Fields("sina_no_kata_no_kanji").Value
             End If
             '▼色/柄
             If paryCsvData(iRecord)(enmSyoColNo.IroGara) <> "" Then
                 .Fields("iro_gara_kanji").Value = paryCsvData(iRecord)(enmSyoColNo.IroGara)
             Else
                 .Fields("iro_gara_kanji").Value = rstIppanKihon.Fields("iro_gara_kanji").Value
             End If
             '▼テースト
             If paryCsvData(iRecord)(enmSyoColNo.Taste) <> "" Then
                 .Fields("taste_kanji").Value = paryCsvData(iRecord)(enmSyoColNo.Taste)
             Else
                 .Fields("taste_kanji").Value = rstIppanKihon.Fields("taste_kanji").Value
             End If
             '▼生産地名
             If paryCsvData(iRecord)(enmSyoColNo.Seisantimei) <> "" Then
                 .Fields("seisanti_mei_kanji").Value = paryCsvData(iRecord)(enmSyoColNo.Seisantimei)
             Else
                 .Fields("seisanti_mei_kanji").Value = rstIppanKihon.Fields("seisanti_mei_kanji").Value
             End If
             '▼特徴
             If paryCsvData(iRecord)(enmSyoColNo.Tokutyou) <> "" Then
                 .Fields("tokutyou").Value = paryCsvData(iRecord)(enmSyoColNo.Tokutyou)
             Else
                 .Fields("tokutyou").Value = rstIppanKihon.Fields("tokutyou").Value
             End If
             
             .Fields("syoumi_siyoukigen_um_kbn").Value = rstIppanKihon.Fields("syoumi_siyoukigen_um_kbn").Value
             .Fields("eco_syouhin_kbn").Value = rstIppanKihon.Fields("eco_syouhin_kbn").Value
             .Fields("green_tekigou_syouhin_kbn").Value = rstIppanKihon.Fields("green_tekigou_syouhin_kbn").Value
             .Fields("gpn_db_keisai_syouhin_kbn").Value = rstIppanKihon.Fields("gpn_db_keisai_syouhin_kbn").Value
             .Fields("touroku_user_id").Value = pstrTantousyaId
             .Fields("touroku_ymd_hms").Value = lstrNow
             .Update
             .Filter = adFilterNone

             strHyoujiAdded = 1
        End With
        
        '販売情報
        With rstHanbai
            If paryCsvData(iRecord)(enmSyoColNo.AtukaiPtn) <> "" Then
                '販売情報の取得
                If Not .BOF Then
                    .MoveFirst
                    .Filter = "jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "' "
                    If .RecordCount <> 0 Then
                        .Delete
                    End If
                End If
            
                '販売情報の追加
                .AddNew
                .Fields("kikaku_syurui_cd").Value = pstrKikakuSyurui
                .Fields("nen_ren_no").Value = pstrKikakuNenrenban
                .Fields("eda_no").Value = pstrKikakuEdaban
                .Fields("jan_cd").Value = paryCsvData(iRecord)(enmSyoColNo.janCD)
                .Fields("atukai_ptn_cd").Value = paryCsvData(iRecord)(enmSyoColNo.AtukaiPtn)
                .Fields("baika").Value = 0
                .Fields("hyoujyun_baika").Value = 0
                .Fields("touroku_user_id").Value = pstrTantousyaId
                .Fields("touroku_ymd_hms").Value = lstrNow
                .Update
                .Filter = adFilterNone
                strHanbaiAdded = 1
            End If
        End With
        
        
        'VMST半期変更情報の取得
        With rstHankiHenkou
            If Not .BOF Then
                .MoveFirst
                .Filter = "jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "' "
            End If
        
        End With
        
        '半期情報
        With rstHanki
            If paryCsvData(iRecord)(enmSyoColNo.HyoujyunBaika) <> "" _
                Or paryCsvData(iRecord)(enmSyoColNo.SikiriGenka) <> "" Then
                '半期情報の取得
                If Not .BOF Then
                    .MoveFirst
                    .Filter = "jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "' "
                    If .RecordCount <> 0 Then
                        .Delete
                    End If
                End If
                
                    '半期情報の追加
                    .AddNew
                    .Fields("kikaku_syurui_cd").Value = pstrKikakuSyurui
                    .Fields("nen_ren_no").Value = pstrKikakuNenrenban
                    .Fields("eda_no").Value = pstrKikakuEdaban
                    .Fields("jan_cd").Value = paryCsvData(iRecord)(enmSyoColNo.janCD)
                    If paryCsvData(iRecord)(enmSyoColNo.HyoujyunBaika) <> "" Then
                        .Fields("hyoujyun_baika").Value = CLng(paryCsvData(iRecord)(enmSyoColNo.HyoujyunBaika))
                    Else
                        .Fields("hyoujyun_baika").Value = dbNull.value
                    End If
                    If paryCsvData(iRecord)(enmSyoColNo.SikiriGenka) <> "" Then
                        .Fields("sikiri_genka").Value = CDbl(paryCsvData(iRecord)(enmSyoColNo.SikiriGenka))
                    Else
                        .Fields("sikiri_genka").Value = dbNull.value
                    End If
                    If rstHankiHenkou.RecordCount = 0 Then
                        .Fields("touroku_kbn").Value = 1 'VMST半期変更情報に対象データが存在した場合は１を設定する。
                    Else
                        .Fields("touroku_kbn").Value = 2 'VMST半期変更情報に対象データが存在しない場合は２を設定する。
                    End If
                    .Fields("touroku_user_id").Value = pstrTantousyaId
                    .Fields("touroku_ymd_hms").Value = lstrNow
                    .Update
                    .Filter = adFilterNone
                    strHankiAdded = 1
            End If
        End With
        
        rstHankiHenkou.Filter = adFilterNone
        '売場体系情報
        With rstUriba
            '変数
            Dim allJigyouCd As Object
            Dim allJuyouCd As Object
            '初期設定
            allJigyouCd = ""
            allJuyouCd = ""
            If paryCsvData(iRecord)(enmSyoColNo.VhJyuyouCd) <> "" Then
                    allJigyouCd = allJigyouCd & "'" & JYU_VH & "',"
                    allJuyouCd = allJuyouCd & "'" & paryCsvData(iRecord)(enmSyoColNo.VhJyuyouCd) & "',"
            End If
            If paryCsvData(iRecord)(enmSyoColNo.RdJyuyouCd) <> "" Then
                    allJigyouCd = allJigyouCd & "'" & JYU_RD & "',"
                    allJuyouCd = allJuyouCd & "'" & paryCsvData(iRecord)(enmSyoColNo.RdJyuyouCd) & "',"
            End If
            If paryCsvData(iRecord)(enmSyoColNo.VcsJyuyouCd) <> "" Then
                    allJigyouCd = allJigyouCd & "'" & JYU_VCS & "',"
                    allJuyouCd = allJuyouCd & "'" & paryCsvData(iRecord)(enmSyoColNo.VcsJyuyouCd) & "',"
            End If
            If paryCsvData(iRecord)(enmSyoColNo.SvhJyuyouCd) <> "" Then
                    allJigyouCd = allJigyouCd & "'" & JYU_SVH & "',"
                    allJuyouCd = allJuyouCd & "'" & paryCsvData(iRecord)(enmSyoColNo.SvhJyuyouCd) & "',"
            End If
            
            '売場体系情報の取得
            If allJigyouCd <> "" Then
                allJigyouCd = Left(allJigyouCd, Len(allJigyouCd) - 1)
                
                '売場体系情報の追加
                For i = 0 To UBound(Split(Replace(allJigyouCd, "'", ""), ","))
                    If Not .BOF Then
                        .MoveFirst
                        .Filter = "jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "'" _
                                    & "AND jigyou_cd = '" & Split(Replace(allJigyouCd, "'", ""), ",")(i) & "'"
                        If .RecordCount <> 0 Then
                            .Delete
                        End If
                    End If
            
                    .AddNew
                    .Fields("kikaku_syurui_cd").Value = pstrKikakuSyurui
                    .Fields("nen_ren_no").Value = pstrKikakuNenrenban
                    .Fields("eda_no").Value = pstrKikakuEdaban
                    .Fields("jan_cd").Value = paryCsvData(iRecord)(enmSyoColNo.janCD)
                    If Split(Replace(allJigyouCd, "'", ""), ",")(i) = JYU_VH Then
                        .Fields("jigyou_cd").Value = JYU_VH
                        .Fields("jyuyou_cd").Value = paryCsvData(iRecord)(enmSyoColNo.VhJyuyouCd)
                    ElseIf Split(Replace(allJigyouCd, "'", ""), ",")(i) = JYU_RD Then
                        .Fields("jigyou_cd").Value = JYU_RD
                        .Fields("jyuyou_cd").Value = paryCsvData(iRecord)(enmSyoColNo.RdJyuyouCd)
                    ElseIf Split(Replace(allJigyouCd, "'", ""), ",")(i) = JYU_VCS Then
                        .Fields("jigyou_cd").Value = JYU_VCS
                        .Fields("jyuyou_cd").Value = paryCsvData(iRecord)(enmSyoColNo.VcsJyuyouCd)
                    ElseIf Split(Replace(allJigyouCd, "'", ""), ",")(i) = JYU_SVH Then
                        .Fields("jigyou_cd").Value = JYU_SVH
                        .Fields("jyuyou_cd").Value = paryCsvData(iRecord)(enmSyoColNo.SvhJyuyouCd)
                    End If
                    .Fields("henkou_date").Value = Format(lstrNow, "yyyy/mm/dd")
                    .Fields("touroku_user_id").Value = pstrTantousyaId
                    .Fields("touroku_ymd_hms").Value = lstrNow
                    .Update
                    strUribaAdded = 1
                Next
                .Filter = adFilterNone
            End If
        End With

        'VMST社内基本情報の取得
        With rstSyanaiKihon
            If Not .BOF Then
                .MoveFirst
                .Find ("jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "'")
            End If
        End With
        
        'POP情報
        If paryCsvData(iRecord)(enmSyoColNo.SyouhinCopy) <> "" Or _
            paryCsvData(iRecord)(enmSyoColNo.TuujyouKakaku) <> "" Or _
            paryCsvData(iRecord)(enmSyoColNo.MakerKakaku) <> "" Or _
            paryCsvData(iRecord)(enmSyoColNo.POPLayout1) <> "" Or _
            paryCsvData(iRecord)(enmSyoColNo.POPLayout2) <> "" Or _
            paryCsvData(iRecord)(enmSyoColNo.PopSize) <> "" Then
            With rstPOP
                 'POP情報の取得
                 If Not .BOF Then
                     .MoveFirst
                     .Filter = "jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "' "
                     If .RecordCount <> 0 Then
                         .Delete
                     End If
                 End If
             
                 'POP情報の追加
                 .AddNew
                 .Fields("kikaku_syurui_cd").Value = pstrKikakuSyurui
                 .Fields("nen_ren_no").Value = pstrKikakuNenrenban
                 .Fields("eda_no").Value = pstrKikakuEdaban
                 .Fields("jan_cd").Value = paryCsvData(iRecord)(enmSyoColNo.janCD)
                 '▼商品コピー
                 If paryCsvData(iRecord)(enmSyoColNo.SyouhinCopy) <> "" Then
                     .Fields("syouhin_copy").Value = paryCsvData(iRecord)(enmSyoColNo.SyouhinCopy)
                 Else
                     .Fields("syouhin_copy").Value = rstSyanaiKihon.Fields("syouhin_copy").Value
                 End If
                 '▼通常価格
                 If paryCsvData(iRecord)(enmSyoColNo.TuujyouKakaku) <> "" Then
                     .Fields("tuujyou_kakaku").Value = paryCsvData(iRecord)(enmSyoColNo.TuujyouKakaku)
                 Else
                     .Fields("tuujyou_kakaku").Value = rstSyanaiKihon.Fields("tuujyou_kakaku").Value
                 End If
                 '▼メーカー希望小売価格
                 If paryCsvData(iRecord)(enmSyoColNo.MakerKakaku) <> "" Then
                     .Fields("maker_kiboukouri_kakaku").Value = paryCsvData(iRecord)(enmSyoColNo.MakerKakaku)
                 Else
                     .Fields("maker_kiboukouri_kakaku").Value = rstSyanaiKihon.Fields("maker_kiboukouri_kakaku").Value
                 End If
                 '▼POPレイアウト１
                 If paryCsvData(iRecord)(enmSyoColNo.POPLayout1) <> "" Then
                     .Fields("pop_layout_cd_1").Value = paryCsvData(iRecord)(enmSyoColNo.POPLayout1)
                 Else
                     .Fields("pop_layout_cd_1").Value = rstSyanaiKihon.Fields("pop_layout_cd_1").Value
                 End If
                 '▼POPレイアウト２
                 If paryCsvData(iRecord)(enmSyoColNo.POPLayout2) <> "" Then
                     .Fields("pop_layout_cd_2").Value = paryCsvData(iRecord)(enmSyoColNo.POPLayout2)
                 Else
                     .Fields("pop_layout_cd_2").Value = rstSyanaiKihon.Fields("pop_layout_cd_2").Value
                 End If
                 '▼POPサイズ
                 If paryCsvData(iRecord)(enmSyoColNo.PopSize) <> "" Then
                     .Fields("pop_size_cd").Value = paryCsvData(iRecord)(enmSyoColNo.PopSize)
                 Else
                     .Fields("pop_size_cd").Value = rstSyanaiKihon.Fields("pop_size_cd").Value
                 End If
                 .Fields("touroku_user_id").Value = pstrTantousyaId
                 .Fields("touroku_ymd_hms").Value = lstrNow
                 .Update
                 .Filter = adFilterNone
    
                 strPOPAdded = 1
            End With
        End If

        'WF管理項目を最新情報に更新する
        With rstWfKanri
            If Not .BOF Then
                .MoveFirst
                .Find ("jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "'")
            End If
            'WF管理項目に無い場合は、新規追加
            If .EOF Then
                .AddNew
                .Fields("kikaku_syurui_cd").Value = pstrKikakuSyurui
                .Fields("nen_ren_no").Value = pstrKikakuNenrenban
                .Fields("eda_no").Value = pstrKikakuEdaban
                .Fields("jan_cd").Value = paryCsvData(iRecord)(enmSyoColNo.janCD)
                .Fields("jan_umu_kbn").Value = 1
                .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                .Fields("jyogai_kbn").Value = 0
                .Fields("touroku_tantousya").Value = pstrTantousyaId
                .Fields("keikakusuu").Value = 0
                .Fields("kikaku_no").Value = pstrKikakuSyurui & pstrKikakuNenrenban & "-" & pstrKikakuEdaban
                .Fields("kikaku_kahi_kbn").Value = 0
                .Fields("kikaku_syouninsya_comment").Value = dbNull.value
                .Fields("hinsyu_kahi_kbn").Value = 0
                .Fields("hinsyu_syouninsya_comment").Value = dbNull.value
                .Fields("naibu_cd").Value = rstIppanKihon.Fields("key_cd").Value
                If strHyoujiAdded = 0 Then
                    .Fields("hyouji_siyou_touroku_kbn").Value = 0
                Else
                    .Fields("hyouji_siyou_touroku_kbn").Value = 1
                End If
                .Fields("sekisai_size_touroku_kbn").Value = 0
                'POP情報登録区分
                If strPOPAdded = 0 Then
                    .Fields("pop_touroku_kbn").Value = 0
                Else
                    .Fields("pop_touroku_kbn").Value = 1
                End If
                .Fields("syanai_kihon_touroku_kbn").Value = 0
                .Fields("kowake_touroku_kbn").Value = 0
                .Fields("jigyoubetu_touroku_kbn").Value = 0
                If strHanbaiAdded = 0 Then
                    .Fields("hanbai_touroku_kbn").Value = 0
                Else
                    .Fields("hanbai_touroku_kbn").Value = 1
                End If
                .Fields("siire_touroku_kbn").Value = 0
                .Fields("hinsyu_taikei_touroku_kbn").Value = 0
                If strUribaAdded = 0 Then
                    .Fields("uriba_touroku_kbn").Value = 0
                Else
                    .Fields("uriba_touroku_kbn").Value = 1
                End If
                .Fields("baika_henkou_touroku_kbn").Value = 0
                .Fields("siire_henkou_touroku_kbn").Value = 0
                .Fields("center_cd_henkou_touroku_kbn").Value = 0
                .Fields("kowake_henkou_touroku_kbn").Value = 0
                If strHankiAdded = 0 Then
                    .Fields("hanki_henkou_touroku_kbn").Value = 0
                Else
                    .Fields("hanki_henkou_touroku_kbn").Value = 1
                End If
                .Fields("touroku_user_id").Value = pstrTantousyaId
                .Fields("kousin_user_id").Value = dbNull.value
                .Fields("touroku_ymd_hms").Value = lstrNow
                .Fields("kousin_ymd_hms").Value = dbNull.value
    
            'WF管理項目にある場合は、更新
            Else
                .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                .Fields("touroku_tantousya").Value = pstrTantousyaId
                .Fields("kikaku_no").Value = pstrKikakuSyurui & pstrKikakuNenrenban & "-" & pstrKikakuEdaban
                .Fields("naibu_cd").Value = rstIppanKihon.Fields("key_cd").Value
                If strHyoujiAdded = 0 Then
                Else
                    .Fields("hyouji_siyou_touroku_kbn").Value = 1
                End If
                'POP情報登録区分
                If strPOPAdded = 0 Then
                Else
                    .Fields("pop_touroku_kbn").Value = 1
                End If
                If strHanbaiAdded = 0 Then
                Else
                    .Fields("hanbai_touroku_kbn").Value = 1
                End If
                If strUribaAdded = 0 Then
                Else
                    .Fields("uriba_touroku_kbn").Value = 1
                End If
                If strHankiAdded = 0 Then
                Else
                    .Fields("hanki_henkou_touroku_kbn").Value = 1
                End If
                .Fields("kousin_user_id").Value = pstrTantousyaId
                .Fields("kousin_ymd_hms").Value = lstrNow
            End If
            .Update
        End With
    Next iRecord

    Call objCDBAccess.UpdateDCRecordset("VSWF", rstHyoujiSiyou)
    Call objCDBAccess.UpdateDCRecordset("VSWF", rstHanbai)
    Call objCDBAccess.UpdateDCRecordset("VSWF", rstHanki)
    Call objCDBAccess.UpdateDCRecordset("VSWF", rstUriba)
    Call objCDBAccess.UpdateDCRecordset("VSWF", rstPOP)
    Call objCDBAccess.UpdateDCRecordset("VSWF", rstWfKanri)
    
    'RSのバッチ更新
    'GetObjectContext.SetComplete
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(オブジェクト開放)"
    If Not rstHyoujiSiyou Is Nothing Then
            rstHyoujiSiyou = Nothing
    End If

    If Not rstWfKanri Is Nothing Then
            rstWfKanri = Nothing
    End If

    If Not rstHanbai Is Nothing Then
            rstHanbai = Nothing
    End If

    If Not rstHanki Is Nothing Then
            rstHanki = Nothing
    End If

    If Not rstUriba Is Nothing Then
            rstUriba = Nothing
    End If

    If Not rstPOP Is Nothing Then
            rstPOP = Nothing
    End If

    If Not rstIppanKihon Is Nothing Then
            rstIppanKihon = Nothing
    End If
    
    If Not rstSyanaiKihon Is Nothing Then
            rstSyanaiKihon = Nothing
    End If
    
    If Not rstHankiHenkou Is Nothing Then
            rstHankiHenkou = Nothing
    End If
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(登録時利用したのでここで開放する)"
    
    Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If

    If Not rstHyoujiSiyou Is Nothing Then
            rstHyoujiSiyou = Nothing
    End If

    If Not rstWfKanri Is Nothing Then
            rstWfKanri = Nothing
    End If

    If Not rstHanbai Is Nothing Then
            rstHanbai = Nothing
    End If

    If Not rstHanki Is Nothing Then
            rstHanki = Nothing
    End If

    If Not rstUriba Is Nothing Then
            rstUriba = Nothing
    End If

    If Not rstPOP Is Nothing Then
            rstPOP = Nothing
    End If

    If Not rstIppanKihon Is Nothing Then
            rstIppanKihon = Nothing
    End If
    
    If Not rstSyanaiKihon Is Nothing Then
            rstSyanaiKihon = Nothing
    End If

    If Not rstHankiHenkou Is Nothing Then
            rstHankiHenkou = Nothing
    End If

    'If Not GetObjectContext Is Nothing Then
        'GetObjectContext.SetAbort
    End If

    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
            'Err.Number, Err.Source, Err.Description
End Sub



'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckDB_JyuyouCode
'*【処理概要】
'*?    第1分類コード、第2分類コード、第3分類コードが有効かをチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intNowLine; Integer; CSVファイルの現在の行
'*P    intFieldType; Integer; RET_FLD_BUN:bunrui情報
'*【戻り値】
'*    < Name Type Comment >
'*R    strKeyCode ; String JANコードに対応する統合マスタのキーコードを格納する
'*R    FncCheckDB_JyuyouCode; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2005/09/02; 高附　周造(IC); 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheckDB_JyuyouCode(ByRef arryCSVRecord() As String, intNowLine As Integer, intFieldType As Integer) As Boolean
    
    'a.分類情報が有効か確認する。
    FncCheckDB_JyuyouCode = True
    
    '　(a)売場体系マスタに同一の需要コードが存在する事を確認する。
    With pChkJyuyouCodeRS
        If Not .BOF Then
                .MoveFirst
            If arryCSVRecord(enmSyoColNo.VhJyuyouCd) <> "" Then
                .Filter = "jigyou_cd = '" & JYU_VH & "' " _
                        & "AND jyuyou_cd = '" & arryCSVRecord(enmSyoColNo.VhJyuyouCd) & "'"
                If .EOF = True Then
                    FncSetError (intNowLine, enmSyoColNo.VhJyuyouCd, PARAM_NOTFINDMASTER
                    FncCheckDB_JyuyouCode = False
                End If
                .Filter = adFilterNone
            End If
            If arryCSVRecord(enmSyoColNo.RdJyuyouCd) <> "" Then
                .Filter = "jigyou_cd = '" & JYU_RD & "' " _
                        & "AND jyuyou_cd = '" & arryCSVRecord(enmSyoColNo.RdJyuyouCd) & "'"
                If .EOF = True Then
                    FncSetError (intNowLine, enmSyoColNo.RdJyuyouCd, PARAM_NOTFINDMASTER
                    FncCheckDB_JyuyouCode = False
                End If
                .Filter = adFilterNone
            End If
            If arryCSVRecord(enmSyoColNo.VcsJyuyouCd) <> "" Then
                .Filter = "jigyou_cd = '" & JYU_VCS & "' " _
                        & "AND jyuyou_cd = '" & arryCSVRecord(enmSyoColNo.VcsJyuyouCd) & "'"
                If .EOF = True Then
                    FncSetError (intNowLine, enmSyoColNo.VcsJyuyouCd, PARAM_NOTFINDMASTER
                    FncCheckDB_JyuyouCode = False
                End If
                .Filter = adFilterNone
            End If
            If arryCSVRecord(enmSyoColNo.SvhJyuyouCd) <> "" Then
                .Filter = "jigyou_cd = '" & JYU_SVH & "' " _
                        & "AND jyuyou_cd = '" & arryCSVRecord(enmSyoColNo.SvhJyuyouCd) & "'"
                If .EOF = True Then
                    FncSetError (intNowLine, enmSyoColNo.SvhJyuyouCd, PARAM_NOTFINDMASTER
                    FncCheckDB_JyuyouCode = False
                End If
                .Filter = adFilterNone
            End If
        End If
    End With
            
End Function

'*< UNISYSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckDB_AtukaiPtnCd
'*【処理概要】
'*?    扱いパターンコードが有効かをチェックします。
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intNowLine; Integer; CSVファイルの現在の行
'*P    intFieldType; Integer; RET_FLD_SYO:商品情報
'*【戻り値】
'*    < Name Type Comment >
'*R    FncCheckDB_AtukaiPtnCd; Boolean;チェックOK:True,NG:False
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2006/09/22; 和田(IC); 商品選定ＷＦ問合せ撲滅改善;
'*< Comment End >**************************************************'
Private Function FncCheckDB_AtukaiPtnCd(ByVal strAtukaiPtnCd As String, _
                                      ByVal intNowLine As Integer, _
                                      ByVal intFieldType As Integer) As Boolean
    
    'a.扱いパターン情報が有効か確認する。
    FncCheckDB_AtukaiPtnCd = True
    
    '　(a)扱いパターンマスタに存在する事を確認する。
    With pChkAtukaiPtnCdRS
        If Not .EOF Then
                .MoveFirst
            If strAtukaiPtnCd <> "" Then
                .Filter = "atukai_ptn_cd = '" & strAtukaiPtnCd & "' "
                If .EOF = True Then
                    FncSetError (intNowLine, enmSyoColNo.AtukaiPtn, PARAM_NOTFINDMASTER
                    FncCheckDB_AtukaiPtnCd = False
                End If
                .Filter = adFilterNone
            End If
        End If
    End With
            
End Function

'*< TOSTEM >*******************************************************************
'*
'*【プロシージャ名】
'*!     GetSaibanNo
'*
'*【処理概要】
'*?     採番NOの取得。
'*
'*【パラメータ】
'*      Name                　　  Type                Comment
'*P     strSaibanKbn;             string              採番区分
'*
'*【戻り値】
'*      Name                　　  Type                Comment
'*R     GetSaibanNo;              integer;            採番NO;
'*
'******************************************************************************
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)       EDP-No      説明
'*M 2008/02/29; 和田(IC);         P-32685;    商品選定ＷＦシステムの原価・売価変更登録時の値入率チェック機能追加;
'*< Comment End >**************************************************************
Public Function GetSaibanNo(ByVal strSaibanKbn As String) As Integer
                               
    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    Dim objCDBAccess    As COMDBUtil.CDBAccess
    Dim strSQL          As String
    Dim rsSaiban        As ADODB.Recordset
    
    On Error GoTo DefectErrorHandler

    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.GetSaibanNo(" & strSaibanKbn & _
                                                            ")"
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL文
    strSQL = ""
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     saiban_no + 1 AS saiban_no"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     w_syousen_neire_keikoku_saiban WITH(UPDLOCK)"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     saiban_kbn = ?"

    '切断Recordset取得
        rsSaiban = objCDBAccess.GetDCRecordset("VSWF", strSQL, MakeParam("@saiban_kbn", DT_Char, 1, strSaibanKbn))
    
    GetSaibanNo = rsSaiban.Fields("saiban_no").Value
    If GetSaibanNo >= 1000 Then
        GetSaibanNo = 0
    End If
    rsSaiban.Close
        rsSaiban = Nothing
    
     'SQL文
    strSQL = ""
    strSQL = strSQL & " UPDATE"
    strSQL = strSQL & "     w_syousen_neire_keikoku_saiban"
    strSQL = strSQL & " SET saiban_no = ?"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     saiban_kbn = ?"
   
    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no", DT_Int, 2, GetSaibanNo), _
                        MakeParam("@saiban_kbn", DT_Char, 1, strSaibanKbn) _
                        )
    
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    'インスタンス破棄
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Function

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
             'Err.Number, Err.Source, Err.Description
End Function

'*< TOSTEM >****************************************************'
'*【プロシージャ名】
'*!    NeirerituCheck
'*【処理概要】
'*?    値入率のチェック
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    aryData;             CSVファイル情報
'*P    strKikakuSyuruiCd;   企画種類コード
'*P    strNenRenNo;         年連番
'*P    strEdaNo;            枝番
'*P    strLoginUser;        ログインユーザー
'*P    intSaibanNo;         採番NO
'*P    sngSyouhizeiRitu;    消費税率(%)
'*【戻り値】
'*    < Name Type Comment >
'*R    なし
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2008/03/03; 和田(IC); P-32685 商品選定WFシステムの原価・売価変更登録時の値入率チェック機能追加;
'*< Comment End >**************************************************'
Public Sub NeirerituCheck(ByVal aryData As Object _
                      , ByVal strKikakuSyuruiCD As String _
                      , ByVal strNenRenNO As String _
                      , ByVal strEdaNO As String _
                      , ByVal strLoginUser As String _
                      , ByVal intSaibanNo As Integer _
                      , ByVal sngSyouhizeiRitu As Single)
    
    Dim strMethodCallInfo As String
    Dim strSaibanKbn As String
    
    On Error GoTo DefectErrorHandler
   
    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.NeirerituCheck(aryData" & _
                                                        "," & strKikakuSyuruiCD & _
                                                        "," & strNenRenNO & _
                                                        "," & strEdaNO & _
                                                        "," & strLoginUser & _
                                                        "," & intSaibanNo & _
                                                        "," & sngSyouhizeiRitu & _
                                                        ")"
    '■ワークテーブルの初期化
    Call DeleteWork(intSaibanNo)
    
    '■JANコードのダブり排除、登録
    Call DistinctJanCd(aryData, strKikakuSyuruiCD, strNenRenNO, strEdaNO, intSaibanNo)
   
    '■値入率警告情報の初期化
    Call DeleteNeireKeikoku(intSaibanNo)
   
    '■値入率警告情報の登録
    Call InsertNeireKeikoku(strLoginUser, intSaibanNo, sngSyouhizeiRitu)
    
    '■ワークテーブルの初期化
    'Call DeleteWork(intSaibanNo)
Exit Sub

DefectErrorHandler:
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >****************************************************'
'*【プロシージャ名】
'*!    DeleteWork
'*【処理概要】
'*?    ワークテーブルの初期化
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intSaibanNo;         採番NO
'*【戻り値】
'*    < Name Type Comment >
'*R    なし
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2008/02/27; 和田(IC); P-32685 商品選定WFシステムの原価・売価変更登録時の値入率チェック機能追加;
'*< Comment End >**************************************************'
Private Sub DeleteWork(ByVal intSaibanNo As Integer)
    
    Dim strMethodCallInfo As String
    Dim strSQL As String
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    
    On Error GoTo DefectErrorHandler
   
    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.DeleteWork(" & intSaibanNo & _
                                                            ")"
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
       
    '▼商品選定WF(値入率警告)WORK
    'SQL文
    strSQL = ""
    strSQL = strSQL & " DELETE FROM"
    strSQL = strSQL & "     w_syousen_neire_keikoku"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     saiban_no = ?"

    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, MakeParam("@saiban_no", DT_Int, 2, intSaibanNo))
  
    '▼商品選定WF(値入率警告キー)WORK
    'SQL文
    strSQL = ""
    strSQL = strSQL & " DELETE FROM"
    strSQL = strSQL & "     w_syousen_neire_keikoku_key"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     saiban_no = ?"

    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, MakeParam("@saiban_no", DT_Int, 2, intSaibanNo))
  
    '▼商品選定WF(値入率警告JAN)WORK
    'SQL文
    strSQL = ""
    strSQL = strSQL & " DELETE FROM"
    strSQL = strSQL & "     w_syousen_neire_keikoku_jan"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     saiban_no = ?"

    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, MakeParam("@saiban_no", DT_Int, 2, intSaibanNo))
  
    'インスタンス破棄
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >****************************************************'
'*【プロシージャ名】
'*!    DistinctJanCd
'*【処理概要】
'*?    CSV情報のJANコードをDISTINCT、登録
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    aryData;                  CSVファイル情報
'*     strKikakuSyuruiCd;        企画種類コード
'*     strNenRenNo;              年連番
'*     strEdaNo;                 枝番
'*P    intSaibanNo;              採番NO
'*【戻り値】
'*    < Name Type Comment >
'*R    なし
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2008/02/27; 和田(IC); P-32685 商品選定WFシステムの原価・売価変更登録時の値入率チェック機能追加;
'*< Comment End >**************************************************'
Private Sub DistinctJanCd(ByVal aryData As Object _
                  , ByVal strKikakuSyuruiCD As String _
                  , ByVal strNenRenNO As String _
                  , ByVal strEdaNO As String _
                  , ByVal intSaibanNo As String)
    
    Dim strMethodCallInfo As String
    Dim i As Integer
    Dim rsJan As ADODB.Recordset
    
    On Error GoTo DefectErrorHandler
   
    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.DistinctJanCd(aryData" & _
                                                            "," & strKikakuSyuruiCD & _
                                                            "," & strNenRenNO & _
                                                            "," & strEdaNO & _
                                                            "," & intSaibanNo & _
                                                            ")"
    For i = 0 To UBound(aryData)
        'JANコード存在チェック
            rsJan = GetJanInfo(strKikakuSyuruiCD _
                                , strNenRenNO _
                                , strEdaNO _
                                , aryData(i)(enmTenColNo.janCD) _
                                , aryData(i)(enmTenColNo.HenkouDate) _
                                , intSaibanNo)
        If rsJan.EOF Then
            '存在しない場合、登録
            Call InsertJan(strKikakuSyuruiCD _
                                , strNenRenNO _
                                , strEdaNO _
                                , aryData(i)(enmTenColNo.janCD) _
                                , aryData(i)(enmTenColNo.HenkouDate) _
                                , intSaibanNo)
        End If
        rsJan.Close
    Next

Exit Sub

DefectErrorHandler:
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >*******************************************************************
'*
'*【プロシージャ名】
'*!     GetJanInfo
'*
'*【処理概要】
'*?     商品選定WF(値入率警告JAN)WORKレコードセットの取得。
'*
'*【パラメータ】
'*      Name                　　  Type                Comment
'*      strKikakuSyuruiCd;        String;             企画種類コード;
'*      strNenRenNo;              String;             年連番;
'*      strEdaNo;                 String;             枝番;
'*      strJanCd;                 String;             JANコード;
'*      strHenkouDate;            String;             変更日;
'*      intSaibanNo;              integer;            採番NO;
'*
'*【戻り値】
'*      Name                　　  Type                Comment
'*R     GetBaikaHenkou;           ADODB.Recordset;    商品選定WF(売価変更情報)レコードセット;
'*
'******************************************************************************
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)       EDP-No      説明
'*M 2008/02/29; 和田(IC);         P-32685;    商品選定ＷＦシステムの原価・売価変更登録時の値入率チェック機能追加;
'*< Comment End >**************************************************************
Private Function GetJanInfo(ByVal strKikakuSyuruiCD As String, _
                              ByVal strNenRenNO As String, _
                              ByVal strEdaNO As String, _
                              ByVal strJanCd As String, _
                              ByVal strHenkouDate As String, _
                              ByVal intSaibanNo As Integer) As ADODB.Recordset
                               
    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    Dim objCDBAccess    As COMDBUtil.CDBAccess
    Dim strSQL          As String
    
    On Error GoTo DefectErrorHandler

    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.GetJanInfo(" & strKikakuSyuruiCD & _
                                                            "," & strNenRenNO & _
                                                            "," & strEdaNO & _
                                                            "," & strJanCd & _
                                                            "," & strHenkouDate & _
                                                            "," & intSaibanNo & _
                                                            ")"
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL文
    strSQL = ""
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,henkou_date"
    strSQL = strSQL & "     ,saiban_no"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     w_syousen_neire_keikoku_jan WITH(READCOMMITTED)"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     kikaku_syurui_cd = ?"
    strSQL = strSQL & "     AND nen_ren_no = ?"
    strSQL = strSQL & "     AND eda_no = ?"
    strSQL = strSQL & "     AND jan_cd = ?"
    strSQL = strSQL & "     AND henkou_date = ?"
    strSQL = strSQL & "     AND saiban_no = ?"

    'MakeParamの作成
    vntAryParam = Array(MakeParam("@kikaku_syurui_cd", DT_Char, 3, strKikakuSyuruiCD), _
                        MakeParam("@nen_ren_no", DT_Char, 10, strNenRenNO), _
                        MakeParam("@eda_no", DT_Char, 3, strEdaNO), _
                        MakeParam("@jan_cd", DT_VarChar, 20, strJanCd), _
                        MakeParam("@henkou_date", DT_Char, 8, strHenkouDate), _
                        MakeParam("@saiban_no", DT_Int, 2, intSaibanNo) _
                        )
    
    '切断Recordset取得
        GetJanInfo = objCDBAccess.GetDCRecordset("VSWF", strSQL, vntAryParam)
    
    'インスタンス破棄
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Function

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
             'Err.Number, Err.Source, Err.Description
End Function

'*< TOSTEM >****************************************************'
'*【プロシージャ名】
'*!    InsertJan
'*【処理概要】
'*?    商品選定WF(値入率警告JAN)WORKの登録
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strKikakuSyuruiCd;   企画種類コード
'*P    strNenRenNo;         年連番
'*P    strEdaNo;            枝番
'*     strJanCd;            JANコード
'*     strHenkouDate;       変更日
'*     intSaibanNo;         採番NO
'*【戻り値】
'*    < Name Type Comment >
'*R    なし
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2008/02/27; 和田(IC); P-32685 商品選定WFシステムの原価・売価変更登録時の値入率チェック機能追加;
'*< Comment End >**************************************************'
Private Sub InsertJan(ByVal strKikakuSyuruiCD As String _
                          , ByVal strNenRenNO As String _
                          , ByVal strEdaNO As String _
                          , ByVal strJanCd As String _
                          , ByVal strHenkouDate As String _
                          , ByVal intSaibanNo As Integer)
    
    Dim strMethodCallInfo As String
    Dim strSQL As String
    Dim vntAryParam As Object
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim strSyoriKbn As String
    
    On Error GoTo DefectErrorHandler
   
    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.InsertJan(" & strKikakuSyuruiCD & _
                                                            "," & strNenRenNO & _
                                                            "," & strEdaNO & _
                                                            "," & strJanCd & _
                                                            "," & strHenkouDate & _
                                                            "," & intSaibanNo & _
                                                            ")"
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
       
    'SQL文
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku_jan("
    strSQL = strSQL & "     kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,henkou_date"
    strSQL = strSQL & "     ,saiban_no)"
    strSQL = strSQL & " VALUES("
    strSQL = strSQL & "     ?"
    strSQL = strSQL & "     ,?"
    strSQL = strSQL & "     ,?"
    strSQL = strSQL & "     ,?"
    strSQL = strSQL & "     ,?"
    strSQL = strSQL & "     ,?)"

    'MakeParamの作成
    vntAryParam = Array(MakeParam("@kikaku_syurui_cd", DT_Char, 3, strKikakuSyuruiCD), _
                        MakeParam("@nen_ren_no", DT_Char, 10, strNenRenNO), _
                        MakeParam("@eda_no", DT_Char, 3, strEdaNO), _
                        MakeParam("@jan_cd", DT_VarChar, 20, strJanCd), _
                        MakeParam("@henkou_date", DT_Char, 8, strHenkouDate), _
                        MakeParam("@saiban_no", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
  
    'インスタンス破棄
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >****************************************************'
'*【プロシージャ名】
'*!    DeleteNeireKeikoku
'*【処理概要】
'*?    値入率警告情報の初期化
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intSaibanNo;         採番NO
'*【戻り値】
'*    < Name Type Comment >
'*R    なし
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2008/02/27; 和田(IC); P-32685 商品選定WFシステムの原価・売価変更登録時の値入率チェック機能追加;
'*< Comment End >**************************************************'
Private Sub DeleteNeireKeikoku(ByVal intSaibanNo As Integer)
    
    Dim strMethodCallInfo As String
    Dim strSQL As String
    Dim vntAryParam As Object
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    
    On Error GoTo DefectErrorHandler
   
    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.DeleteNeireKeikoku(" & intSaibanNo & _
                                                                    ")"
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
       
    'SQL文
    strSQL = ""
    strSQL = strSQL & " DELETE"
    strSQL = strSQL & "     keikoku"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     t_syousen_neire_keikoku keikoku"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         w_syousen_neire_keikoku_jan jan"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         keikoku.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "         AND keikoku.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "         AND keikoku.eda_no = jan.eda_no"
    strSQL = strSQL & "         AND keikoku.jan_cd = jan.jan_cd"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     jan.saiban_no = ?"

    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, MakeParam("@saiban_no", DT_Int, 2, intSaibanNo))
  
    'インスタンス破棄
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >****************************************************'
'*【プロシージャ名】
'*!    InsertNeireKeikoku
'*【処理概要】
'*?    値入率警告の登録
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strLoginUser;        ログインユーザー
'*P    intSaibanNo;         採番NO
'*P    sngSyouhizeiRitu;    消費税率(%)
'*【戻り値】
'*    < Name Type Comment >
'*R    なし
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2008/02/27; 和田(IC); P-32685 商品選定WFシステムの原価・売価変更登録時の値入率チェック機能追加;
'*< Comment End >**************************************************'
Private Sub InsertNeireKeikoku(ByVal strLoginUser As String _
                          , ByVal intSaibanNo As Integer _
                          , ByVal sngSyouhizeiRitu As Single)
    
    Dim strMethodCallInfo As String
    
    On Error GoTo DefectErrorHandler
   
    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.InsertNeireKeikoku(" & strLoginUser & _
                                                            "," & intSaibanNo & _
                                                            "," & sngSyouhizeiRitu & _
                                                            ")"
    '▼売価変更情報のチェック
        '直近過去処理
        '商品選定WF(値入率警告)WORK
    Call InsertBaikaWork(strLoginUser _
                    , intSaibanNo _
                    , sngSyouhizeiRitu _
                    )

        '商品選定WF(値入率警告KEY)WORK
    Call InsertBaikaKeyWork(intSaibanNo)
        
        '商品選定WF(値入率警告)
    Call InsertNeire(intSaibanNo, "1")
    
    '▼仕入変更情報のチェック
        '直近過去処理
        '商品選定WF(値入率警告)WORK
    Call InsertGenkaWork(strLoginUser _
                    , intSaibanNo _
                    , sngSyouhizeiRitu _
                    )

        '商品選定WF(値入率警告KEY)WORK
    Call InsertGenkaKeyWork(intSaibanNo)
        
        '商品選定WF(値入率警告)
    Call InsertNeire(intSaibanNo, "2")
    
    
    'Commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >*******************************************************************
'*
'*【プロシージャ名】
'*!     InsertBaikaWork
'*
'*【処理概要】
'*?     商品選定WF(値入率警告)WORKに値入率異常分を登録
'*
'*【パラメータ】
'*      Name                　　  Type                Comment
'*      strLoginUser;             String;             ログインユーザー;
'*      intSaibanNo;              Integer;            採番NO;
'*      sngSyouhizeiRitu;         Single;             消費税率(%);
'*
'*【戻り値】
'*      Name                　　  Type                Comment
'*R     なし
'*
'******************************************************************************
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)       EDP-No      説明
'*M 2008/02/29; 和田(IC);         P-32685;    商品選定ＷＦシステムの原価・売価変更登録時の値入率チェック機能追加;
'*M 2009/04/22; Sureerat P.;      P-35471;    Mod check 値入率上限,値入率下限;
'*M 2014/01/22; 李松涛(大連) ;　P-46854; 税制改正に伴う対応(商品選定WF・新商品情報);
'*M 2014/03/04; 和田(IC);         P-46854;    処理を分割(タイムアウト対応);
'*M 2014/04/10; 和田(IC);         P-46854;    KEY登録簡易化;
'*M 2014/04/18; 和田(IC);         P-46854;    仕入変更情報(タイムアウト対応);
'*< Comment End >**************************************************************
Private Sub InsertBaikaWork(ByVal strLoginUser As String _
                      , ByVal intSaibanNo As Integer _
                      , ByVal sngSyouhizeiRitu As Single)

    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim strSQL As String
    
    On Error GoTo DefectErrorHandler

    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.InsertBaikaWork(" & strLoginUser & _
                                                            "," & intSaibanNo & _
                                                            "," & sngSyouhizeiRitu & _
                                                             ")"
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    Dim syouhizeiRitu '消費税率
    Dim zeiHanteiKbn  '税判定区分
    Dim objVswf
    Dim objStr
    
    '商品情報登録COMのインスタンス生成
        objVswf = CreateObject("VSWF01BUS005.CVswfSyouhizei")
    
    '消費税率と税判定区分データを取得する
    objStr = objVswf.GetKaisibiJyouhou()
    syouhizeiRitu = Split(objStr, "|")(0)
    zeiHanteiKbn = Split(objStr, "|")(1)
    
        objVswf = Nothing
    'SQL文
    '▼商品選定WF(仕入変更情報)とのチェック
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,baika"
    strSQL = strSQL & "     ,baika_kikaku_no"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,genka"
    strSQL = strSQL & "     ,genka_kikaku_no"
    strSQL = strSQL & "     ,neire_ritu"
    strSQL = strSQL & "     ,touroku_naiyou"
    strSQL = strSQL & "     ,kousin_user_id"
    strSQL = strSQL & "     ,kousin_ymd_hms"
    strSQL = strSQL & "     ,saiban_no"
    strSQL = strSQL & "     ,table_kbn)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     sel2.kikaku_syurui_cd"
    strSQL = strSQL & "     ,sel2.nen_ren_no"
    strSQL = strSQL & "     ,sel2.eda_no"
    strSQL = strSQL & "     ,sel2.jan_cd"
    strSQL = strSQL & "     ,sel2.ten_cd"
    strSQL = strSQL & "     ,sel2.baika_henkou_date"
    strSQL = strSQL & "     ,sel2.baika"
    strSQL = strSQL & "     ,sel2.baika_kikaku_no"
    strSQL = strSQL & "     ,sel2.genka_henkou_date"
    strSQL = strSQL & "     ,sel2.genka"
    strSQL = strSQL & "     ,sel2.genka_kikaku_no"
    strSQL = strSQL & "     ,CASE"
    strSQL = strSQL & "         WHEN ROUND(sel2.neire_ritu,1) > 9999.9 THEN 9999.9"
    strSQL = strSQL & "         WHEN ROUND(sel2.neire_ritu,1) < -9999.9 THEN -9999.9"
    strSQL = strSQL & "         ELSE ROUND(sel2.neire_ritu,1)"
    strSQL = strSQL & "         END"
    strSQL = strSQL & "     ,sel2.touroku_naiyou"
    strSQL = strSQL & "     ,sel2.kousin_user_id"
    strSQL = strSQL & "     ,sel2.kousin_ymd_hms"
    strSQL = strSQL & "     ,sel2.saiban_no"
    strSQL = strSQL & "     ,sel2.table_kbn"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     (SELECT"
    strSQL = strSQL & "         sel.kikaku_syurui_cd"
    strSQL = strSQL & "         ,sel.nen_ren_no"
    strSQL = strSQL & "         ,sel.eda_no"
    strSQL = strSQL & "         ,sel.jan_cd"
    strSQL = strSQL & "         ,sel.ten_cd"
    strSQL = strSQL & "         ,sel.baika_henkou_date"
    strSQL = strSQL & "         ,sel.baika"
    strSQL = strSQL & "         ,NULL AS baika_kikaku_no"
    strSQL = strSQL & "         ,sel.genka_henkou_date"
    strSQL = strSQL & "         ,sel.genka"
    strSQL = strSQL & "         ,sel.kikaku_no AS genka_kikaku_no"
   
    If zeiHanteiKbn = "0" Then  ' 税込み の場合
        '|--------------------(売価(税抜き)------------------------||--原価---||--------------------(売価(税抜き)--------------|
        '((売価(税込み) × 消費税率 ÷ (消費税率 ＋ 100)－ 消費税額) - 原価)  /  (売価(税込み) × 消費税率 ÷ (消費税率 ＋ 100)
        strSQL = strSQL & "         ,((sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) - sel.genka) / (sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) * 100 AS neire_ritu"     '値入率
    Else                        ' 税抜き の場合
        '(売価(税抜き) － 原価) ÷ 売価(税抜き)
        strSQL = strSQL & "         ,round((sel.baika- sel.genka) / sel.baika,3) * 100 AS neire_ritu "      '値入率
    End If
    
    strSQL = strSQL & "         ,'1' AS touroku_naiyou"         '売価変更(固定値)
    strSQL = strSQL & "         ,'" & strLoginUser & "' AS kousin_user_id"
    strSQL = strSQL & "         ,GETDATE() AS kousin_ymd_hms"
    strSQL = strSQL & "         ,? AS saiban_no"
    strSQL = strSQL & "         ,sel.table_kbn"
    strSQL = strSQL & "         ,ISNULL(a.neireritu_jyougen,60) AS  neireritu_jyougen"
    strSQL = strSQL & "         ,ISNULL(a.neireritu_kagen,5) AS neireritu_kagen"
    strSQL = strSQL & "     FROM"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             w_baika.kikaku_syurui_cd"
    strSQL = strSQL & "             ,w_baika.nen_ren_no"
    strSQL = strSQL & "             ,w_baika.eda_no"
    strSQL = strSQL & "             ,w_baika.jan_cd"
    strSQL = strSQL & "             ,w_baika.ten_cd"
    strSQL = strSQL & "             ,w_baika.henkou_date AS baika_henkou_date"
    strSQL = strSQL & "             ,w_baika.baika"
    strSQL = strSQL & "             ,w_siire.henkou_date AS genka_henkou_date"
    strSQL = strSQL & "             ,w_siire.genka"
    strSQL = strSQL & "             ,w_siire.kikaku_syurui_cd + w_siire.nen_ren_no + '-' + w_siire.eda_no AS kikaku_no"
    strSQL = strSQL & "             ,'2' AS table_kbn"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             t_syousen_baika_henkou w_baika WITH(READCOMMITTED)"
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 w_syousen_neire_keikoku_jan jan WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 w_baika.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "                 AND w_baika.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "                 AND w_baika.eda_no = jan.eda_no"
    strSQL = strSQL & "                 AND w_baika.jan_cd = jan.jan_cd"
    strSQL = strSQL & "                 AND w_baika.henkou_date = jan.henkou_date"
                                    '商品選定WF(仕入変更情報)
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 t_syousen_siire_henkou w_siire WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 w_baika.kikaku_syurui_cd = w_siire.kikaku_syurui_cd"
    strSQL = strSQL & "                 AND w_baika.nen_ren_no = w_siire.nen_ren_no"
    strSQL = strSQL & "                 AND w_baika.eda_no = w_siire.eda_no"
    strSQL = strSQL & "                 AND w_baika.jan_cd = w_siire.jan_cd"
    strSQL = strSQL & "                 AND w_baika.ten_cd = w_siire.ten_cd"
    strSQL = strSQL & "         WHERE"
    strSQL = strSQL & "             jan.saiban_no = ?"
    strSQL = strSQL & "             AND w_baika.henkou_date >= w_siire.henkou_date"
    strSQL = strSQL & "             AND w_baika.touroku_kbn <> '9'"
    strSQL = strSQL & "             AND w_siire.touroku_kbn <> '9'"
    strSQL = strSQL & "             AND w_baika.baika <> 0"
    strSQL = strSQL & "             AND ISNULL(w_siire.genka,0) <> 0"
    strSQL = strSQL & "         ) sel"
    strSQL = strSQL & "         LEFT JOIN"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             v_key_1.jan_cd"
    strSQL = strSQL & "             ,m_neireritu.neireritu_jyougen"
    strSQL = strSQL & "             ,m_neireritu.neireritu_kagen"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             m_neireritu  WITH(READCOMMITTED)"
    strSQL = strSQL & "             LEFT JOIN"
    strSQL = strSQL & "                 v_syanai_kihon  WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_syanai_kihon.bumon_cd = m_neireritu.daibunrui_cd"
    strSQL = strSQL & "             LEFT JOIN"
    strSQL = strSQL & "                 v_key_1  WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_key_1.key_cd = v_syanai_kihon.key_cd"
    strSQL = strSQL & "         ) a"
    strSQL = strSQL & "         ON"
    strSQL = strSQL & "             a.jan_cd = sel.jan_cd"
    strSQL = strSQL & "      ) sel2"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     sel2.neire_ritu < sel2.neireritu_kagen"
    strSQL = strSQL & "     OR"
    strSQL = strSQL & "     sel2.neire_ritu > sel2.neireritu_jyougen"

    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@saiban_no2", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)


    'SQL文
    '▼仕入変更情報とのチェック
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,baika"
    strSQL = strSQL & "     ,baika_kikaku_no"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,genka"
    strSQL = strSQL & "     ,genka_kikaku_no"
    strSQL = strSQL & "     ,neire_ritu"
    strSQL = strSQL & "     ,touroku_naiyou"
    strSQL = strSQL & "     ,kousin_user_id"
    strSQL = strSQL & "     ,kousin_ymd_hms"
    strSQL = strSQL & "     ,saiban_no"
    strSQL = strSQL & "     ,table_kbn)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     sel2.kikaku_syurui_cd"
    strSQL = strSQL & "     ,sel2.nen_ren_no"
    strSQL = strSQL & "     ,sel2.eda_no"
    strSQL = strSQL & "     ,sel2.jan_cd"
    strSQL = strSQL & "     ,sel2.ten_cd"
    strSQL = strSQL & "     ,sel2.baika_henkou_date"
    strSQL = strSQL & "     ,sel2.baika"
    strSQL = strSQL & "     ,sel2.baika_kikaku_no"
    strSQL = strSQL & "     ,sel2.genka_henkou_date"
    strSQL = strSQL & "     ,sel2.genka"
    strSQL = strSQL & "     ,sel2.genka_kikaku_no"
    strSQL = strSQL & "     ,CASE"
    strSQL = strSQL & "         WHEN ROUND(sel2.neire_ritu,1) > 9999.9 THEN 9999.9"
    strSQL = strSQL & "         WHEN ROUND(sel2.neire_ritu,1) < -9999.9 THEN -9999.9"
    strSQL = strSQL & "         ELSE ROUND(sel2.neire_ritu,1)"
    strSQL = strSQL & "         END"
    strSQL = strSQL & "     ,sel2.touroku_naiyou"
    strSQL = strSQL & "     ,sel2.kousin_user_id"
    strSQL = strSQL & "     ,sel2.kousin_ymd_hms"
    strSQL = strSQL & "     ,sel2.saiban_no"
    strSQL = strSQL & "     ,sel2.table_kbn"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     (SELECT"
    strSQL = strSQL & "         sel.kikaku_syurui_cd"
    strSQL = strSQL & "         ,sel.nen_ren_no"
    strSQL = strSQL & "         ,sel.eda_no"
    strSQL = strSQL & "         ,sel.jan_cd"
    strSQL = strSQL & "         ,sel.ten_cd"
    strSQL = strSQL & "         ,sel.baika_henkou_date"
    strSQL = strSQL & "         ,sel.baika"
    strSQL = strSQL & "         ,NULL AS baika_kikaku_no"
    strSQL = strSQL & "         ,sel.genka_henkou_date"
    strSQL = strSQL & "         ,sel.genka"
    strSQL = strSQL & "         ,sel.kikaku_no AS genka_kikaku_no"
   
    If zeiHanteiKbn = "0" Then  ' 税込み の場合
        '|--------------------(売価(税抜き)------------------------||--原価---||--------------------(売価(税抜き)--------------|
        '((売価(税込み) × 消費税率 ÷ (消費税率 ＋ 100)－ 消費税額) - 原価)  /  (売価(税込み) × 消費税率 ÷ (消費税率 ＋ 100)
        strSQL = strSQL & "         ,((sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) - sel.genka) / (sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) * 100 AS neire_ritu"     '値入率
    Else                        ' 税抜き の場合
        '(売価(税抜き) － 原価) ÷ 売価(税抜き)
        strSQL = strSQL & "         ,round((sel.baika- sel.genka) / sel.baika,3) * 100 AS neire_ritu "      '値入率
    End If
    
    strSQL = strSQL & "         ,'1' AS touroku_naiyou"         '売価変更(固定値)
    strSQL = strSQL & "         ,'" & strLoginUser & "' AS kousin_user_id"
    strSQL = strSQL & "         ,GETDATE() AS kousin_ymd_hms"
    strSQL = strSQL & "         ,? AS saiban_no"
    strSQL = strSQL & "         ,sel.table_kbn"
    strSQL = strSQL & "         ,ISNULL(a.neireritu_jyougen,60) AS  neireritu_jyougen"
    strSQL = strSQL & "         ,ISNULL(a.neireritu_kagen,5) AS neireritu_kagen"
    strSQL = strSQL & "     FROM"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             w_baika.kikaku_syurui_cd"
    strSQL = strSQL & "             ,w_baika.nen_ren_no"
    strSQL = strSQL & "             ,w_baika.eda_no"
    strSQL = strSQL & "             ,w_baika.jan_cd"
    strSQL = strSQL & "             ,w_baika.ten_cd"
    strSQL = strSQL & "             ,w_baika.henkou_date AS baika_henkou_date"
    strSQL = strSQL & "             ,w_baika.baika"
    strSQL = strSQL & "             ,v_siire.henkou_date AS genka_henkou_date"
    strSQL = strSQL & "             ,v_siire.genka AS genka"
    strSQL = strSQL & "             ,v_siire.kikaku_syurui_cd + v_siire.nen_ren_no + '-' + v_siire.eda_no AS kikaku_no"
    strSQL = strSQL & "             ,'1' AS table_kbn"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             t_syousen_baika_henkou w_baika WITH(READCOMMITTED)"
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 w_syousen_neire_keikoku_jan jan WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 w_baika.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "                 AND w_baika.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "                 AND w_baika.eda_no = jan.eda_no"
    strSQL = strSQL & "                 AND w_baika.jan_cd = jan.jan_cd"
    strSQL = strSQL & "                 AND w_baika.henkou_date = jan.henkou_date"
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 v_key_1 v_key WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 w_baika.jan_cd = v_key.jan_cd"
                                    '仕入変更情報
    strSQL = strSQL & "             INNER JOIN"
    'strSQL = strSQL & "                 v_siire_henkou v_siire WITH(READCOMMITTED)"
    strSQL = strSQL & "                 VMST.dbo.t_siire_henkou v_siire WITH(READCOMMITTED,INDEX(t_siire_henkou_idx1))"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_key.key_cd = v_siire.key_cd"
    strSQL = strSQL & "                 AND w_baika.ten_cd = v_siire.ten_cd"
    strSQL = strSQL & "         WHERE"
    strSQL = strSQL & "             jan.saiban_no = ?"
    strSQL = strSQL & "             AND w_baika.henkou_date >= v_siire.henkou_date"
    strSQL = strSQL & "             AND w_baika.touroku_kbn <> '9'"
    strSQL = strSQL & "             AND v_siire.touroku_kbn <> '9'"
    strSQL = strSQL & "             AND w_baika.baika <> 0"
    strSQL = strSQL & "             AND ISNULL(v_siire.genka,0) <> 0"
    strSQL = strSQL & "         ) sel"
    strSQL = strSQL & "         LEFT JOIN"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             v_key_1.jan_cd"
    strSQL = strSQL & "             ,m_neireritu.neireritu_jyougen"
    strSQL = strSQL & "             ,m_neireritu.neireritu_kagen"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             m_neireritu  WITH(READCOMMITTED)"
    strSQL = strSQL & "             LEFT JOIN"
    strSQL = strSQL & "                 v_syanai_kihon  WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_syanai_kihon.bumon_cd = m_neireritu.daibunrui_cd"
    strSQL = strSQL & "             LEFT JOIN"
    strSQL = strSQL & "                 v_key_1  WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_key_1.key_cd = v_syanai_kihon.key_cd"
    strSQL = strSQL & "         ) a"
    strSQL = strSQL & "         ON"
    strSQL = strSQL & "             a.jan_cd = sel.jan_cd"
    strSQL = strSQL & "      ) sel2"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     sel2.neire_ritu < sel2.neireritu_kagen"
    strSQL = strSQL & "     OR"
    strSQL = strSQL & "     sel2.neire_ritu > sel2.neireritu_jyougen"

    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@saiban_no2", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)


    'SQL文
    '▼店別情報とのチェック
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,baika"
    strSQL = strSQL & "     ,baika_kikaku_no"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,genka"
    strSQL = strSQL & "     ,genka_kikaku_no"
    strSQL = strSQL & "     ,neire_ritu"
    strSQL = strSQL & "     ,touroku_naiyou"
    strSQL = strSQL & "     ,kousin_user_id"
    strSQL = strSQL & "     ,kousin_ymd_hms"
    strSQL = strSQL & "     ,saiban_no"
    strSQL = strSQL & "     ,table_kbn)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     sel2.kikaku_syurui_cd"
    strSQL = strSQL & "     ,sel2.nen_ren_no"
    strSQL = strSQL & "     ,sel2.eda_no"
    strSQL = strSQL & "     ,sel2.jan_cd"
    strSQL = strSQL & "     ,sel2.ten_cd"
    strSQL = strSQL & "     ,sel2.baika_henkou_date"
    strSQL = strSQL & "     ,sel2.baika"
    strSQL = strSQL & "     ,sel2.baika_kikaku_no"
    strSQL = strSQL & "     ,sel2.genka_henkou_date"
    strSQL = strSQL & "     ,sel2.genka"
    strSQL = strSQL & "     ,sel2.genka_kikaku_no"
    strSQL = strSQL & "     ,CASE"
    strSQL = strSQL & "         WHEN ROUND(sel2.neire_ritu,1) > 9999.9 THEN 9999.9"
    strSQL = strSQL & "         WHEN ROUND(sel2.neire_ritu,1) < -9999.9 THEN -9999.9"
    strSQL = strSQL & "         ELSE ROUND(sel2.neire_ritu,1)"
    strSQL = strSQL & "         END"
    strSQL = strSQL & "     ,sel2.touroku_naiyou"
    strSQL = strSQL & "     ,sel2.kousin_user_id"
    strSQL = strSQL & "     ,sel2.kousin_ymd_hms"
    strSQL = strSQL & "     ,sel2.saiban_no"
    strSQL = strSQL & "     ,sel2.table_kbn"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     (SELECT"
    strSQL = strSQL & "         sel.kikaku_syurui_cd"
    strSQL = strSQL & "         ,sel.nen_ren_no"
    strSQL = strSQL & "         ,sel.eda_no"
    strSQL = strSQL & "         ,sel.jan_cd"
    strSQL = strSQL & "         ,sel.ten_cd"
    strSQL = strSQL & "         ,sel.baika_henkou_date"
    strSQL = strSQL & "         ,sel.baika"
    strSQL = strSQL & "         ,NULL AS baika_kikaku_no"
    strSQL = strSQL & "         ,sel.genka_henkou_date"
    strSQL = strSQL & "         ,sel.genka"
    strSQL = strSQL & "         ,sel.kikaku_no AS genka_kikaku_no"
   
    If zeiHanteiKbn = "0" Then  ' 税込み の場合
        '|--------------------(売価(税抜き)------------------------||--原価---||--------------------(売価(税抜き)--------------|
        '((売価(税込み) × 消費税率 ÷ (消費税率 ＋ 100)－ 消費税額) - 原価)  /  (売価(税込み) × 消費税率 ÷ (消費税率 ＋ 100)
        strSQL = strSQL & "         ,((sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) - sel.genka) / (sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) * 100 AS neire_ritu"     '値入率
    Else                        ' 税抜き の場合
        '(売価(税抜き) － 原価) ÷ 売価(税抜き)
        strSQL = strSQL & "         ,round((sel.baika- sel.genka) / sel.baika,3) * 100 AS neire_ritu "      '値入率
    End If
    
    strSQL = strSQL & "         ,'1' AS touroku_naiyou"         '売価変更(固定値)
    strSQL = strSQL & "         ,'" & strLoginUser & "' AS kousin_user_id"
    strSQL = strSQL & "         ,GETDATE() AS kousin_ymd_hms"
    strSQL = strSQL & "         ,? AS saiban_no"
    strSQL = strSQL & "         ,sel.table_kbn"
    strSQL = strSQL & "         ,ISNULL(a.neireritu_jyougen,60) AS  neireritu_jyougen"
    strSQL = strSQL & "         ,ISNULL(a.neireritu_kagen,5) AS neireritu_kagen"
    strSQL = strSQL & "     FROM"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             w_baika.kikaku_syurui_cd"
    strSQL = strSQL & "             ,w_baika.nen_ren_no"
    strSQL = strSQL & "             ,w_baika.eda_no"
    strSQL = strSQL & "             ,w_baika.jan_cd"
    strSQL = strSQL & "             ,w_baika.ten_cd"
    strSQL = strSQL & "             ,w_baika.henkou_date AS baika_henkou_date"
    strSQL = strSQL & "             ,w_baika.baika"
    strSQL = strSQL & "             ,'19900101' AS genka_henkou_date"            'Vマスタの変更日は1990/01/01をセット
    strSQL = strSQL & "             ,mise.genka"
    strSQL = strSQL & "             ,NULL AS kikaku_no"
    strSQL = strSQL & "             ,'0' AS table_kbn"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             t_syousen_baika_henkou w_baika WITH(READCOMMITTED)"
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 w_syousen_neire_keikoku_jan jan WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 w_baika.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "                 AND w_baika.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "                 AND w_baika.eda_no = jan.eda_no"
    strSQL = strSQL & "                 AND w_baika.jan_cd = jan.jan_cd"
    strSQL = strSQL & "                 AND w_baika.henkou_date = jan.henkou_date"
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 v_key_1 v_key WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 w_baika.jan_cd = v_key.jan_cd"
                                    '店別情報
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 v_mise_betu mise WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_key.key_cd = mise.key_cd"
    strSQL = strSQL & "                 AND w_baika.ten_cd = mise.ten_cd"
    strSQL = strSQL & "         WHERE"
    strSQL = strSQL & "             jan.saiban_no = ?"
    strSQL = strSQL & "             AND w_baika.touroku_kbn <> '9'"
    strSQL = strSQL & "             AND mise.touroku_kbn <> '9'"
    strSQL = strSQL & "             AND w_baika.baika <> 0"
    strSQL = strSQL & "             AND mise.genka <> 0"
    strSQL = strSQL & "         ) sel"
    strSQL = strSQL & "         LEFT JOIN"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             v_key_1.jan_cd"
    strSQL = strSQL & "             ,m_neireritu.neireritu_jyougen"
    strSQL = strSQL & "             ,m_neireritu.neireritu_kagen"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             m_neireritu  WITH(READCOMMITTED)"
    strSQL = strSQL & "             LEFT JOIN"
    strSQL = strSQL & "                 v_syanai_kihon  WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_syanai_kihon.bumon_cd = m_neireritu.daibunrui_cd"
    strSQL = strSQL & "             LEFT JOIN"
    strSQL = strSQL & "                 v_key_1  WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_key_1.key_cd = v_syanai_kihon.key_cd"
    strSQL = strSQL & "         ) a"
    strSQL = strSQL & "         ON"
    strSQL = strSQL & "             a.jan_cd = sel.jan_cd"
    strSQL = strSQL & "      ) sel2"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     sel2.neire_ritu < sel2.neireritu_kagen"
    strSQL = strSQL & "     OR"
    strSQL = strSQL & "     sel2.neire_ritu > sel2.neireritu_jyougen"

    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@saiban_no2", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)


    'SQL文
    '▼社内基本情報とのチェック
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,baika"
    strSQL = strSQL & "     ,baika_kikaku_no"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,genka"
    strSQL = strSQL & "     ,genka_kikaku_no"
    strSQL = strSQL & "     ,neire_ritu"
    strSQL = strSQL & "     ,touroku_naiyou"
    strSQL = strSQL & "     ,kousin_user_id"
    strSQL = strSQL & "     ,kousin_ymd_hms"
    strSQL = strSQL & "     ,saiban_no"
    strSQL = strSQL & "     ,table_kbn)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     sel2.kikaku_syurui_cd"
    strSQL = strSQL & "     ,sel2.nen_ren_no"
    strSQL = strSQL & "     ,sel2.eda_no"
    strSQL = strSQL & "     ,sel2.jan_cd"
    strSQL = strSQL & "     ,sel2.ten_cd"
    strSQL = strSQL & "     ,sel2.baika_henkou_date"
    strSQL = strSQL & "     ,sel2.baika"
    strSQL = strSQL & "     ,sel2.baika_kikaku_no"
    strSQL = strSQL & "     ,sel2.genka_henkou_date"
    strSQL = strSQL & "     ,sel2.genka"
    strSQL = strSQL & "     ,sel2.genka_kikaku_no"
    strSQL = strSQL & "     ,CASE"
    strSQL = strSQL & "         WHEN ROUND(sel2.neire_ritu,1) > 9999.9 THEN 9999.9"
    strSQL = strSQL & "         WHEN ROUND(sel2.neire_ritu,1) < -9999.9 THEN -9999.9"
    strSQL = strSQL & "         ELSE ROUND(sel2.neire_ritu,1)"
    strSQL = strSQL & "         END"
    strSQL = strSQL & "     ,sel2.touroku_naiyou"
    strSQL = strSQL & "     ,sel2.kousin_user_id"
    strSQL = strSQL & "     ,sel2.kousin_ymd_hms"
    strSQL = strSQL & "     ,sel2.saiban_no"
    strSQL = strSQL & "     ,sel2.table_kbn"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     (SELECT"
    strSQL = strSQL & "         sel.kikaku_syurui_cd"
    strSQL = strSQL & "         ,sel.nen_ren_no"
    strSQL = strSQL & "         ,sel.eda_no"
    strSQL = strSQL & "         ,sel.jan_cd"
    strSQL = strSQL & "         ,sel.ten_cd"
    strSQL = strSQL & "         ,sel.baika_henkou_date"
    strSQL = strSQL & "         ,sel.baika"
    strSQL = strSQL & "         ,NULL AS baika_kikaku_no"
    strSQL = strSQL & "         ,sel.genka_henkou_date"
    strSQL = strSQL & "         ,sel.genka"
    strSQL = strSQL & "         ,sel.kikaku_no AS genka_kikaku_no"
   
    If zeiHanteiKbn = "0" Then  ' 税込み の場合
        '|--------------------(売価(税抜き)------------------------||--原価---||--------------------(売価(税抜き)--------------|
        '((売価(税込み) × 消費税率 ÷ (消費税率 ＋ 100)－ 消費税額) - 原価)  /  (売価(税込み) × 消費税率 ÷ (消費税率 ＋ 100)
        strSQL = strSQL & "         ,((sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) - sel.genka) / (sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) * 100 AS neire_ritu"     '値入率
    Else                        ' 税抜き の場合
        '(売価(税抜き) － 原価) ÷ 売価(税抜き)
        strSQL = strSQL & "         ,round((sel.baika- sel.genka) / sel.baika,3) * 100 AS neire_ritu "      '値入率
    End If
    
    strSQL = strSQL & "         ,'1' AS touroku_naiyou"         '売価変更(固定値)
    strSQL = strSQL & "         ,'" & strLoginUser & "' AS kousin_user_id"
    strSQL = strSQL & "         ,GETDATE() AS kousin_ymd_hms"
    strSQL = strSQL & "         ,? AS saiban_no"
    strSQL = strSQL & "         ,sel.table_kbn"
    strSQL = strSQL & "         ,ISNULL(a.neireritu_jyougen,60) AS  neireritu_jyougen"
    strSQL = strSQL & "         ,ISNULL(a.neireritu_kagen,5) AS neireritu_kagen"
    strSQL = strSQL & "     FROM"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             w_baika.kikaku_syurui_cd"
    strSQL = strSQL & "             ,w_baika.nen_ren_no"
    strSQL = strSQL & "             ,w_baika.eda_no"
    strSQL = strSQL & "             ,w_baika.jan_cd"
    strSQL = strSQL & "             ,w_baika.ten_cd"
    strSQL = strSQL & "             ,w_baika.henkou_date AS baika_henkou_date"
    strSQL = strSQL & "             ,w_baika.baika"
    strSQL = strSQL & "             ,'19900101' AS genka_henkou_date"            'Vマスタの変更日は1990/01/01をセット
    strSQL = strSQL & "             ,kihon.hyoujyun_genka AS genka"
    strSQL = strSQL & "             ,NULL AS kikaku_no"
    strSQL = strSQL & "             ,'0' AS table_kbn"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             t_syousen_baika_henkou w_baika WITH(READCOMMITTED)"
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 w_syousen_neire_keikoku_jan jan WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 w_baika.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "                 AND w_baika.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "                 AND w_baika.eda_no = jan.eda_no"
    strSQL = strSQL & "                 AND w_baika.jan_cd = jan.jan_cd"
    strSQL = strSQL & "                 AND w_baika.henkou_date = jan.henkou_date"
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 v_key_1 v_key WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 w_baika.jan_cd = v_key.jan_cd"
                                    '社内基本情報
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 v_syanai_kihon kihon WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_key.key_cd = kihon.key_cd"
    strSQL = strSQL & "         WHERE"
    strSQL = strSQL & "             jan.saiban_no = ?"
    strSQL = strSQL & "             AND w_baika.touroku_kbn <> '9'"
    strSQL = strSQL & "             AND kihon.touroku_kbn <> '9'"
    strSQL = strSQL & "             AND w_baika.baika <> 0"
    strSQL = strSQL & "             AND kihon.hyoujyun_genka <> 0"
    strSQL = strSQL & "             AND w_baika.ten_cd = '0000'"
    strSQL = strSQL & "         ) sel"
    strSQL = strSQL & "         LEFT JOIN"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             v_key_1.jan_cd"
    strSQL = strSQL & "             ,m_neireritu.neireritu_jyougen"
    strSQL = strSQL & "             ,m_neireritu.neireritu_kagen"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             m_neireritu  WITH(READCOMMITTED)"
    strSQL = strSQL & "             LEFT JOIN"
    strSQL = strSQL & "                 v_syanai_kihon  WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_syanai_kihon.bumon_cd = m_neireritu.daibunrui_cd"
    strSQL = strSQL & "             LEFT JOIN"
    strSQL = strSQL & "                 v_key_1  WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_key_1.key_cd = v_syanai_kihon.key_cd"
    strSQL = strSQL & "         ) a"
    strSQL = strSQL & "         ON"
    strSQL = strSQL & "             a.jan_cd = sel.jan_cd"
    strSQL = strSQL & "      ) sel2"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     sel2.neire_ritu < sel2.neireritu_kagen"
    strSQL = strSQL & "     OR"
    strSQL = strSQL & "     sel2.neire_ritu > sel2.neireritu_jyougen"

    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@saiban_no2", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
 
    'インスタンス破棄
    objCDBAccess = Nothing
    
    'commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
             'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >*******************************************************************
'*
'*【プロシージャ名】
'*!     InsertBaikaKeyWork
'*
'*【処理概要】
'*?     商品選定WF(値入率警告KEY)WORKに対象キーを登録
'*
'*【パラメータ】
'*      Name                　　  Type                Comment
'*      intSaibanNo;              Integer;            採番NO;
'*
'*【戻り値】
'*      Name                　　  Type                Comment
'*R     なし
'*
'******************************************************************************
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)       EDP-No      説明
'*M 2008/02/29; 和田(IC);         P-32685;    商品選定ＷＦシステムの原価・売価変更登録時の値入率チェック機能追加;
'*M 2014/03/04; 和田(IC);         P-46854;    処理を分割(タイムアウト対応);
'*< Comment End >**************************************************************
Private Sub InsertBaikaKeyWork(ByVal intSaibanNo As Integer)

    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim strSQL As String
    
    On Error GoTo DefectErrorHandler

    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.InsertBaikaKeyWork(" & intSaibanNo & _
                                                             ")"
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL文
    '▼商品選定WF(仕入変更情報)のキー登録
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku_key2"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,table_kbn"
    strSQL = strSQL & "     ,syori_kbn"
    strSQL = strSQL & "     ,saiban_no)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     w_baika.kikaku_syurui_cd"
    strSQL = strSQL & "     ,w_baika.nen_ren_no"
    strSQL = strSQL & "     ,w_baika.eda_no"
    strSQL = strSQL & "     ,w_baika.jan_cd"
    strSQL = strSQL & "     ,w_baika.ten_cd"
    strSQL = strSQL & "     ,jan.henkou_date AS baika_henkou_date"
    strSQL = strSQL & "     ,w_siire.henkou_date + '2' AS genka_henkou_date"    '原価変更日＋テーブル区分
    strSQL = strSQL & "     ,'2'"      'テーブル区分"
    strSQL = strSQL & "     ,'1' AS syori_kbn"         '売価変更(固定値)
    strSQL = strSQL & "     ,jan.saiban_no"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     t_syousen_baika_henkou w_baika WITH(READCOMMITTED)"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         w_syousen_neire_keikoku_jan jan WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         w_baika.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "         AND w_baika.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "         AND w_baika.eda_no = jan.eda_no"
    strSQL = strSQL & "         AND w_baika.jan_cd = jan.jan_cd"
    strSQL = strSQL & "         AND w_baika.henkou_date = jan.henkou_date"
                            '商品選定WF(仕入変更情報)
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         t_syousen_siire_henkou w_siire WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         w_baika.kikaku_syurui_cd = w_siire.kikaku_syurui_cd"
    strSQL = strSQL & "         AND w_baika.nen_ren_no = w_siire.nen_ren_no"
    strSQL = strSQL & "         AND w_baika.eda_no = w_siire.eda_no"
    strSQL = strSQL & "         AND w_baika.jan_cd = w_siire.jan_cd"
    strSQL = strSQL & "         AND w_baika.ten_cd = w_siire.ten_cd"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     jan.saiban_no = ?"
    strSQL = strSQL & "     AND w_baika.henkou_date >= w_siire.henkou_date"
    strSQL = strSQL & "     AND w_baika.touroku_kbn <> '9'"
    strSQL = strSQL & "     AND w_siire.touroku_kbn <> '9'"
    strSQL = strSQL & "     AND w_baika.baika <> 0"
    strSQL = strSQL & "     AND w_siire.genka <> 0"

    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)


    'SQL文
    '▼仕入変更情報のキー登録
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku_key2"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,table_kbn"
    strSQL = strSQL & "     ,syori_kbn"
    strSQL = strSQL & "     ,saiban_no)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     w_baika.kikaku_syurui_cd"
    strSQL = strSQL & "     ,w_baika.nen_ren_no"
    strSQL = strSQL & "     ,w_baika.eda_no"
    strSQL = strSQL & "     ,w_baika.jan_cd"
    strSQL = strSQL & "     ,w_baika.ten_cd"
    strSQL = strSQL & "     ,jan.henkou_date AS baika_henkou_date"
    strSQL = strSQL & "     ,v_siire.henkou_date + '1' AS genka_henkou_date"    '原価変更日＋テーブル区分
    strSQL = strSQL & "     ,'1'"      'テーブル区分"
    strSQL = strSQL & "     ,'1' AS syori_kbn"         '売価変更(固定値)
    strSQL = strSQL & "     ,jan.saiban_no"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     t_syousen_baika_henkou w_baika WITH(READCOMMITTED)"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         w_syousen_neire_keikoku_jan jan WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         w_baika.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "         AND w_baika.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "         AND w_baika.eda_no = jan.eda_no"
    strSQL = strSQL & "         AND w_baika.jan_cd = jan.jan_cd"
    strSQL = strSQL & "         AND w_baika.henkou_date = jan.henkou_date"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         v_key_1 v_key WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         w_baika.jan_cd = v_key.jan_cd"
                            '仕入変更情報
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         v_siire_henkou v_siire WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         v_key.key_cd = v_siire.key_cd"
    strSQL = strSQL & "         AND w_baika.ten_cd = v_siire.ten_cd"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     jan.saiban_no = ?"
    strSQL = strSQL & "     AND w_baika.henkou_date >= v_siire.henkou_date"
    strSQL = strSQL & "     AND w_baika.touroku_kbn <> '9'"
    strSQL = strSQL & "     AND v_siire.touroku_kbn <> '9'"
    strSQL = strSQL & "     AND w_baika.baika <> 0"
    strSQL = strSQL & "     AND v_siire.genka <> 0"

    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)


    'SQL文
    '▼店別情報のキー登録
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku_key2"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,table_kbn"
    strSQL = strSQL & "     ,syori_kbn"
    strSQL = strSQL & "     ,saiban_no)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     w_baika.kikaku_syurui_cd"
    strSQL = strSQL & "     ,w_baika.nen_ren_no"
    strSQL = strSQL & "     ,w_baika.eda_no"
    strSQL = strSQL & "     ,w_baika.jan_cd"
    strSQL = strSQL & "     ,w_baika.ten_cd"
    strSQL = strSQL & "     ,jan.henkou_date AS baika_henkou_date"
    strSQL = strSQL & "     ,'19900101' + '0'AS genka_henkou_date"            'Vマスタの変更日は1990/01/01をセット 原価変更日＋テーブル区分
    strSQL = strSQL & "     ,'0'"      'テーブル区分"
    strSQL = strSQL & "     ,'1' AS syori_kbn"         '売価変更(固定値)
    strSQL = strSQL & "     ,jan.saiban_no"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     t_syousen_baika_henkou w_baika WITH(READCOMMITTED)"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         w_syousen_neire_keikoku_jan jan WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         w_baika.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "         AND w_baika.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "         AND w_baika.eda_no = jan.eda_no"
    strSQL = strSQL & "         AND w_baika.jan_cd = jan.jan_cd"
    strSQL = strSQL & "         AND w_baika.henkou_date = jan.henkou_date"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         v_key_1 v_key WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         w_baika.jan_cd = v_key.jan_cd"
                            '店別情報
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         v_mise_betu mise WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         v_key.key_cd = mise.key_cd"
    strSQL = strSQL & "         AND w_baika.ten_cd = mise.ten_cd"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     jan.saiban_no = ?"
    strSQL = strSQL & "     AND w_baika.touroku_kbn <> '9'"
    strSQL = strSQL & "     AND mise.touroku_kbn <> '9'"
    strSQL = strSQL & "     AND w_baika.baika <> 0"
    strSQL = strSQL & "     AND mise.genka <> 0"

    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)


    'SQL文
    '▼社内基本情報のキー登録
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku_key2"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,table_kbn"
    strSQL = strSQL & "     ,syori_kbn"
    strSQL = strSQL & "     ,saiban_no)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     w_baika.kikaku_syurui_cd"
    strSQL = strSQL & "     ,w_baika.nen_ren_no"
    strSQL = strSQL & "     ,w_baika.eda_no"
    strSQL = strSQL & "     ,w_baika.jan_cd"
    strSQL = strSQL & "     ,w_baika.ten_cd"
    strSQL = strSQL & "     ,jan.henkou_date"
    strSQL = strSQL & "     ,'19900101' + '0'"            'Vマスタの変更日は1990/01/01をセット 原価変更日＋テーブル区分
    strSQL = strSQL & "     ,'0'"      'テーブル区分"
    strSQL = strSQL & "     ,'1' AS syori_kbn"         '売価変更(固定値)
    strSQL = strSQL & "     ,jan.saiban_no"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     t_syousen_baika_henkou w_baika WITH(READCOMMITTED)"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         w_syousen_neire_keikoku_jan jan WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         w_baika.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "         AND w_baika.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "         AND w_baika.eda_no = jan.eda_no"
    strSQL = strSQL & "         AND w_baika.jan_cd = jan.jan_cd"
    strSQL = strSQL & "         AND w_baika.henkou_date = jan.henkou_date"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         v_key_1 v_key WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         w_baika.jan_cd = v_key.jan_cd"
                            '社内基本情報
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         v_syanai_kihon kihon WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         v_key.key_cd = kihon.key_cd"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     jan.saiban_no = ?"
    strSQL = strSQL & "     AND w_baika.touroku_kbn <> '9'"
    strSQL = strSQL & "     AND kihon.touroku_kbn <> '9'"
    strSQL = strSQL & "     AND w_baika.baika <> 0"
    strSQL = strSQL & "     AND kihon.hyoujyun_genka <> 0"
    strSQL = strSQL & "     AND w_baika.ten_cd = '0000'"

    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)


    'SQL文
    '▼直近のキーを登録
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku_key"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,table_kbn"
    strSQL = strSQL & "     ,syori_kbn"
    strSQL = strSQL & "     ,saiban_no)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     sel.kikaku_syurui_cd"
    strSQL = strSQL & "     ,sel.nen_ren_no"
    strSQL = strSQL & "     ,sel.eda_no"
    strSQL = strSQL & "     ,sel.jan_cd"
    strSQL = strSQL & "     ,sel.ten_cd"
    strSQL = strSQL & "     ,sel.baika_henkou_date"
    strSQL = strSQL & "     ,SUBSTRING(sel.genka_henkou_date,1,8) AS genka_henkou_date"
    strSQL = strSQL & "     ,sel.table_kbn"
    strSQL = strSQL & "     ,sel.syori_kbn"
    strSQL = strSQL & "     ,sel.saiban_no"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     w_syousen_neire_keikoku_key2 sel"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             kikaku_syurui_cd"
    strSQL = strSQL & "             ,nen_ren_no"
    strSQL = strSQL & "             ,eda_no"
    strSQL = strSQL & "             ,jan_cd"
    strSQL = strSQL & "             ,ten_cd"
    strSQL = strSQL & "             ,baika_henkou_date"
    strSQL = strSQL & "             ,MAX(genka_henkou_date) AS genka_henkou_date"
    strSQL = strSQL & "             ,syori_kbn"
    strSQL = strSQL & "             ,saiban_no"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             w_syousen_neire_keikoku_key2"
    strSQL = strSQL & "         WHERE"
    strSQL = strSQL & "             saiban_no = ?"
    strSQL = strSQL & "             AND syori_kbn = '1'"    '売価変更(固定値)
    strSQL = strSQL & "         GROUP BY"
    strSQL = strSQL & "             kikaku_syurui_cd"
    strSQL = strSQL & "             ,nen_ren_no"
    strSQL = strSQL & "             ,eda_no"
    strSQL = strSQL & "             ,jan_cd"
    strSQL = strSQL & "             ,ten_cd"
    strSQL = strSQL & "             ,baika_henkou_date"
    strSQL = strSQL & "             ,syori_kbn"
    strSQL = strSQL & "             ,saiban_no"
    strSQL = strSQL & "         ) sel2"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         sel.kikaku_syurui_cd = sel2.kikaku_syurui_cd"
    strSQL = strSQL & "         AND sel.nen_ren_no = sel2.nen_ren_no"
    strSQL = strSQL & "         AND sel.eda_no = sel2.eda_no"
    strSQL = strSQL & "         AND sel.jan_cd = sel2.jan_cd"
    strSQL = strSQL & "         AND sel.ten_cd = sel2.ten_cd"
    strSQL = strSQL & "         AND sel.baika_henkou_date = sel2.baika_henkou_date"
    strSQL = strSQL & "         AND sel.genka_henkou_date = sel2.genka_henkou_date"
    strSQL = strSQL & "         AND sel.syori_kbn = sel2.syori_kbn"
    strSQL = strSQL & "         AND sel.saiban_no = sel2.saiban_no"

    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)

    'インスタンス破棄
    objCDBAccess = Nothing
    
    'commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
             'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >****************************************************'
'*【プロシージャ名】
'*!    InsertNeire
'*【処理概要】
'*?    値入率警告情報の初期化
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    intSaibanNo;         採番NO
'*P    strSyoriKbn;         1:売価 2:原価
'*【戻り値】
'*    < Name Type Comment >
'*R    なし
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2008/02/27; 和田(IC); P-32685 商品選定WFシステムの原価・売価変更登録時の値入率チェック機能追加;
'*< Comment End >**************************************************'
Private Sub InsertNeire(ByVal intSaibanNo As Integer _
                          , ByVal strSyoriKbn As String)
    
    Dim strMethodCallInfo As String
    Dim strSQL As String
    Dim vntAryParam As Object
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    
    On Error GoTo DefectErrorHandler
   
    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.InsertNeire(" & intSaibanNo & _
                                                             "," & strSyoriKbn & _
                                                                    ")"
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
       
    'SQL文
    strSQL = ""
    strSQL = strSQL & " INSERT INTO t_syousen_neire_keikoku"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,baika"
    strSQL = strSQL & "     ,baika_kikaku_no"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,genka"
    strSQL = strSQL & "     ,genka_kikaku_no"
    strSQL = strSQL & "     ,neire_ritu"
    strSQL = strSQL & "     ,touroku_naiyou"
    strSQL = strSQL & "     ,kousin_user_id"
    strSQL = strSQL & "     ,kousin_ymd_hms"
    strSQL = strSQL & "     ,kakunin_kbn)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     keikoku.kikaku_syurui_cd"
    strSQL = strSQL & "     ,keikoku.nen_ren_no"
    strSQL = strSQL & "     ,keikoku.eda_no"
    strSQL = strSQL & "     ,keikoku.jan_cd"
    strSQL = strSQL & "     ,keikoku.ten_cd"
    strSQL = strSQL & "     ,CASE keikoku.baika_henkou_date"
    strSQL = strSQL & "         WHEN '19900101' THEN NULL"
    strSQL = strSQL & "         ELSE keikoku.baika_henkou_date"
    strSQL = strSQL & "         END"
    strSQL = strSQL & "     ,keikoku.baika"
    strSQL = strSQL & "     ,keikoku.baika_kikaku_no"
    strSQL = strSQL & "     ,CASE keikoku.genka_henkou_date"
    strSQL = strSQL & "         WHEN '19900101' THEN NULL"
    strSQL = strSQL & "         ELSE keikoku.genka_henkou_date"
    strSQL = strSQL & "         END"
    strSQL = strSQL & "     ,keikoku.genka"
    strSQL = strSQL & "     ,keikoku.genka_kikaku_no"
    strSQL = strSQL & "     ,keikoku.neire_ritu"
    strSQL = strSQL & "     ,keikoku.touroku_naiyou"
    strSQL = strSQL & "     ,keikoku.kousin_user_id"
    strSQL = strSQL & "     ,keikoku.kousin_ymd_hms"
    strSQL = strSQL & "     ,NULL"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     w_syousen_neire_keikoku keikoku"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         w_syousen_neire_keikoku_key w_key"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         keikoku.kikaku_syurui_cd = w_key.kikaku_syurui_cd"
    strSQL = strSQL & "         AND keikoku.nen_ren_no = w_key.nen_ren_no"
    strSQL = strSQL & "         AND keikoku.eda_no = w_key.eda_no"
    strSQL = strSQL & "         AND keikoku.jan_cd = w_key.jan_cd"
    strSQL = strSQL & "         AND keikoku.ten_cd = w_key.ten_cd"
    strSQL = strSQL & "         AND keikoku.baika_henkou_date = w_key.baika_henkou_date"
    strSQL = strSQL & "         AND keikoku.genka_henkou_date = w_key.genka_henkou_date"
    strSQL = strSQL & "         AND keikoku.table_kbn = w_key.table_kbn"
    strSQL = strSQL & "         AND keikoku.touroku_naiyou = w_key.syori_kbn"
    strSQL = strSQL & "         AND keikoku.saiban_no = w_key.saiban_no"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     w_key.saiban_no = ?"
    strSQL = strSQL & "     AND w_key.syori_kbn = ?"

    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@syori_kbn", DT_Char, 1, strSyoriKbn) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
  
    'インスタンス破棄
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >*******************************************************************
'*
'*【プロシージャ名】
'*!     InsertGenkaWork
'*
'*【処理概要】
'*?     商品選定WF(値入率警告)WORKに値入率異常分を登録
'*
'*【パラメータ】
'*      Name                　　  Type                Comment
'*      strLoginUser;             String;             ログインユーザー;
'*      intSaibanNo;              Integer;            採番NO;
'*      sngSyouhizeiRitu;         Single;             消費税率(%);
'*
'*【戻り値】
'*      Name                　　  Type                Comment
'*R     なし
'*
'******************************************************************************
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)       EDP-No      説明
'*M 2008/02/29; 和田(IC);         P-32685;    商品選定ＷＦシステムの原価・売価変更登録時の値入率チェック機能追加;
'*M 2009/04/22; Sureerat P.;      P-35471;    Mod check 値入率上限,値入率下限;
'*M 2014/01/22; 李松涛(大連) ;　P-46854; 税制改正に伴う対応(商品選定WF・新商品情報);
'*M 2014/03/04; 和田(IC);         P-46854;    処理を分割(タイムアウト対応);
'*M 2014/04/10; 和田(IC);         P-46854;    KEY登録簡易化;
'*M 2014/04/18; 和田(IC);         P-46854;    売価変更情報(タイムアウト対応);
'*< Comment End >**************************************************************
Private Sub InsertGenkaWork(ByVal strLoginUser As String _
                      , ByVal intSaibanNo As Integer _
                      , ByVal sngSyouhizeiRitu As Single)

    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim strSQL As String
    
    On Error GoTo DefectErrorHandler

    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.InsertGenkaWork(" & strLoginUser & _
                                                            "," & intSaibanNo & _
                                                            "," & sngSyouhizeiRitu & _
                                                             ")"
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")

    Dim syouhizeiRitu '消費税率
    Dim zeiHanteiKbn  '税判定区分
    Dim objVswf
    Dim objStr
    
    '商品情報登録COMのインスタンス生成
        objVswf = CreateObject("VSWF01BUS005.CVswfSyouhizei")
    
    '消費税率と税判定区分データを取得する
    objStr = objVswf.GetKaisibiJyouhou()
    syouhizeiRitu = Split(objStr, "|")(0)
    zeiHanteiKbn = Split(objStr, "|")(1)
    
        objVswf = Nothing

    'SQL文
    '▼商品選定WF(売価変更情報)とのチェック
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,baika"
    strSQL = strSQL & "     ,baika_kikaku_no"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,genka"
    strSQL = strSQL & "     ,genka_kikaku_no"
    strSQL = strSQL & "     ,neire_ritu"
    strSQL = strSQL & "     ,touroku_naiyou"
    strSQL = strSQL & "     ,kousin_user_id"
    strSQL = strSQL & "     ,kousin_ymd_hms"
    strSQL = strSQL & "     ,saiban_no"
    strSQL = strSQL & "     ,table_kbn)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     sel2.kikaku_syurui_cd"
    strSQL = strSQL & "     ,sel2.nen_ren_no"
    strSQL = strSQL & "     ,sel2.eda_no"
    strSQL = strSQL & "     ,sel2.jan_cd"
    strSQL = strSQL & "     ,sel2.ten_cd"
    strSQL = strSQL & "     ,sel2.baika_henkou_date"
    strSQL = strSQL & "     ,sel2.baika"
    strSQL = strSQL & "     ,sel2.baika_kikaku_no"
    strSQL = strSQL & "     ,sel2.genka_henkou_date"
    strSQL = strSQL & "     ,sel2.genka"
    strSQL = strSQL & "     ,sel2.genka_kikaku_no"
    strSQL = strSQL & "     ,CASE"
    strSQL = strSQL & "         WHEN ROUND(sel2.neire_ritu,1) > 9999.9 THEN 9999.9"
    strSQL = strSQL & "         WHEN ROUND(sel2.neire_ritu,1) < -9999.9 THEN -9999.9"
    strSQL = strSQL & "         ELSE ROUND(sel2.neire_ritu,1)"
    strSQL = strSQL & "         END"
    strSQL = strSQL & "     ,sel2.touroku_naiyou"
    strSQL = strSQL & "     ,sel2.kousin_user_id"
    strSQL = strSQL & "     ,sel2.kousin_ymd_hms"
    strSQL = strSQL & "     ,sel2.saiban_no"
    strSQL = strSQL & "     ,sel2.table_kbn"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     (SELECT"
    strSQL = strSQL & "         sel.kikaku_syurui_cd"
    strSQL = strSQL & "         ,sel.nen_ren_no"
    strSQL = strSQL & "         ,sel.eda_no"
    strSQL = strSQL & "         ,sel.jan_cd"
    strSQL = strSQL & "         ,sel.ten_cd"
    strSQL = strSQL & "         ,sel.baika_henkou_date"
    strSQL = strSQL & "         ,sel.baika"
    strSQL = strSQL & "         ,sel.kikaku_no AS baika_kikaku_no"
    strSQL = strSQL & "         ,sel.genka_henkou_date"
    strSQL = strSQL & "         ,sel.genka"
    strSQL = strSQL & "         ,NULL AS genka_kikaku_no"

    If zeiHanteiKbn = "0" Then  ' 税込み の場合
        '|--------------------(売価(税抜き)------------------------||--原価---||--------------------(売価(税抜き)--------------|
        '((売価(税込み) × 消費税率 ÷ (消費税率 ＋ 100)－ 消費税額) - 原価)  /  (売価(税込み) × 消費税率 ÷ (消費税率 ＋ 100)
        strSQL = strSQL & "         ,((sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) - sel.genka) / (sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) * 100 AS neire_ritu"     '値入率
    Else                        ' 税抜き の場合
        '(売価(税抜き) － 原価) ÷ 売価(税抜き)
        strSQL = strSQL & "         ,round((sel.baika- sel.genka) / sel.baika,3) * 100 AS neire_ritu "      '値入率
    End If

    strSQL = strSQL & "         ,'2' AS touroku_naiyou"   '仕入変更(固定値)
    strSQL = strSQL & "         ,'" & strLoginUser & "' AS kousin_user_id"
    strSQL = strSQL & "         ,GETDATE() AS kousin_ymd_hms"
    strSQL = strSQL & "         ,? AS saiban_no"
    strSQL = strSQL & "         ,sel.table_kbn"
    strSQL = strSQL & "         ,ISNULL(a.neireritu_jyougen,60) AS  neireritu_jyougen"
    strSQL = strSQL & "         ,ISNULL(a.neireritu_kagen,5) AS neireritu_kagen"
    strSQL = strSQL & "     FROM"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             w_siire.kikaku_syurui_cd"
    strSQL = strSQL & "             ,w_siire.nen_ren_no"
    strSQL = strSQL & "             ,w_siire.eda_no"
    strSQL = strSQL & "             ,w_siire.jan_cd"
    strSQL = strSQL & "             ,w_siire.ten_cd"
    strSQL = strSQL & "             ,w_siire.henkou_date AS genka_henkou_date"
    strSQL = strSQL & "             ,w_siire.genka"
    strSQL = strSQL & "             ,w_baika.henkou_date AS baika_henkou_date"
    strSQL = strSQL & "             ,w_baika.baika"
    strSQL = strSQL & "             ,w_baika.kikaku_syurui_cd + w_baika.nen_ren_no + '-' + w_baika.eda_no AS kikaku_no"
    strSQL = strSQL & "             ,'2' AS table_kbn"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             t_syousen_siire_henkou w_siire WITH(READCOMMITTED)"
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 w_syousen_neire_keikoku_jan jan WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 w_siire.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "                 AND w_siire.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "                 AND w_siire.eda_no = jan.eda_no"
    strSQL = strSQL & "                 AND w_siire.jan_cd = jan.jan_cd"
    strSQL = strSQL & "                 AND w_siire.henkou_date = jan.henkou_date"
                                    '商品選定WF(売価変更情報)
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 t_syousen_baika_henkou w_baika WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 w_siire.kikaku_syurui_cd = w_baika.kikaku_syurui_cd"
    strSQL = strSQL & "                 AND w_siire.nen_ren_no = w_baika.nen_ren_no"
    strSQL = strSQL & "                 AND w_siire.eda_no = w_baika.eda_no"
    strSQL = strSQL & "                 AND w_siire.jan_cd = w_baika.jan_cd"
    strSQL = strSQL & "                 AND w_siire.ten_cd = w_baika.ten_cd"
    strSQL = strSQL & "         WHERE"
    strSQL = strSQL & "             jan.saiban_no = ?"
    strSQL = strSQL & "             AND w_siire.henkou_date >= w_baika.henkou_date"
    strSQL = strSQL & "             AND w_siire.touroku_kbn <> '9'"
    strSQL = strSQL & "             AND w_baika.touroku_kbn <> '9'"
    strSQL = strSQL & "             AND ISNULL(w_siire.genka,0) <> 0"
    strSQL = strSQL & "             AND w_baika.baika <> 0"
    strSQL = strSQL & "         ) sel"
    strSQL = strSQL & "         LEFT JOIN"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             v_key_1.jan_cd"
    strSQL = strSQL & "             ,m_neireritu.neireritu_jyougen"
    strSQL = strSQL & "             ,m_neireritu.neireritu_kagen"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             m_neireritu  WITH(READCOMMITTED)"
    strSQL = strSQL & "             LEFT JOIN"
    strSQL = strSQL & "                 v_syanai_kihon  WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_syanai_kihon.bumon_cd = m_neireritu.daibunrui_cd"
    strSQL = strSQL & "             LEFT JOIN"
    strSQL = strSQL & "                 v_key_1  WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_key_1.key_cd = v_syanai_kihon.key_cd"
    strSQL = strSQL & "         ) a"
    strSQL = strSQL & "         ON"
    strSQL = strSQL & "             a.jan_cd = sel.jan_cd"
    strSQL = strSQL & "      ) sel2"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     sel2.neire_ritu < sel2.neireritu_kagen"
    strSQL = strSQL & "     OR"
    strSQL = strSQL & "     sel2.neire_ritu > sel2.neireritu_jyougen"
    
    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@saiban_no2", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    
    'SQL文
    '▼売価変更情報とのチェック
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,baika"
    strSQL = strSQL & "     ,baika_kikaku_no"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,genka"
    strSQL = strSQL & "     ,genka_kikaku_no"
    strSQL = strSQL & "     ,neire_ritu"
    strSQL = strSQL & "     ,touroku_naiyou"
    strSQL = strSQL & "     ,kousin_user_id"
    strSQL = strSQL & "     ,kousin_ymd_hms"
    strSQL = strSQL & "     ,saiban_no"
    strSQL = strSQL & "     ,table_kbn)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     sel2.kikaku_syurui_cd"
    strSQL = strSQL & "     ,sel2.nen_ren_no"
    strSQL = strSQL & "     ,sel2.eda_no"
    strSQL = strSQL & "     ,sel2.jan_cd"
    strSQL = strSQL & "     ,sel2.ten_cd"
    strSQL = strSQL & "     ,sel2.baika_henkou_date"
    strSQL = strSQL & "     ,sel2.baika"
    strSQL = strSQL & "     ,sel2.baika_kikaku_no"
    strSQL = strSQL & "     ,sel2.genka_henkou_date"
    strSQL = strSQL & "     ,sel2.genka"
    strSQL = strSQL & "     ,sel2.genka_kikaku_no"
    strSQL = strSQL & "     ,CASE"
    strSQL = strSQL & "         WHEN ROUND(sel2.neire_ritu,1) > 9999.9 THEN 9999.9"
    strSQL = strSQL & "         WHEN ROUND(sel2.neire_ritu,1) < -9999.9 THEN -9999.9"
    strSQL = strSQL & "         ELSE ROUND(sel2.neire_ritu,1)"
    strSQL = strSQL & "         END"
    strSQL = strSQL & "     ,sel2.touroku_naiyou"
    strSQL = strSQL & "     ,sel2.kousin_user_id"
    strSQL = strSQL & "     ,sel2.kousin_ymd_hms"
    strSQL = strSQL & "     ,sel2.saiban_no"
    strSQL = strSQL & "     ,sel2.table_kbn"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     (SELECT"
    strSQL = strSQL & "         sel.kikaku_syurui_cd"
    strSQL = strSQL & "         ,sel.nen_ren_no"
    strSQL = strSQL & "         ,sel.eda_no"
    strSQL = strSQL & "         ,sel.jan_cd"
    strSQL = strSQL & "         ,sel.ten_cd"
    strSQL = strSQL & "         ,sel.baika_henkou_date"
    strSQL = strSQL & "         ,sel.baika"
    strSQL = strSQL & "         ,sel.kikaku_no AS baika_kikaku_no"
    strSQL = strSQL & "         ,sel.genka_henkou_date"
    strSQL = strSQL & "         ,sel.genka"
    strSQL = strSQL & "         ,NULL AS genka_kikaku_no"

    If zeiHanteiKbn = "0" Then  ' 税込み の場合
        '|--------------------(売価(税抜き)------------------------||--原価---||--------------------(売価(税抜き)--------------|
        '((売価(税込み) × 消費税率 ÷ (消費税率 ＋ 100)－ 消費税額) - 原価)  /  (売価(税込み) × 消費税率 ÷ (消費税率 ＋ 100)
        strSQL = strSQL & "         ,((sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) - sel.genka) / (sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) * 100 AS neire_ritu"     '値入率
    Else                        ' 税抜き の場合
        '(売価(税抜き) － 原価) ÷ 売価(税抜き)
        strSQL = strSQL & "         ,round((sel.baika- sel.genka) / sel.baika,3) * 100 AS neire_ritu "      '値入率
    End If

    strSQL = strSQL & "         ,'2' AS touroku_naiyou"   '仕入変更(固定値)
    strSQL = strSQL & "         ,'" & strLoginUser & "' AS kousin_user_id"
    strSQL = strSQL & "         ,GETDATE() AS kousin_ymd_hms"
    strSQL = strSQL & "         ,? AS saiban_no"
    strSQL = strSQL & "         ,sel.table_kbn"
    strSQL = strSQL & "         ,ISNULL(a.neireritu_jyougen,60) AS  neireritu_jyougen"
    strSQL = strSQL & "         ,ISNULL(a.neireritu_kagen,5) AS neireritu_kagen"
    strSQL = strSQL & "     FROM"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             w_siire.kikaku_syurui_cd"
    strSQL = strSQL & "             ,w_siire.nen_ren_no"
    strSQL = strSQL & "             ,w_siire.eda_no"
    strSQL = strSQL & "             ,w_siire.jan_cd"
    strSQL = strSQL & "             ,w_siire.ten_cd"
    strSQL = strSQL & "             ,w_siire.henkou_date AS genka_henkou_date"
    strSQL = strSQL & "             ,w_siire.genka"
    strSQL = strSQL & "             ,v_baika.henkou_date AS baika_henkou_date"
    strSQL = strSQL & "             ,v_baika.baika"
    strSQL = strSQL & "             ,v_baika.kikaku_syurui_cd + v_baika.nen_ren_no + '-' + v_baika.eda_no AS kikaku_no"
    strSQL = strSQL & "             ,'1' AS table_kbn"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             t_syousen_siire_henkou w_siire WITH(READCOMMITTED)"
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 w_syousen_neire_keikoku_jan jan WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 w_siire.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "                 AND w_siire.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "                 AND w_siire.eda_no = jan.eda_no"
    strSQL = strSQL & "                 AND w_siire.jan_cd = jan.jan_cd"
    strSQL = strSQL & "                 AND w_siire.henkou_date = jan.henkou_date"
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 v_key_1 v_key WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 w_siire.jan_cd = v_key.jan_cd"
                                    '売価変更情報
    strSQL = strSQL & "             INNER JOIN"
    'strSQL = strSQL & "                 v_baika_henkou v_baika WITH(READCOMMITTED)"
    strSQL = strSQL & "                 VMST.dbo.t_baika_henkou v_baika WITH(READCOMMITTED,INDEX(t_baika_henkou_idx1))"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_key.key_cd = v_baika.key_cd"
    strSQL = strSQL & "                 AND w_siire.ten_cd = v_baika.ten_cd"
    strSQL = strSQL & "         WHERE"
    strSQL = strSQL & "             jan.saiban_no = ?"
    strSQL = strSQL & "             AND w_siire.henkou_date >= v_baika.henkou_date"
    strSQL = strSQL & "             AND w_siire.touroku_kbn <> '9'"
    strSQL = strSQL & "             AND v_baika.touroku_kbn <> '9'"
    strSQL = strSQL & "             AND ISNULL(w_siire.genka,0) <> 0"
    strSQL = strSQL & "             AND v_baika.baika <> 0"
    strSQL = strSQL & "         ) sel"
    strSQL = strSQL & "         LEFT JOIN"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             v_key_1.jan_cd"
    strSQL = strSQL & "             ,m_neireritu.neireritu_jyougen"
    strSQL = strSQL & "             ,m_neireritu.neireritu_kagen"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             m_neireritu  WITH(READCOMMITTED)"
    strSQL = strSQL & "             LEFT JOIN"
    strSQL = strSQL & "                 v_syanai_kihon  WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_syanai_kihon.bumon_cd = m_neireritu.daibunrui_cd"
    strSQL = strSQL & "             LEFT JOIN"
    strSQL = strSQL & "                 v_key_1  WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_key_1.key_cd = v_syanai_kihon.key_cd"
    strSQL = strSQL & "         ) a"
    strSQL = strSQL & "         ON"
    strSQL = strSQL & "             a.jan_cd = sel.jan_cd"
    strSQL = strSQL & "      ) sel2"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     sel2.neire_ritu < sel2.neireritu_kagen"
    strSQL = strSQL & "     OR"
    strSQL = strSQL & "     sel2.neire_ritu > sel2.neireritu_jyougen"
    
    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@saiban_no2", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    'SQL文
    '▼店別情報とのチェック
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,baika"
    strSQL = strSQL & "     ,baika_kikaku_no"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,genka"
    strSQL = strSQL & "     ,genka_kikaku_no"
    strSQL = strSQL & "     ,neire_ritu"
    strSQL = strSQL & "     ,touroku_naiyou"
    strSQL = strSQL & "     ,kousin_user_id"
    strSQL = strSQL & "     ,kousin_ymd_hms"
    strSQL = strSQL & "     ,saiban_no"
    strSQL = strSQL & "     ,table_kbn)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     sel2.kikaku_syurui_cd"
    strSQL = strSQL & "     ,sel2.nen_ren_no"
    strSQL = strSQL & "     ,sel2.eda_no"
    strSQL = strSQL & "     ,sel2.jan_cd"
    strSQL = strSQL & "     ,sel2.ten_cd"
    strSQL = strSQL & "     ,sel2.baika_henkou_date"
    strSQL = strSQL & "     ,sel2.baika"
    strSQL = strSQL & "     ,sel2.baika_kikaku_no"
    strSQL = strSQL & "     ,sel2.genka_henkou_date"
    strSQL = strSQL & "     ,sel2.genka"
    strSQL = strSQL & "     ,sel2.genka_kikaku_no"
    strSQL = strSQL & "     ,CASE"
    strSQL = strSQL & "         WHEN ROUND(sel2.neire_ritu,1) > 9999.9 THEN 9999.9"
    strSQL = strSQL & "         WHEN ROUND(sel2.neire_ritu,1) < -9999.9 THEN -9999.9"
    strSQL = strSQL & "         ELSE ROUND(sel2.neire_ritu,1)"
    strSQL = strSQL & "         END"
    strSQL = strSQL & "     ,sel2.touroku_naiyou"
    strSQL = strSQL & "     ,sel2.kousin_user_id"
    strSQL = strSQL & "     ,sel2.kousin_ymd_hms"
    strSQL = strSQL & "     ,sel2.saiban_no"
    strSQL = strSQL & "     ,sel2.table_kbn"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     (SELECT"
    strSQL = strSQL & "         sel.kikaku_syurui_cd"
    strSQL = strSQL & "         ,sel.nen_ren_no"
    strSQL = strSQL & "         ,sel.eda_no"
    strSQL = strSQL & "         ,sel.jan_cd"
    strSQL = strSQL & "         ,sel.ten_cd"
    strSQL = strSQL & "         ,sel.baika_henkou_date"
    strSQL = strSQL & "         ,sel.baika"
    strSQL = strSQL & "         ,sel.kikaku_no AS baika_kikaku_no"
    strSQL = strSQL & "         ,sel.genka_henkou_date"
    strSQL = strSQL & "         ,sel.genka"
    strSQL = strSQL & "         ,NULL AS genka_kikaku_no"

    If zeiHanteiKbn = "0" Then  ' 税込み の場合
        '|--------------------(売価(税抜き)------------------------||--原価---||--------------------(売価(税抜き)--------------|
        '((売価(税込み) × 消費税率 ÷ (消費税率 ＋ 100)－ 消費税額) - 原価)  /  (売価(税込み) × 消費税率 ÷ (消費税率 ＋ 100)
        strSQL = strSQL & "         ,((sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) - sel.genka) / (sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) * 100 AS neire_ritu"     '値入率
    Else                        ' 税抜き の場合
        '(売価(税抜き) － 原価) ÷ 売価(税抜き)
        strSQL = strSQL & "         ,round((sel.baika- sel.genka) / sel.baika,3) * 100 AS neire_ritu "      '値入率
    End If

    strSQL = strSQL & "         ,'2' AS touroku_naiyou"   '仕入変更(固定値)
    strSQL = strSQL & "         ,'" & strLoginUser & "' AS kousin_user_id"
    strSQL = strSQL & "         ,GETDATE() AS kousin_ymd_hms"
    strSQL = strSQL & "         ,? AS saiban_no"
    strSQL = strSQL & "         ,sel.table_kbn"
    strSQL = strSQL & "         ,ISNULL(a.neireritu_jyougen,60) AS  neireritu_jyougen"
    strSQL = strSQL & "         ,ISNULL(a.neireritu_kagen,5) AS neireritu_kagen"
    strSQL = strSQL & "     FROM"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             w_siire.kikaku_syurui_cd"
    strSQL = strSQL & "             ,w_siire.nen_ren_no"
    strSQL = strSQL & "             ,w_siire.eda_no"
    strSQL = strSQL & "             ,w_siire.jan_cd"
    strSQL = strSQL & "             ,w_siire.ten_cd"
    strSQL = strSQL & "             ,w_siire.henkou_date genka_henkou_date"
    strSQL = strSQL & "             ,w_siire.genka"
    strSQL = strSQL & "             ,'19900101' baika_henkou_date"    'Vマスタの変更日は1990/01/01をセット
    strSQL = strSQL & "             ,mise.baika"
    strSQL = strSQL & "             ,NULL AS kikaku_no"
    strSQL = strSQL & "             ,'0' AS table_kbn"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             t_syousen_siire_henkou w_siire WITH(READCOMMITTED)"
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 w_syousen_neire_keikoku_jan jan WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 w_siire.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "                 AND w_siire.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "                 AND w_siire.eda_no = jan.eda_no"
    strSQL = strSQL & "                 AND w_siire.jan_cd = jan.jan_cd"
    strSQL = strSQL & "                 AND w_siire.henkou_date = jan.henkou_date"
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 v_key_1 v_key WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 w_siire.jan_cd = v_key.jan_cd"
                                    '店別情報
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 v_mise_betu mise WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_key.key_cd = mise.key_cd"
    strSQL = strSQL & "                 AND w_siire.ten_cd = mise.ten_cd"
    strSQL = strSQL & "         WHERE"
    strSQL = strSQL & "             jan.saiban_no = ?"
    strSQL = strSQL & "             AND w_siire.touroku_kbn <> '9'"
    strSQL = strSQL & "             AND mise.touroku_kbn <> '9'"
    strSQL = strSQL & "             AND ISNULL(w_siire.genka,0) <> 0"
    strSQL = strSQL & "             AND mise.baika <> 0"
    strSQL = strSQL & "         ) sel"
    strSQL = strSQL & "         LEFT JOIN"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             v_key_1.jan_cd"
    strSQL = strSQL & "             ,m_neireritu.neireritu_jyougen"
    strSQL = strSQL & "             ,m_neireritu.neireritu_kagen"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             m_neireritu  WITH(READCOMMITTED)"
    strSQL = strSQL & "             LEFT JOIN"
    strSQL = strSQL & "                 v_syanai_kihon  WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_syanai_kihon.bumon_cd = m_neireritu.daibunrui_cd"
    strSQL = strSQL & "             LEFT JOIN"
    strSQL = strSQL & "                 v_key_1  WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_key_1.key_cd = v_syanai_kihon.key_cd"
    strSQL = strSQL & "         ) a"
    strSQL = strSQL & "         ON"
    strSQL = strSQL & "             a.jan_cd = sel.jan_cd"
    strSQL = strSQL & "      ) sel2"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     sel2.neire_ritu < sel2.neireritu_kagen"
    strSQL = strSQL & "     OR"
    strSQL = strSQL & "     sel2.neire_ritu > sel2.neireritu_jyougen"
    
    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@saiban_no2", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    'SQL文
    '▼社内基本情報とのチェック
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,baika"
    strSQL = strSQL & "     ,baika_kikaku_no"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,genka"
    strSQL = strSQL & "     ,genka_kikaku_no"
    strSQL = strSQL & "     ,neire_ritu"
    strSQL = strSQL & "     ,touroku_naiyou"
    strSQL = strSQL & "     ,kousin_user_id"
    strSQL = strSQL & "     ,kousin_ymd_hms"
    strSQL = strSQL & "     ,saiban_no"
    strSQL = strSQL & "     ,table_kbn)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     sel2.kikaku_syurui_cd"
    strSQL = strSQL & "     ,sel2.nen_ren_no"
    strSQL = strSQL & "     ,sel2.eda_no"
    strSQL = strSQL & "     ,sel2.jan_cd"
    strSQL = strSQL & "     ,sel2.ten_cd"
    strSQL = strSQL & "     ,sel2.baika_henkou_date"
    strSQL = strSQL & "     ,sel2.baika"
    strSQL = strSQL & "     ,sel2.baika_kikaku_no"
    strSQL = strSQL & "     ,sel2.genka_henkou_date"
    strSQL = strSQL & "     ,sel2.genka"
    strSQL = strSQL & "     ,sel2.genka_kikaku_no"
    strSQL = strSQL & "     ,CASE"
    strSQL = strSQL & "         WHEN ROUND(sel2.neire_ritu,1) > 9999.9 THEN 9999.9"
    strSQL = strSQL & "         WHEN ROUND(sel2.neire_ritu,1) < -9999.9 THEN -9999.9"
    strSQL = strSQL & "         ELSE ROUND(sel2.neire_ritu,1)"
    strSQL = strSQL & "         END"
    strSQL = strSQL & "     ,sel2.touroku_naiyou"
    strSQL = strSQL & "     ,sel2.kousin_user_id"
    strSQL = strSQL & "     ,sel2.kousin_ymd_hms"
    strSQL = strSQL & "     ,sel2.saiban_no"
    strSQL = strSQL & "     ,sel2.table_kbn"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     (SELECT"
    strSQL = strSQL & "         sel.kikaku_syurui_cd"
    strSQL = strSQL & "         ,sel.nen_ren_no"
    strSQL = strSQL & "         ,sel.eda_no"
    strSQL = strSQL & "         ,sel.jan_cd"
    strSQL = strSQL & "         ,sel.ten_cd"
    strSQL = strSQL & "         ,sel.baika_henkou_date"
    strSQL = strSQL & "         ,sel.baika"
    strSQL = strSQL & "         ,sel.kikaku_no AS baika_kikaku_no"
    strSQL = strSQL & "         ,sel.genka_henkou_date"
    strSQL = strSQL & "         ,sel.genka"
    strSQL = strSQL & "         ,NULL AS genka_kikaku_no"

    If zeiHanteiKbn = "0" Then  ' 税込み の場合
        '|--------------------(売価(税抜き)------------------------||--原価---||--------------------(売価(税抜き)--------------|
        '((売価(税込み) × 消費税率 ÷ (消費税率 ＋ 100)－ 消費税額) - 原価)  /  (売価(税込み) × 消費税率 ÷ (消費税率 ＋ 100)
        strSQL = strSQL & "         ,((sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) - sel.genka) / (sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) * 100 AS neire_ritu"     '値入率
    Else                        ' 税抜き の場合
        '(売価(税抜き) － 原価) ÷ 売価(税抜き)
        strSQL = strSQL & "         ,round((sel.baika- sel.genka) / sel.baika,3) * 100 AS neire_ritu "      '値入率
    End If

    strSQL = strSQL & "         ,'2' AS touroku_naiyou"   '仕入変更(固定値)
    strSQL = strSQL & "         ,'" & strLoginUser & "' AS kousin_user_id"
    strSQL = strSQL & "         ,GETDATE() AS kousin_ymd_hms"
    strSQL = strSQL & "         ,? AS saiban_no"
    strSQL = strSQL & "         ,sel.table_kbn"
    strSQL = strSQL & "         ,ISNULL(a.neireritu_jyougen,60) AS  neireritu_jyougen"
    strSQL = strSQL & "         ,ISNULL(a.neireritu_kagen,5) AS neireritu_kagen"
    strSQL = strSQL & "     FROM"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             w_siire.kikaku_syurui_cd"
    strSQL = strSQL & "             ,w_siire.nen_ren_no"
    strSQL = strSQL & "             ,w_siire.eda_no"
    strSQL = strSQL & "             ,w_siire.jan_cd"
    strSQL = strSQL & "             ,w_siire.ten_cd"
    strSQL = strSQL & "             ,w_siire.henkou_date AS genka_henkou_date"
    strSQL = strSQL & "             ,w_siire.genka"
    strSQL = strSQL & "             ,'19900101' AS baika_henkou_date"    'Vマスタの変更日は1990/01/01をセット
    strSQL = strSQL & "             ,kihon.baika"
    strSQL = strSQL & "             ,NULL AS kikaku_no"
    strSQL = strSQL & "             ,'0' AS table_kbn"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             t_syousen_siire_henkou w_siire WITH(READCOMMITTED)"
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 w_syousen_neire_keikoku_jan jan WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 w_siire.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "                 AND w_siire.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "                 AND w_siire.eda_no = jan.eda_no"
    strSQL = strSQL & "                 AND w_siire.jan_cd = jan.jan_cd"
    strSQL = strSQL & "                 AND w_siire.henkou_date = jan.henkou_date"
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 v_key_1 v_key WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 w_siire.jan_cd = v_key.jan_cd"
                                    '社内基本情報
    strSQL = strSQL & "             INNER JOIN"
    strSQL = strSQL & "                 v_syanai_kihon kihon WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_key.key_cd = kihon.key_cd"
    strSQL = strSQL & "         WHERE"
    strSQL = strSQL & "             jan.saiban_no = ?"
    strSQL = strSQL & "             AND w_siire.touroku_kbn <> '9'"
    strSQL = strSQL & "             AND kihon.touroku_kbn <> '9'"
    strSQL = strSQL & "             AND ISNULL(w_siire.genka,0) <> 0"
    strSQL = strSQL & "             AND kihon.baika <> 0"
    strSQL = strSQL & "             AND w_siire.ten_cd = '0000'"
    strSQL = strSQL & "         ) sel"
    strSQL = strSQL & "         LEFT JOIN"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             v_key_1.jan_cd"
    strSQL = strSQL & "             ,m_neireritu.neireritu_jyougen"
    strSQL = strSQL & "             ,m_neireritu.neireritu_kagen"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             m_neireritu  WITH(READCOMMITTED)"
    strSQL = strSQL & "             LEFT JOIN"
    strSQL = strSQL & "                 v_syanai_kihon  WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_syanai_kihon.bumon_cd = m_neireritu.daibunrui_cd"
    strSQL = strSQL & "             LEFT JOIN"
    strSQL = strSQL & "                 v_key_1  WITH(READCOMMITTED)"
    strSQL = strSQL & "             ON"
    strSQL = strSQL & "                 v_key_1.key_cd = v_syanai_kihon.key_cd"
    strSQL = strSQL & "         ) a"
    strSQL = strSQL & "         ON"
    strSQL = strSQL & "             a.jan_cd = sel.jan_cd"
    strSQL = strSQL & "      ) sel2"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     sel2.neire_ritu < sel2.neireritu_kagen"
    strSQL = strSQL & "     OR"
    strSQL = strSQL & "     sel2.neire_ritu > sel2.neireritu_jyougen"
    
    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@saiban_no2", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    'インスタンス破棄
    objCDBAccess = Nothing
    
    'commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
             'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >*******************************************************************
'*
'*【プロシージャ名】
'*!     InsertGenkaKeyWork
'*
'*【処理概要】
'*?     商品選定WF(値入率警告)WORKに値入率異常分を登録
'*
'*【パラメータ】
'*      Name                　　  Type                Comment
'*      intSaibanNo;              Integer;            採番NO;
'*
'*【戻り値】
'*      Name                　　  Type                Comment
'*R     なし
'*
'******************************************************************************
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)       EDP-No      説明
'*M 2008/02/29; 和田(IC);         P-32685;    商品選定ＷＦシステムの原価・売価変更登録時の値入率チェック機能追加;
'*M 2014/03/04; 和田(IC);         P-46854;    処理を分割(タイムアウト対応);
'*< Comment End >**************************************************************
Private Sub InsertGenkaKeyWork(ByVal intSaibanNo As Integer)

    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim strSQL As String
    
    On Error GoTo DefectErrorHandler

    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.InsertGenkaKeyWork(" & intSaibanNo & _
                                                             ")"
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL文
    '▼商品選定WF(売価変更情報)のキー登録
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku_key2"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,table_kbn"
    strSQL = strSQL & "     ,syori_kbn"
    strSQL = strSQL & "     ,saiban_no)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     w_siire.kikaku_syurui_cd"
    strSQL = strSQL & "     ,w_siire.nen_ren_no"
    strSQL = strSQL & "     ,w_siire.eda_no"
    strSQL = strSQL & "     ,w_siire.jan_cd"
    strSQL = strSQL & "     ,w_siire.ten_cd"
    strSQL = strSQL & "     ,w_baika.henkou_date + '2' AS baika_henkou_date"    '売価変更日＋テーブル区分
    strSQL = strSQL & "     ,jan.henkou_date AS genka_henkou_date"
    strSQL = strSQL & "     ,'2'"      'テーブル区分"
    strSQL = strSQL & "     ,'2' AS syori_kbn"         '原価変更(固定値)
    strSQL = strSQL & "     ,jan.saiban_no"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     t_syousen_siire_henkou w_siire WITH(READCOMMITTED)"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         w_syousen_neire_keikoku_jan jan WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         w_siire.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "         AND w_siire.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "         AND w_siire.eda_no = jan.eda_no"
    strSQL = strSQL & "         AND w_siire.jan_cd = jan.jan_cd"
    strSQL = strSQL & "         AND w_siire.henkou_date = jan.henkou_date"
                            '商品選定WF(売価変更情報)
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         t_syousen_baika_henkou w_baika WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         w_siire.kikaku_syurui_cd = w_baika.kikaku_syurui_cd"
    strSQL = strSQL & "         AND w_siire.nen_ren_no = w_baika.nen_ren_no"
    strSQL = strSQL & "         AND w_siire.eda_no = w_baika.eda_no"
    strSQL = strSQL & "         AND w_siire.jan_cd = w_baika.jan_cd"
    strSQL = strSQL & "         AND w_siire.ten_cd = w_baika.ten_cd"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     jan.saiban_no = ?"
    strSQL = strSQL & "     AND w_siire.henkou_date >= w_baika.henkou_date"
    strSQL = strSQL & "     AND w_siire.touroku_kbn <> '9'"
    strSQL = strSQL & "     AND w_baika.touroku_kbn <> '9'"
    strSQL = strSQL & "     AND w_siire.genka <> 0"
    strSQL = strSQL & "     AND w_baika.baika <> 0"
    
    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    
    'SQL文
    '▼売価変更情報のキー登録
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku_key2"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,table_kbn"
    strSQL = strSQL & "     ,syori_kbn"
    strSQL = strSQL & "     ,saiban_no)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     w_siire.kikaku_syurui_cd"
    strSQL = strSQL & "     ,w_siire.nen_ren_no"
    strSQL = strSQL & "     ,w_siire.eda_no"
    strSQL = strSQL & "     ,w_siire.jan_cd"
    strSQL = strSQL & "     ,w_siire.ten_cd"
    strSQL = strSQL & "     ,v_baika.henkou_date + '1' "    '売価変更日＋テーブル区分
    strSQL = strSQL & "     ,jan.henkou_date"
    strSQL = strSQL & "     ,'1'"      'テーブル区分"
    strSQL = strSQL & "     ,'2' AS syori_kbn"         '原価変更(固定値)
    strSQL = strSQL & "     ,jan.saiban_no"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     t_syousen_siire_henkou w_siire WITH(READCOMMITTED)"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         w_syousen_neire_keikoku_jan jan WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         w_siire.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "         AND w_siire.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "         AND w_siire.eda_no = jan.eda_no"
    strSQL = strSQL & "         AND w_siire.jan_cd = jan.jan_cd"
    strSQL = strSQL & "         AND w_siire.henkou_date = jan.henkou_date"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         v_key_1 v_key WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         w_siire.jan_cd = v_key.jan_cd"
                            '売価変更情報
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         v_baika_henkou v_baika WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         v_key.key_cd = v_baika.key_cd"
    strSQL = strSQL & "         AND w_siire.ten_cd = v_baika.ten_cd"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     jan.saiban_no = ?"
    strSQL = strSQL & "     AND w_siire.henkou_date >= v_baika.henkou_date"
    strSQL = strSQL & "     AND w_siire.touroku_kbn <> '9'"
    strSQL = strSQL & "     AND v_baika.touroku_kbn <> '9'"
    strSQL = strSQL & "     AND w_siire.genka <> 0"
    strSQL = strSQL & "     AND v_baika.baika <> 0"
    
    
    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    
    'SQL文
    '▼店別情報のキー登録
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku_key2"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,table_kbn"
    strSQL = strSQL & "     ,syori_kbn"
    strSQL = strSQL & "     ,saiban_no)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     w_siire.kikaku_syurui_cd"
    strSQL = strSQL & "     ,w_siire.nen_ren_no"
    strSQL = strSQL & "     ,w_siire.eda_no"
    strSQL = strSQL & "     ,w_siire.jan_cd"
    strSQL = strSQL & "     ,w_siire.ten_cd"
    strSQL = strSQL & "     ,'19900101' + '0'"    'Vマスタの変更日は1990/01/01をセット 売価変更日＋テーブル区分
    strSQL = strSQL & "     ,jan.henkou_date"
    strSQL = strSQL & "     ,'0'"      'テーブル区分"
    strSQL = strSQL & "     ,'2' AS syori_kbn"         '原価変更(固定値)
    strSQL = strSQL & "     ,jan.saiban_no"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     t_syousen_siire_henkou w_siire WITH(READCOMMITTED)"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         w_syousen_neire_keikoku_jan jan WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         w_siire.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "         AND w_siire.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "         AND w_siire.eda_no = jan.eda_no"
    strSQL = strSQL & "         AND w_siire.jan_cd = jan.jan_cd"
    strSQL = strSQL & "         AND w_siire.henkou_date = jan.henkou_date"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         v_key_1 v_key WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         w_siire.jan_cd = v_key.jan_cd"
                            '店別情報
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         v_mise_betu mise WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         v_key.key_cd = mise.key_cd"
    strSQL = strSQL & "         AND w_siire.ten_cd = mise.ten_cd"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     jan.saiban_no = ?"
    strSQL = strSQL & "     AND w_siire.touroku_kbn <> '9'"
    strSQL = strSQL & "     AND mise.touroku_kbn <> '9'"
    strSQL = strSQL & "     AND w_siire.genka <> 0"
    strSQL = strSQL & "     AND mise.baika <> 0"
    
    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    
    'SQL文
    '▼社内基本情報のキー登録
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku_key2"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,table_kbn"
    strSQL = strSQL & "     ,syori_kbn"
    strSQL = strSQL & "     ,saiban_no)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     w_siire.kikaku_syurui_cd"
    strSQL = strSQL & "     ,w_siire.nen_ren_no"
    strSQL = strSQL & "     ,w_siire.eda_no"
    strSQL = strSQL & "     ,w_siire.jan_cd"
    strSQL = strSQL & "     ,w_siire.ten_cd"
    strSQL = strSQL & "     ,'19900101' + '0'"    'Vマスタの変更日は1990/01/01をセット 売価変更日＋テーブル区分
    strSQL = strSQL & "     ,jan.henkou_date"
    strSQL = strSQL & "     ,'0'"      'テーブル区分"
    strSQL = strSQL & "     ,'2' AS syori_kbn"         '原価変更(固定値)
    strSQL = strSQL & "     ,jan.saiban_no"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     t_syousen_siire_henkou w_siire WITH(READCOMMITTED)"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         w_syousen_neire_keikoku_jan jan WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         w_siire.kikaku_syurui_cd = jan.kikaku_syurui_cd"
    strSQL = strSQL & "         AND w_siire.nen_ren_no = jan.nen_ren_no"
    strSQL = strSQL & "         AND w_siire.eda_no = jan.eda_no"
    strSQL = strSQL & "         AND w_siire.jan_cd = jan.jan_cd"
    strSQL = strSQL & "         AND w_siire.henkou_date = jan.henkou_date"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         v_key_1 v_key WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         w_siire.jan_cd = v_key.jan_cd"
                            '社内基本情報
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         v_syanai_kihon kihon WITH(READCOMMITTED)"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         v_key.key_cd = kihon.key_cd"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     jan.saiban_no = ?"
    strSQL = strSQL & "     AND w_siire.touroku_kbn <> '9'"
    strSQL = strSQL & "     AND kihon.touroku_kbn <> '9'"
    strSQL = strSQL & "     AND w_siire.genka <> 0"
    strSQL = strSQL & "     AND kihon.baika <> 0"
    strSQL = strSQL & "     AND w_siire.ten_cd = '0000'"
    
    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    
    '▼直近のキーを登録
    strSQL = ""
    strSQL = strSQL & " INSERT INTO w_syousen_neire_keikoku_key"
    strSQL = strSQL & "     (kikaku_syurui_cd"
    strSQL = strSQL & "     ,nen_ren_no"
    strSQL = strSQL & "     ,eda_no"
    strSQL = strSQL & "     ,jan_cd"
    strSQL = strSQL & "     ,ten_cd"
    strSQL = strSQL & "     ,baika_henkou_date"
    strSQL = strSQL & "     ,genka_henkou_date"
    strSQL = strSQL & "     ,table_kbn"
    strSQL = strSQL & "     ,syori_kbn"
    strSQL = strSQL & "     ,saiban_no)"
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     sel.kikaku_syurui_cd"
    strSQL = strSQL & "     ,sel.nen_ren_no"
    strSQL = strSQL & "     ,sel.eda_no"
    strSQL = strSQL & "     ,sel.jan_cd"
    strSQL = strSQL & "     ,sel.ten_cd"
    strSQL = strSQL & "     ,SUBSTRING(sel.baika_henkou_date,1,8) AS baika_henkou_date"
    strSQL = strSQL & "     ,sel.genka_henkou_date"
    strSQL = strSQL & "     ,sel.table_kbn"
    strSQL = strSQL & "     ,sel.syori_kbn"
    strSQL = strSQL & "     ,sel.saiban_no"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     w_syousen_neire_keikoku_key2 sel"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             kikaku_syurui_cd"
    strSQL = strSQL & "             ,nen_ren_no"
    strSQL = strSQL & "             ,eda_no"
    strSQL = strSQL & "             ,jan_cd"
    strSQL = strSQL & "             ,ten_cd"
    strSQL = strSQL & "             ,MAX(baika_henkou_date) AS baika_henkou_date"
    strSQL = strSQL & "             ,genka_henkou_date"
    strSQL = strSQL & "             ,syori_kbn"
    strSQL = strSQL & "             ,saiban_no"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             w_syousen_neire_keikoku_key2"
    strSQL = strSQL & "         WHERE"
    strSQL = strSQL & "             saiban_no = ?"
    strSQL = strSQL & "             AND syori_kbn = '2'"    '原価変更(固定値)
    strSQL = strSQL & "         GROUP BY"
    strSQL = strSQL & "             kikaku_syurui_cd"
    strSQL = strSQL & "             ,nen_ren_no"
    strSQL = strSQL & "             ,eda_no"
    strSQL = strSQL & "             ,jan_cd"
    strSQL = strSQL & "             ,ten_cd"
    strSQL = strSQL & "             ,genka_henkou_date"
    strSQL = strSQL & "             ,syori_kbn"
    strSQL = strSQL & "             ,saiban_no"
    strSQL = strSQL & "         ) sel2"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         sel.kikaku_syurui_cd = sel2.kikaku_syurui_cd"
    strSQL = strSQL & "         AND sel.nen_ren_no = sel2.nen_ren_no"
    strSQL = strSQL & "         AND sel.eda_no = sel2.eda_no"
    strSQL = strSQL & "         AND sel.jan_cd = sel2.jan_cd"
    strSQL = strSQL & "         AND sel.ten_cd = sel2.ten_cd"
    strSQL = strSQL & "         AND sel.baika_henkou_date = sel2.baika_henkou_date"
    strSQL = strSQL & "         AND sel.genka_henkou_date = sel2.genka_henkou_date"
    strSQL = strSQL & "         AND sel.syori_kbn = sel2.syori_kbn"
    strSQL = strSQL & "         AND sel.saiban_no = sel2.saiban_no"

    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    'インスタンス破棄
    objCDBAccess = Nothing
    
    'commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
             'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >*******************************************************************
'*
'*【プロシージャ名】
'*!     GetNVHInfo
'*
'*【処理概要】
'*?     NVH店舗の取得。
'*
'*【パラメータ】
'*      Name                　　  Type                Comment
'*      strJanCd;                 String;             JANコード;
'*
'*【戻り値】
'*      Name                　　  Type                Comment
'*R     GetNVHInfo;               ADODB.Recordset;   レコードセット;
'*
'******************************************************************************
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)       EDP-No      説明
'*M 2008/09/19; 和田(IC);         P-34312;    ＮＶＨ店舗への商品マスター変更登録作業の簡素化;
'*< Comment End >**************************************************************
Private Function GetNVHInfo(ByVal strJanCd As String) As ADODB.Recordset
                               
    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    Dim objCDBAccess    As COMDBUtil.CDBAccess
    Dim strSQL          As String
    
    On Error GoTo DefectErrorHandler

    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.GetNVHInfo(" & strJanCd & _
                                                            ")"
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL文
    strSQL = ""
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     nvh.ten_cd"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     m_nvh_tenpo nvh WITH(READCOMMITTED)"
    strSQL = strSQL & " INNER JOIN"
    strSQL = strSQL & "     v_syanai_kihon syanai"
    strSQL = strSQL & " ON"
    strSQL = strSQL & "     nvh.atukai_ptn_cd = syanai.atukai_ptn_cd"
    strSQL = strSQL & " INNER JOIN"
    strSQL = strSQL & "     v_key_1 key1"
    strSQL = strSQL & " ON"
    strSQL = strSQL & "     syanai.key_cd = key1.key_cd"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     key1.jan_cd=?"

    'MakeParamの作成
    vntAryParam = Array(MakeParam("@strJanCd", DT_Char, 13, strJanCd))
    
    '切断Recordset取得
        GetNVHInfo = objCDBAccess.GetDCRecordset("VSWF", strSQL, vntAryParam)
    
    'インスタンス破棄
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Function

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
             'Err.Number, Err.Source, Err.Description
End Function

'*< TOSTEM >****************************************************'
'*【プロシージャ名】
'*!    FncCheckDB_Group
'*【処理概要】
'*?    グループ品チェック(グループ品の場合エラー)
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strJanCd;         JANコード
'*【戻り値】
'*    < Name Type Comment >
'*R    なし
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2009/05/12; 和田; P-35471 新規作成;
'*< Comment End >**************************************************'
Private Function FncCheckDB_Group(ByVal strJanCd As String) As Boolean
    
    Dim strMethodCallInfo As String
    Dim strSQL As String
    Dim vntAryParam As Object
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim rsGrp As ADODB.Recordset
    On Error GoTo DefectErrorHandler
   
    FncCheckDB_Group = True
    
    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncCheckDB_Group(" & strJanCd & _
                                                                    ")"
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
       
    'SQL文
    strSQL = ""
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     tanpin_kbn"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     v_syanai_kihon sya"
    strSQL = strSQL & "     INNER JOIN"
    strSQL = strSQL & "         v_key_1 key1"
    strSQL = strSQL & "     ON"
    strSQL = strSQL & "         sya.key_cd = key1.key_cd"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     key1.jan_cd = ?"

    'MakeParamの作成
    vntAryParam = Array(MakeParam("@strJanCd", DT_Char, 13, strJanCd))
    
    '切断Recordset取得
        rsGrp = objCDBAccess.GetDCRecordset("VSWF", strSQL, vntAryParam)
    
    If Not rsGrp.EOF Then
        If rsGrp.Fields("tanpin_kbn").Value = "1" Then
            FncCheckDB_Group = False
        End If
    End If
    
    'インスタンス破棄
    rsGrp.Close
        rsGrp = Nothing
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Function

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
             'Err.Number, Err.Source, Err.Description
End Function
'******************************************************************'
'*【プロシージャ名】
'*!    UpdGenka_HattuTani
'*【処理概要】
'*?    INSERT文作成（商品選定WF(戻し価格)）
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSVデータを１行分格納する配列へのポインタ
'*【戻り値】
'*    < Name Type Comment >
'*R
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2009/06/08; Supattra T(Tostem Thai); 新規作成;
'*< Comment End >**************************************************'
Private Sub UpdGenka_HattuTani(ByVal arryData As Object)

    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim strMethodCallInfo As String
    Dim lstrKikakuSyurui As String
    lstrKikakuSyurui = pstrKikakuSyurui
    
    Dim lstrKikakuNenrenban As String
    lstrKikakuNenrenban = pstrKikakuNenrenban
    
    Dim lstrKikakuEdaban As String
    lstrKikakuEdaban = pstrKikakuEdaban
    
    Dim lstrTantousyaId As String
    lstrTantousyaId = pstrTantousyaId
    
    Dim lstrNow As Object
    lstrNow = Format(Now(), "yyyy/mm/dd hh:nn:ss")
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.UpdGenka_HattuTani()"
    
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    On Error GoTo DefectErrorHandler
    
    '事業所マスタ(v_jigyousyo)より、店舗を取得する。
    Dim strSQL  As String

    'strMethodCallInfo = App.Title & ".VswfIkatuTou.UpdGenka_HattuTani(一括登録)"
    
    
    strSQL = ""
    strSQL = strSQL & " UPDATE t_syousen_siire_henkou  " & vbCrLf
    '2009.06.25 Mod by Supattra T.(Tostem Thai)Start.
    'strSQL = strSQL & " SET genka = SH.genka " & vbCrLf
    'strSQL = strSQL & " ,hattyuu_tani = SH.hattyuu_tani " & vbCrLf
    strSQL = strSQL & " SET genka = MB.genka " & vbCrLf
    strSQL = strSQL & " ,hattyuu_tani = MB.hattyuu_tani " & vbCrLf
    '2009.06.25 Mod by Supattra T.(Tostem Thai)End.
    strSQL = strSQL & " ,kousin_user_id =  '" & lstrTantousyaId & "'" & vbCrLf
    strSQL = strSQL & " ,kousin_ymd_hms =  '" & lstrNow & "'" & vbCrLf
    strSQL = strSQL & " FROM " & vbCrLf
    strSQL = strSQL & "  t_syousen_siire_henkou SH " & vbCrLf
    strSQL = strSQL & " INNER JOIN " & vbCrLf
    strSQL = strSQL & "  v_mise_betu MB " & vbCrLf
    strSQL = strSQL & " ON MB.ten_cd= SH.ten_cd " & vbCrLf
    strSQL = strSQL & " INNER JOIN " & vbCrLf
    strSQL = strSQL & "  v_key_1 K1  " & vbCrLf
    strSQL = strSQL & " ON K1.jan_cd = SH.jan_cd " & vbCrLf
    strSQL = strSQL & " AND K1.key_cd = MB.key_cd " & vbCrLf
    strSQL = strSQL & " WHERE SH.genka > MB.genka " & vbCrLf
    strSQL = strSQL & " AND SH.kikaku_syurui_cd = '" & lstrKikakuSyurui & "'" & vbCrLf
    strSQL = strSQL & " AND SH.nen_ren_no = '" & lstrKikakuNenrenban & "'" & vbCrLf
    strSQL = strSQL & " AND SH.eda_no = '" & lstrKikakuEdaban & "'" & vbCrLf

    Call objCDBAccess.ExecuteSql("VSWF", strSQL)

    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If

  
    Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
    
    'If Not GetObjectContext Is Nothing Then
        'GetObjectContext.SetAbort
    End If

    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >****************************************************'
'*【プロシージャ名】
'*!    DeleteJyogaiTenpo
'*【処理概要】
'*?    変更除外で作成された店舗のNULL情報を削除する
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*【パラメータ】
'*    < Name Type Comment >
'*P    strKikakuNo;         企画No
'*【戻り値】
'*    < Name Type Comment >
'*R    なし
'******************************************************************'
'*【変更履歴】
'*    < 日付 名前(会社名) 説明 >
'*M    2009/07/23; 和田(IC); P-36097 期間限定販促商品（”まる得”など）のワークフローへの登録システム構築;
'*< Comment End >**************************************************'
Public Sub DeleteJyogaiTenpo(ByVal strKikakuNo As String)
    
    Dim strMethodCallInfo As String
    Dim strSQL As String
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim vntAryParam As Object
    
    On Error GoTo DefectErrorHandler
   
    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.DeleteJyogaiTenpo(" & strKikakuNo & ")"
    
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
       
    'MakeParamの作成
    vntAryParam = Array(MakeParam("@KikakuSyuruiCd", DT_Char, 3, Left(strKikakuNo, 3)) _
                        , MakeParam("@NenRenNo", DT_Char, 10, Mid(strKikakuNo, 4, 10)) _
                        , MakeParam("@EdaNo", DT_Char, 3, Right(strKikakuNo, 3)) _
                        )
    
    '▼商品選定WF(売価変更情報)削除
    'SQL文
    strSQL = ""
    strSQL = strSQL & " DELETE FROM"
    strSQL = strSQL & "     t_syousen_baika_henkou"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     kikaku_syurui_cd = ?"
    strSQL = strSQL & "     AND nen_ren_no = ?"
    strSQL = strSQL & "     AND eda_no = ?"
    strSQL = strSQL & "     AND baika IS NULL"

    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
  
    '▼商品選定WF(仕入変更情報)削除
    'SQL文
    strSQL = ""
    strSQL = strSQL & " DELETE FROM"
    strSQL = strSQL & "     t_syousen_siire_henkou"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     kikaku_syurui_cd = ?"
    strSQL = strSQL & "     AND nen_ren_no = ?"
    strSQL = strSQL & "     AND eda_no = ?"
    strSQL = strSQL & "     AND (torihikisaki_cd IS NULL"
    strSQL = strSQL & "         AND hattyuu_tani IS NULL"
    strSQL = strSQL & "         AND genka IS NULL"
    strSQL = strSQL & "         )"

    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
  
    '▼商品選定DB(センターコード変更情報)削除
    'SQL文
    strSQL = ""
    strSQL = strSQL & " DELETE FROM"
    strSQL = strSQL & "     t_syousen_center_cd_henkou"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     kikaku_syurui_cd = ?"
    strSQL = strSQL & "     AND nen_ren_no = ?"
    strSQL = strSQL & "     AND eda_no = ?"
    strSQL = strSQL & "     AND center_kbn IS NULL"

    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    '▼戻し価格削除
    'SQL文
    strSQL = ""
    strSQL = strSQL & " DELETE FROM"
    strSQL = strSQL & "     t_modosi_kakaku"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     nengetu ="
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "             CONVERT(CHAR(6),tenkai_kaisi_ym_date,112)"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             t_kikaku_jyouhou WITH (READCOMMITTED)"
    strSQL = strSQL & "         WHERE"
    strSQL = strSQL & "             kikaku_syurui_cd = ?"
    strSQL = strSQL & "             AND nen_ren_no = ?"
    strSQL = strSQL & "             AND eda_no = ?"
    strSQL = strSQL & "         )"
    strSQL = strSQL & "         AND genka IS NULL"
    strSQL = strSQL & "         AND baika IS NULL"
    strSQL = strSQL & "         AND hattyuu_tani IS NULL"

    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
  
    'インスタンス破棄
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >*******************************************************************'
'*【プロシージャ名】
'*!     UpdateNVHModosi
'*【処理概要】
'*?     NVH店舗変更情報登録処理
'*【パラメータ】
'*      Name                Type            Comment
'*P     strKikakuNo;        <String>;       企画NO;
'*P     strJanCd;　         <String>;       JANコード;
'*P     strModosiGenka;     <String>;       戻し原価;
'*P     strModosiBaika;     <String>;       戻し売価;
'*P     strModosiHattyuuTani;     <String>;       戻し発注単位;
'*P     strPOPSize;         <String>;       POPサイズ;
'*P     strTenCd;           <String>;       店コード;
'*P     strTantousyaId;     <String>;       担当者ID;
'*【戻り値】
'*      Name                    Type        Comment
'*R     <UpdateNVHModosi>;            <Boolean>;  成功/失敗
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)            EDP-No      説明
'*M 2009/08/06; 和田(IC);              P-36097;    新規作成;
'*M 2010/03/26; 和田(IC);              P-37515;    変更除外店舗登録がない場合の対応;
'*< Comment End >**************************************************************
Public Function UpdateNVHModosi(ByVal strKikakuNo As String, ByVal strJanCd As String, _
                    ByVal strModosiGenka As String, ByVal strModosiBaika As String, _
                    ByVal strModosiHattyuuTani As String, ByVal strPOPSize As String, _
                    ByVal strTenCd As String, ByVal strTantousyaId As String) As Boolean

    Dim objCDBAccess            As COMDBUtil.CDBAccess  'DBオブジェクト
    Dim strSQL                  As String   'SQL
    Dim strMethodCallInfo       As String
    Dim strKikakuSyuruiCD       As String
    Dim strNenRenNO             As String
    Dim strEdaNO                As String
    Dim lngSqlRtn               As Long
    Dim objRs                   As ADODB.Recordset
    Dim objNengetu              As ADODB.Recordset
    Dim objJyogai               As ADODB.Recordset
    Dim strNengetu              As String
    Dim vntAryParam             As Object
    Dim strGenkaJyogaiKbn       As String   '原価除外区分
    Dim strbaikaJyogaiKbn       As String   '売価除外区分
    Dim strHattyuuTaniJyogaiKbn     As String   '発注単位除外区分

    On Error GoTo DefectErrorHandler

    'strMethodCallInfo = App.Title & ".CHacTenCom.UpdateNVHModosi(" & _
                        strKikakuNo & "," & strJanCd & "," & _
                        strModosiGenka & "," & strModosiBaika & "," & _
                        strModosiHattyuuTani & "," & strPOPSize & "," & _
                        strTenCd & "," & strTantousyaId & "," & ")"
                            
    '除外店舗情報の取得
        objJyogai = GetJyogaiKbn(strKikakuNo, strTenCd)
    
    If Not objJyogai.EOF Then
        '原価除外区分
        strGenkaJyogaiKbn = objJyogai.Fields("genka_jyogai_kbn").Value
        '売価除外区分
        strbaikaJyogaiKbn = objJyogai.Fields("baika_jyogai_kbn").Value
        '発注単位除外区分
        strHattyuuTaniJyogaiKbn = objJyogai.Fields("hattyuu_tani_jyogai_kbn").Value
    Else
        '原価除外区分
        strGenkaJyogaiKbn = "0"
        '売価除外区分
        strbaikaJyogaiKbn = "0"
        '発注単位除外区分
        strHattyuuTaniJyogaiKbn = "0"
    End If

    'オブジェクト破棄
    objJyogai.Close
        objJyogai = Nothing

    If strGenkaJyogaiKbn = "1" Then strModosiGenka = ""
    If strbaikaJyogaiKbn = "1" Then strModosiBaika = ""
    If strHattyuuTaniJyogaiKbn = "1" Then strModosiHattyuuTani = ""
    
    '企画NOから企画種類コード、年連番、枝番を取得
    strKikakuSyuruiCD = Left(strKikakuNo, 3)
    strNenRenNO = Mid(strKikakuNo, 4, 10)
    strEdaNO = Right(strKikakuNo, 3)
       
    'インスタンス作成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
                
    '年月の取得
    strSQL = ""
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     CONVERT(CHAR(6),tenkai_kaisi_ym_date,112) nengetu"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     t_kikaku_jyouhou WITH (READCOMMITTED)"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     kikaku_syurui_cd = ?"
    strSQL = strSQL & "     AND nen_ren_no = ?"
    strSQL = strSQL & "     AND eda_no = ?"
    
    vntAryParam = Array(MakeParam("@KikakuSyuruiCd", DT_Char, 3, strKikakuSyuruiCD) _
                        , MakeParam("@NenRenNO", DT_Char, 10, strNenRenNO) _
                        , MakeParam("@EdaNO", DT_Char, 3, strEdaNO) _
                        )
    
        objNengetu = objCDBAccess.GetDCRecordset("VSWF", strSQL, vntAryParam)
    strNengetu = objNengetu("nengetu").Value
    
    'オブジェクト破棄
    objNengetu.Close
        objNengetu = Nothing
    
    '存在チェック
        objRs = GetModosiCount(strNengetu, strJanCd, strTenCd)
    If objRs.EOF Then
        strSQL = ""
        strSQL = strSQL & " INSERT INTO t_modosi_kakaku"
        strSQL = strSQL & "     (nengetu"
        strSQL = strSQL & "     ,jan_cd"
        strSQL = strSQL & "     ,ten_cd"
        strSQL = strSQL & "     ,genka"
        strSQL = strSQL & "     ,baika"
        strSQL = strSQL & "     ,hattyuu_tani"
        strSQL = strSQL & "     ,pop_size"
        strSQL = strSQL & "     ,touroku_user_id"
        strSQL = strSQL & "     ,kousin_user_id"
        strSQL = strSQL & "     ,touroku_ymd_hms"
        strSQL = strSQL & "     ,kousin_ymd_hms)"
        strSQL = strSQL & " VALUES"
        strSQL = strSQL & "     ('" & strNengetu & "'"
        strSQL = strSQL & "     ,'" & strJanCd & "'"
        strSQL = strSQL & "     ,'" & strTenCd & "'"
        If strModosiGenka <> "" Then
            strSQL = strSQL & "     ,'" & strModosiGenka & "'"
        Else
            strSQL = strSQL & "     ,NULL"
        End If
        If strModosiBaika <> "" Then
            strSQL = strSQL & "     ,'" & strModosiBaika & "'"
        Else
            strSQL = strSQL & "     ,NULL"
        End If
        If strModosiHattyuuTani <> "" Then
            strSQL = strSQL & "     ,'" & strModosiHattyuuTani & "'"
        Else
            strSQL = strSQL & "     ,NULL"
        End If
        If strPOPSize <> "" Then
            strSQL = strSQL & "     ,'" & strPOPSize & "'"
        Else
            strSQL = strSQL & "     ,NULL"
        End If
        strSQL = strSQL & "     ,'" & strTantousyaId & "'"
        strSQL = strSQL & "     ,'" & strTantousyaId & "'"
        strSQL = strSQL & "     ,GETDATE()"
        strSQL = strSQL & "     ,GETDATE()"
        strSQL = strSQL & "     )"
    
    Else
        strSQL = ""
        strSQL = strSQL & " UPDATE"
        strSQL = strSQL & "     t_modosi_kakaku"
        strSQL = strSQL & " SET"
        strSQL = strSQL & "     kousin_user_id = '" & strTantousyaId & "'"
        strSQL = strSQL & "     ,kousin_ymd_hms = GETDATE()"
        If strModosiGenka <> "" Then
            strSQL = strSQL & "     ,genka = " & strModosiGenka
        End If
        If strModosiBaika <> "" Then
            strSQL = strSQL & "     ,baika = " & strModosiBaika
        End If
        If strModosiHattyuuTani <> "" Then
            strSQL = strSQL & "     ,hattyuu_tani = " & strModosiHattyuuTani
        End If
        If strPOPSize <> "" Then
            strSQL = strSQL & "     ,pop_size = '" & strPOPSize & "'"
        End If
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & "     nengetu = '" & strNengetu & "'"
        strSQL = strSQL & "     AND jan_cd = '" & strJanCd & "'"
        strSQL = strSQL & "     AND ten_cd = '" & strTenCd & "'"
    
    End If
    
    'SQL実行
    lngSqlRtn = objCDBAccess.ExecuteSql("VSWF", strSQL)
    
    'インスタンス破棄
    objRs.Close
        objRs = Nothing
    objCDBAccess = Nothing
    
    'GetObjectContext.SetComplete
    
    UpdateNVHModosi = True
    
Exit Function
DefectErrorHandler:

    UpdateNVHModosi = False
    
    'エラー処理
    If Not objRs Is Nothing Then
        objRs.Close
            objRs = Nothing
    End If
    
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
      
    'If Not GetObjectContext Is Nothing Then
        'GetObjectContext.SetAbort
    End If
      
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
                  'Err.Number, Err.Source, Err.Description
End Function

'*< TOSTEM >*******************************************************************'
'*【プロシージャ名】
'*!     GetModosiCount
'*【処理概要】
'*?     戻し価格テーブルの存在チェック
'*【パラメータ】
'*      Name                Type            Comment
'*P     strNengetu;         <String>;       年月;
'*P     strJanCd;           <String>;       JANコード;
'*P     strTenCd;           <String>;       店コード;
'*【戻り値】
'*      Name                    Type                Comment
'*R     <GetModosiCount>;       <ADODB.Recordset>;  存在チェック
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)            EDP-No      説明
'*M 2009/08/06; 和田(IC);              P-36097;    新規作成;

'*< Comment End >**************************************************************
Private Function GetModosiCount(ByVal strNengetu As String, ByVal strJanCd As String, _
                           ByVal strTenCd As String) As ADODB.Recordset

    Dim strSQL                  As String
    Dim objCDBAccess            As COMDBUtil.CDBAccess  'DBオブジェクト
    Dim strMethodCallInfo       As String
    Dim strKikakuSyuruiCD       As String
    Dim strNenRenNO             As String
    Dim strEdaNO                As String
    Dim vntAryParam             As Object

    On Error GoTo DefectErrorHandler

    'strMethodCallInfo = App.Title & ".CHacTenCom.GetModosiCount(" & _
                        strNengetu & "," & strJanCd & "," & _
                        strTenCd & ")"
            
    strSQL = ""
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     nengetu"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     t_modosi_kakaku WITH (READCOMMITTED)"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     nengetu = ?"
    strSQL = strSQL & "     AND jan_cd = ?"
    strSQL = strSQL & "     AND ten_cd = ?"
    
    vntAryParam = Array(MakeParam("@Nengetu", DT_Char, 6, strNengetu) _
                        , MakeParam("@JanCd", DT_Char, 13, strJanCd) _
                        , MakeParam("@TenCd", DT_Char, 4, strTenCd) _
                        )
    
    'インスタンス作成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'Recordset取得
        GetModosiCount = objCDBAccess.GetDCRecordset("VSWF", strSQL, vntAryParam)
    'インスタンス破棄
    objCDBAccess = Nothing
    
Exit Function
DefectErrorHandler:

    'エラー処理
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
      
    'If Not GetObjectContext Is Nothing Then
        'GetObjectContext.SetAbort
    End If
      
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
                  'Err.Number, Err.Source, Err.Description
End Function

'*< TOSTEM >*******************************************************************'
'*【プロシージャ名】
'*!     GetJyogaiKbn
'*【処理概要】
'*?     商品選定DB(変更除外店舗)情報を取得する
'*【パラメータ】
'*      Name                Type            Comment
'*P     strKikakuNo;        <String>;       企画NO;
'*P     strTenCd;           <String>;       店コード;
'*【戻り値】
'*      Name                    Type                 Comment
'*R     <GetSiireJyogaiKbn>;    <ADODB.Recordset>;   仕入除外区分
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)            EDP-No      説明
'*M 2009/08/06; 和田(IC);              P-36097;    新規作成;
'*< Comment End >**************************************************************
Public Function GetJyogaiKbn(ByVal strKikakuNo As String, ByVal strTenCd As String) As ADODB.Recordset

    Dim objCDBAccess            As COMDBUtil.CDBAccess  'DBオブジェクト
    Dim strSQL                  As String   'SQL
    Dim strMethodCallInfo       As String
    Dim strKikakuSyuruiCD       As String
    Dim strNenRenNO             As String
    Dim strEdaNO                As String
    Dim vntAryParam             As Object
    
    On Error GoTo DefectErrorHandler
    
    'strMethodCallInfo = App.Title & ".CHacTenCom.GetJyogaiKbn(" & _
                        strKikakuNo & "," & strTenCd & "," & ")"
                                                
    '企画NOから企画種類コード、年連番、枝番を取得
    strKikakuSyuruiCD = Left(strKikakuNo, 3)
    strNenRenNO = Mid(strKikakuNo, 4, 10)
    strEdaNO = Right(strKikakuNo, 3)

    'インスタンス作成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
                
    '更新用SQL作成
    strSQL = ""
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     baika_jyogai_kbn"
    strSQL = strSQL & "     ,genka_jyogai_kbn"
    strSQL = strSQL & "     ,hattyuu_tani_jyogai_kbn"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     t_syousen_jyogai_tenpo WITH (READCOMMITTED)"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     kikaku_syurui_cd = ?"
    strSQL = strSQL & "     AND nen_ren_no = ?"
    strSQL = strSQL & "     AND eda_no = ?"
    strSQL = strSQL & "     AND ten_cd = ?"
    
    vntAryParam = Array(MakeParam("@KikakuSyuruiCd", DT_Char, 3, strKikakuSyuruiCD) _
                        , MakeParam("@NenRenNO", DT_Char, 10, strNenRenNO) _
                        , MakeParam("@EdaNO", DT_Char, 3, strEdaNO) _
                        , MakeParam("@TenCd", DT_Char, 4, strTenCd) _
                        )
    
    'Recordset取得
        GetJyogaiKbn = objCDBAccess.GetDCRecordset("VSWF", strSQL, vntAryParam)

    'インスタンス破棄
    objCDBAccess = Nothing
    
    'GetObjectContext.SetComplete
    
Exit Function
DefectErrorHandler:

    'エラー処理
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
      
    'If Not GetObjectContext Is Nothing Then
        'GetObjectContext.SetAbort
    End If
      
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
                  'Err.Number, Err.Source, Err.Description
End Function

'***** 2018/11/16 ログを細かく出したい為、改修 ***********************************************************************************
''*< TOSTEM >*******************************************************************'
''*【プロシージャ名】
''*!     SetLog
''*【処理概要】
''*?     一括変更登録の処理開始時と終了時にログを登録する
''*【パラメータ】
''*      Name                Type            Comment
''*P     strSyoriKbn;        <String>;       一括変更登録処理;
''*【戻り値】
''*      Name                    Type                 Comment
''*R     なし;
''******************************************************************************
''*【変更履歴】
''*  日付        名前(会社名)            EDP-No      説明
''*M 2015/09/03; 和田(IC);               P-51672;    新規作成;
''*< Comment End >**************************************************************
'Public Sub SetLog(ByVal lngRenban As Long, _
'                    ByVal strSyoriKbn As String, _
'                    ByVal strLoginUser As String)
'
'    Dim strMethodCallInfo As String
'    Dim vntAryParam As Object
'    Dim objCDBAccess    As COMDBUtil.CDBAccess
'    Dim strSQL          As String
'
'    On Error GoTo DefectErrorHandler
'
'    'エラー表示用
'    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.SetLog(" & strSyoriKbn & ")"
'
'    'CDBAccessインスタンス生成
'    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
'
'    'SQL文
'    strSQL = ""
'    strSQL = strSQL & " INSERT INTO"
'    strSQL = strSQL & "     t_ikkatu_henkou_touroku_log"
'    strSQL = strSQL & " SELECT ?,?,?,GETDATE()"
'
'    'MakeParamの作成
'    vntAryParam = Array(MakeParam("@renban", DT_Long, 4, lngRenban), _
'                        MakeParam("@syori_kbn", DT_String, 2, strSyoriKbn), _
'                        MakeParam("@user_id", DT_String, 40, strLoginUser) _
'                        )
'
'    'SQL実行
'    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
'
'    'インスタンス破棄
'    objCDBAccess = Nothing
'
'    'Commit
'    'GetObjectContext.SetComplete
'Exit Sub
'
'DefectErrorHandler:
'    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
'    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
'    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
'             'Err.Number, Err.Source, Err.Description
'End Sub

'*< TOSTEM >*******************************************************************'
'*【プロシージャ名】
'*!     SetLog
'*【処理概要】
'*?     一括変更登録処理内でログを登録・更新する
'*【パラメータ】
'*      Name                Type            Comment
'*P1    lngRenban;          <String>;       連番;
'*P2    strSyoriFile;       <String>;       処理ファイル(発注中止日、商品基本情報、分類情報、店別情報);
'*P3    lngDataCnt;         <Long>;         データ数;
'*P5    strSyoriMethod;     <String>;       処理メソッド;
'*P6    strSyoriSyousai;    <String>;       処理詳細;
'*P7    strLoginUser;       <String>;       ユーザーID;
'*P8    strUpdateKbn;       <String>;       データ更新区分(1：登録開始、2：処理開始、３：処理終了、４：登録終了);
'*【戻り値】
'*      Name                    Type                 Comment
'*R     なし;
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)            EDP-No      説明
'*M 2018/11/19; 山田(IC);               P-XXXXX;    新規作成;
'*< Comment End >**************************************************************
Public Sub SetLog(ByVal lngRenban As Long, _
                    ByVal strSyoriFile As String, _
                    ByVal lngDateCnt As Long, _
                    ByVal strSyoriMethod As String, _
                    ByVal strSyoriSyousai As String, _
                    ByVal strLoginUser As String, _
                    ByVal strUpdateKbn As String)

    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    Dim objCDBAccess    As COMDBUtil.CDBAccess
    Dim strSQL          As String

    On Error GoTo DefectErrorHandler

    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.SetLog(" & strSyoriFile & ")"
    
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL文
    strSQL = ""
    
    '更新区分ごとに実行するSQLを設定
    Select Case strUpdateKbn
        Case REGIST_STA
            strSQL = strSQL & " INSERT INTO"
            strSQL = strSQL & "     t_ikkatu_henkou_touroku_log"
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & "     '" & lngRenban & " '"
            strSQL = strSQL & "     ,'" & strSyoriFile & "' "
            strSQL = strSQL & "     ,'" & lngDateCnt & " ' "
            strSQL = strSQL & "     ,'" & strSyoriMethod & " ' "
            strSQL = strSQL & "     ,'" & strSyoriSyousai & " ' "
            strSQL = strSQL & "     ,'" & strLoginUser & " ' "
            strSQL = strSQL & "     ,GETDATE() "
            strSQL = strSQL & "     ,NULL "
            strSQL = strSQL & "     ,NULL "
            strSQL = strSQL & "     ,NULL "

            
'            'MakeParamの作成
'            vntAryParam = Array(MakeParam("@Renban", DT_Long, 4, lngRenban), _
'                                MakeParam("@syori_file", DT_String, 10, strSyoriFile), _
'                                MakeParam("@data_cnt", DT_Long, 4, lngDateCnt), _
'                                MakeParam("@syori_method", DT_String, 40, strSyoriMethod), _
'                                MakeParam("@syori_syousai", DT_String, 100, strSyoriSyousai), _
'                                MakeParam("@user_id", DT_String, 40, strLoginUser))
        
        Case PROCESS_STA
            strSQL = strSQL & " UPDATE "
            strSQL = strSQL & "     t_ikkatu_henkou_touroku_log"
            strSQL = strSQL & " SET "
            strSQL = strSQL & "     syori_file = '" & strSyoriFile & "'"
            strSQL = strSQL & "    ,data_cnt = '" & lngDateCnt & "'"
            strSQL = strSQL & "    ,syori_method = '" & strSyoriMethod & "'"
            strSQL = strSQL & "    ,syori_syousai = '" & strSyoriSyousai & "'"
            strSQL = strSQL & "    ,syori_kaishi = GETDATE() "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & "     renban = '" & lngRenban & "'"
            
        
        Case PROCESS_END
            strSQL = strSQL & " UPDATE "
            strSQL = strSQL & "     t_ikkatu_henkou_touroku_log"
            strSQL = strSQL & " SET "
            strSQL = strSQL & "    syori_method = '" & strSyoriMethod & "'"
            strSQL = strSQL & "    ,syori_syousai = '" & strSyoriSyousai & "'"
            strSQL = strSQL & "    ,syori_end = GETDATE() "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & "     renban = '" & lngRenban & "'"
            
            
        Case REGIST_END
            strSQL = strSQL & " UPDATE "
            strSQL = strSQL & "     t_ikkatu_henkou_touroku_log"
            strSQL = strSQL & " SET "
            strSQL = strSQL & "    syori_method = '" & strSyoriMethod & "'"
            strSQL = strSQL & "    ,syori_syousai = '" & strSyoriSyousai & "'"
            strSQL = strSQL & "    ,touroku_end = GETDATE() "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & "     renban = '" & lngRenban & "'"
            
        Case Else
    End Select
    
    'SQL実行
    'Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    'インスタンス破棄
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
             'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >*******************************************************************
'*
'*【プロシージャ名】
'*!     GetSaibanNo2
'*
'*【処理概要】
'*?     採番NOの取得。
'*
'*【パラメータ】
'*      Name                　　  Type                Comment
'*P     strSaibanKbn;             string              採番区分
'*
'*【戻り値】
'*      Name                　　  Type                Comment
'*R     GetSaibanNo;              integer;            採番NO;
'*
'******************************************************************************
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)       EDP-No      説明
'*M 2015/09/03; 和田(IC);          P-51672;    新規作成;
'*< Comment End >**************************************************************
Public Function GetSaibanNo2(ByVal strSaibanKbn As String) As Long
                               
    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    Dim objCDBAccess    As COMDBUtil.CDBAccess
    Dim strSQL          As String
    Dim rsSaiban        As ADODB.Recordset
    
    On Error GoTo DefectErrorHandler

    'エラー表示用
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.GetSaibanNo2(" & strSaibanKbn & ")"
    
    'CDBAccessインスタンス生成
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL文
    strSQL = ""
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     saiban_no + 1 AS saiban_no"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     w_syousen_neire_keikoku_saiban WITH(UPDLOCK)"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     saiban_kbn = ?"

    '切断Recordset取得
        rsSaiban = objCDBAccess.GetDCRecordset("VSWF", strSQL, MakeParam("@saiban_kbn", DT_Char, 1, strSaibanKbn))
    
    GetSaibanNo2 = rsSaiban.Fields("saiban_no").Value
    rsSaiban.Close
        rsSaiban = Nothing
    
     'SQL文
    strSQL = ""
    strSQL = strSQL & " UPDATE"
    strSQL = strSQL & "     w_syousen_neire_keikoku_saiban"
    strSQL = strSQL & " SET"
    strSQL = strSQL & "     saiban_no = ?"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     saiban_kbn = ?"
   
    'MakeParamの作成
    vntAryParam = Array(MakeParam("@saiban_no", DT_Long, 4, GetSaibanNo2), _
                        MakeParam("@saiban_kbn", DT_Char, 1, strSaibanKbn) _
                        )
    
    'SQL実行
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    'インスタンス破棄
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Function

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
             'Err.Number, Err.Source, Err.Description
End Function


'*< TOSTEM >*******************************************************************'
'*【プロシージャ名】
'*!     LockTableIkkatu
'*【処理概要】
'*?     一括変更登録処理内でログを登録・更新する
'*【パラメータ】
'*      Name                Type            Comment
'*【戻り値】
'*      Name                    Type                 Comment
'*R     なし;
'******************************************************************************
'*【変更履歴】
'*  日付        名前(会社名)            EDP-No      説明
'*M 2018/11/19; 李松涛(大連);               P-XXXXX;    新規作成;
'*< Comment End >**************************************************************
Public Sub LockTableIkkatu(ByRef objCDBAccess)
        '2018/12/14; 李松涛(大連) 一括登録時、毎回一つユーザーのみ実行できる 追加↓
        Dim strSQL
        strSQL = strSQL & "WHILE Object_id('tempdb..##LockTableIkkatu') IS NOT NULL "
        strSQL = strSQL & "BEGIN "
        strSQL = strSQL & "    IF Datediff(second, (SELECT ins_date FROM ##LockTableIkkatu), Getdate()) > 60 "
        strSQL = strSQL & "    BEGIN "
        strSQL = strSQL & "        DROP TABLE ##LockTableIkkatu "
        strSQL = strSQL & ""
        strSQL = strSQL & "        BREAK "
        strSQL = strSQL & "    END "
        strSQL = strSQL & "    WAITFOR delay '00:00:01.000' "
        strSQL = strSQL & "END "
        strSQL = strSQL & ""
        strSQL = strSQL & "IF Object_id('tempdb..##LockTableIkkatu') IS NULL "
        strSQL = strSQL & "BEGIN"
        strSQL = strSQL & "     SELECT Getdate() AS ins_date INTO   ##LockTableIkkatu "
        strSQL = strSQL & "END"
        'SQL実行
        Call objCDBAccess.ExecuteSql("VSWF", strSQL)
        Dim mydt
        mydt = Now()
        pStrRenban = Year(mydt) & Month(mydt) & Day(mydt) & "_" & Hour(mydt) & Minute(mydt) & Second(mydt) & "." & Right("000" & ((Timer() * 1000) Mod 1000), 3)
        
        '2018/12/14; 李松涛(大連) 一括登録時、毎回一つユーザーのみ実行できる 追加↑
End Sub


Public Sub SetIkkatuLog(ByVal syori_step As String, _
                            ByVal syori_mei As String, _
                            ByVal syori_method As String, _
                            ByVal syori_syousai As String, _
                            ByVal syori_kekka As String, _
                            ByVal user_id As String, _
                            ByVal kbn As String)
                            

    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    'Dim objCDBAccess    As COMDBUtil.CDBAccess
    Dim strSQL          As String

    On Error GoTo DefectErrorHandler

    'エラー表示用
    ''strMethodCallInfo = "VswfIkatuTou.CHacTenCom.SetIkkatuLog()"
    
    'CDBAccessインスタンス生成
    'objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL文
    strSQL = ""
    
    syori_step = Right("                    " & syori_step, 20)
    

    Dim mydt
    
    If pStrRenban = "" Then
        mydt = Now()
        pStrRenban = Year(mydt) & Month(mydt) & Day(mydt) & "_" & Hour(mydt) & Minute(mydt) & Second(mydt) & "." & Right("000" & ((Timer() * 1000) Mod 1000), 3)
    End If
    
    '更新区分ごとに実行するSQLを設定
    Select Case kbn
        Case "INS"
            strSQL = strSQL & " INSERT INTO"
            strSQL = strSQL & "     t_ikkatu_henkou_touroku_log_moto"
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & "     '" & pStrRenban & " '"
            strSQL = strSQL & "     ,'" & syori_step & "' "
            strSQL = strSQL & "     ,GETDATE() "
            strSQL = strSQL & "     ,null "
            strSQL = strSQL & "     ,'" & syori_mei & " ' "
            strSQL = strSQL & "     ,'" & syori_method & " ' "
                        strSQL = strSQL & "     ,'" & syori_syousai & " ' "
                        strSQL = strSQL & "     ,'' "
                        strSQL = strSQL & "     ,'" & user_id & " ' "


        Case "UPD"
            strSQL = strSQL & " UPDATE "
            strSQL = strSQL & "     t_ikkatu_henkou_touroku_log_moto"
            strSQL = strSQL & " SET "
            strSQL = strSQL & "     end_time = GETDATE()"
            strSQL = strSQL & "    ,syori_kekka = '" & syori_kekka & "'"
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & "     renban = '" & pStrRenban & "'"
            strSQL = strSQL & "AND  syori_step = '" & syori_step & "'"
        Case Else
        
    End Select
    
    Dim adoConnectionNew
        adoConnectionNew = New ADODB.Connection
    adoConnectionNew.ConnectionString = "Provider=SQLOLEDB.1;Password=lixil@2014;Persist Security Info=True;User ID=sa;Initial Catalog=VSWF;Data Source=ot2414;"
    adoConnectionNew.Open
    Dim rtv
        rtv = adoConnectionNew.Execute(strSQL)
    adoConnectionNew.Close
        adoConnectionNew = Nothing
    'SQL実行
    'Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    'インスタンス破棄
    'objCDBAccess = Nothing
    
    'Commit
    ''GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    'If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    ''If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "予期しないエラーです。", _
             'Err.Number, Err.Source, Err.Description
End Sub


    Private Function IsNull(ByVal v As Object) As Boolean
        Return IsDBNull.vaue = True
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub
End Class