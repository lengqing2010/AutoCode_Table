Option Explicit Off

Imports Microsoft.VisualBasic
Imports System.IO

Partial Class _Default
    Inherits System.Web.UI.Page




'******************************************************************'
'*�y���W���[�����z
'�@    CHacTenCom.cls

'*�y�T�v�z
'�@    �������~���ꊇ�o�^or�X�ʏ��ꊇ�ύX�o�^

'*�y�ݒ�z
'�@�@�@@TransactionModeProperty
'           RequiresTransaction
'�@�@�@@InstancingProperty
'           MultiUse
'�@�@�@@ReferenceSettings
'           COM+ Services Type Library
'           Microsoft ActiveX Data Object 2.6 Library
'           COMDBUtil(DBAccess���ʃR���|�[�l���g)
'           Microsoft Scripting Runtime

'******************************************************************'

'CSV�t�@�C����ԍ� (���̕��т�CSV�̕��тɂȂ�܂��B)
Private Enum enmHacColNo    '�������~���pCSV�f�[�^����
    jigCD = 0               '�Y�����ƃR�[�h
    SyuryouDate             '�����I����
    janCD                   'JAN�R�[�h
    Syobun                  '�������@
    ChushiKbn               '�������~���R�敪
    DaiJanCd                '��֏��iJAN�R�[�h
End Enum

Private Enum enmTenColNo    '�X�ʗpCSV�f�[�^����
    TenCd = 0               '�Y���X�R�[�h
    HenkouDate              '�ύX��
    janCD                   'JAN�R�[�h
    Genka                   '����
    Baika                   '����
    BaikaRiyuu              '�����ύX���R
    Torihikisaki            '�����R�[�h
    Hacchuutani             '�����P��
    CenterKbn               '�Z���^�[�敪
End Enum
'2009.06.06 Add by Supattra T.(Tostem Thai)Start.
Private Enum enmModColNo
    TenCd = 0               '�Y���X�R�[�h
    HenkouDate              '�ύX��
    janCD                   'JAN�R�[�h
    Genka                   '����
    Baika                   '����
    BaikaRiyuu              '�����ύX���R
    Torihikisaki            '�����R�[�h
    Hacchuutani             '�����P��
    CenterKbn               '�Z���^�[�敪
    centerCD                '�Z���^�[�R�[�h
    MoGenka                 '�߂�����
    MoBaika                 '�߂�����
    HattyuuTani             '�߂������P��
    PopSize                 'POP�T�C�Y
End Enum
'2009.06.06 Add by Supattra T.(Tostem Thai)End.

'********** 2005/08/31 ���� �ǉ� **********��
'CSV�t�@�C����ԍ� (���̕��т�CSV�̕��тɂȂ�܂��B)
Private Enum enmSyoColNo    '���i��{�pCSV�f�[�^����
    janCD = 0               'JAN�R�[�h
    SyouhinmeiKana          '�J�i���i��(TA�`�[�p)
    SyouhinmeiKanji         '�������i��(SA/��ײ����ޗp)
    ReceiptmeiKana          '���V�[�g��(��)
    ReceiptmeiKanji         '���V�[�g��(����)
    SyouhihinmeiHansoku     '���i��(�̑��p)
    
    MakerMei                '���[�J�[��
    KikakuSize              '�K�i/�T�C�Y
    HinbanKataban           '�i��/�^��
    IroGara                 '�F/��
    Taste                   '�e�[�X�g
    Seisantimei             '���Y�n��
    Tokutyou                '����
    SyouhinCopy             '���i�R�s�[
    TuujyouKakaku           '�ʏ퉿�i
    MakerKakaku             '���[�J�[��]�������i
    POPLayout1              'POP���C�A�E�g�P
    POPLayout2              'POP���C�A�E�g�Q
    PopSize                 'POP�T�C�Y
    
    AtukaiPtn               '�����p�^�[��
    HyoujyunBaika           '�W������
    SikiriGenka             '�d�،���
    VhJyuyouCd              'VH���v�R�[�h
    RdJyuyouCd              'RD���v�R�[�h
    VcsJyuyouCd             'VCS���v�R�[�h
    SvhJyuyouCd             'SVH���v�R�[�h
End Enum

Private Enum enmBunColNo    '���ޗpCSV�f�[�^����
    HenkouDate = 0          '�ύX��
    janCD                   'JAN�R�[�h
    Bunrui1                 '��1���ރR�[�h
    Bunrui2                 '��2���ރR�[�h
    Bunrui3                 '��3���ރR�[�h
End Enum
'********** 2005/08/31 ���� �ǉ� **********��

'CSV�t�@�C�� �t�B�[���h��
Private Const FLD_HACCHUU = "�Y�����ƃR�[�h,�����I����,JAN�R�[�h,�������@,�������~���R�敪,��֏��i�R�[�h"
'*M 2014/01/22; ������(��A) ;�@P-46854; �Ő������ɔ����Ή�(���i�I��WF�E�V���i���);�@START
Private Const FLD_TENBETU = "�Y���X�R�[�h,�ύX��,JAN�R�[�h,����,����,�����ύX���R,�����,�����P��,�Z���^�[�敪,�Z���^�[�R�[�h,�߂�����,�߂�����,�߂������P��,POP�T�C�Y"
'*M 2014/01/22; ������(��A) ;�@P-46854; �Ő������ɔ����Ή�(���i�I��WF�E�V���i���);�@END
Private Const FLD_SYOUHIN = "JAN�R�[�h,�J�i���i��(TA�`�[�p),�������i��(SA/��ײ����ޗp),���V�[�g��(��),���V�[�g��(����),���i��(�̑��p),���[�J�[��,�K�i/�T�C�Y,�i��/�^��,�F/��,�e�[�X�g,���Y�n��,����,���i�R�s�[,�ʏ퉿�i,Ұ����]�������i,POP���C�A�E�g�P,POP���C�A�E�g�Q,POP�T�C�Y,�����p�^�[��,�W������,�d�،���,VH���v�R�[�h,R&D���v�R�[�h,VCS���v�R�[�h,SVH���v�R�[�h"
Private Const FLD_BUNRUI = "�ύX��,JAN�R�[�h,��1���ރR�[�h,��2���ރR�[�h,��3���ރR�[�h"
'2009.06.05 Add by Supattra T.(Tostem Thai)Start.
'Private Const FLD_MODOSI = "�Y���X�R�[�h,�ύX��,JAN�R�[�h,����,����,�����ύX���R,�����,�����P��,�Z���^�[�敪,�߂�����,�߂�����,�߂������P��,POP�T�C�Y"
'*M 2014/01/22; ������(��A) ;�@P-46854; �Ő������ɔ����Ή�(���i�I��WF�E�V���i���);�@START
'Private Const FLD_MODOSI = "�Y���X�R�[�h,�ύX��,JAN�R�[�h,����,����,�����ύX���R,�����,�����P��,�Z���^�[�敪,�Z���^�[�R�[�h,�߂�����,�߂�����,�߂������P��,POP�T�C�Y"
Private Const FLD_MODOSI = "�Y���X�R�[�h,�ύX��,JAN�R�[�h,����,����(�ŕ�),�����ύX���R,�����,�����P��,�Z���^�[�敪,�Z���^�[�R�[�h,�߂�����,�߂�����,�߂������P��,POP�T�C�Y"
'*M 2014/01/22; ������(��A) ;�@P-46854; �Ő������ɔ����Ή�(���i�I��WF�E�V���i���);�@END
'2009.06.05 Add by Supattra T.(Tostem Thai)End.

'�X�E���ƃR�[�h�i�S�X,�S����,�{���j
Private Const ALL_TEN_CD = "9999"                 '�S�X
Private Const ALL_JIG_CD = "99"                   '�S����
Private Const HON_TEN_CD = "0000"                 '�{��

'�Ăяo�����֖߂��ׂ̖߂�l
Private Const RET_FLD_HACH As Long = 1                  '�������~��
Private Const RET_FLD_TENB As Long = 2                  '�X�ʏ��
Private Const RET_FLD_SYO As Long = 3                   '���i��{���
Private Const RET_FLD_BUN As Long = 4                   '���ޏ��
Private Const RET_NO_DATA As Long = 0                   '�f�[�^����
'2009.06.05 Add by Supattra T.(Tostem Thai)Start.
Private Const RET_FLD_MODO As Long = 5
'2009.06.05 Add by Supattra T.(Tostem Thai)End.

'���ƕʎ��v�R�[�h
Private Const JYU_VH As String = "01"                   'VH���v�R�[�h
Private Const JYU_RD As String = "02"                   'RD���v�R�[�h
Private Const JYU_VCS As String = "03"                  'VCS���v�R�[�h
Private Const JYU_SVH As String = "06"                  'SVH���v�R�[�h


'KET�敪
Private Const KEYCD_BAIKARIYUU = "025"            '�����ύX���R

'�Z���^�[�R�[�h
'Private Const CENTERCD_KODAMA = "0151"            '(���ʃZ���^�[)
Private Const CENTERCD_KODAMA = "0810"            '(�吶�����~�Z���^�[)

'���i�I��WF�̃e�[�u����
Private Const TBL_BAI_HEN As String = "t_syousen_baika_henkou"      '���i�I��DB(�����ύX���)
Private Const TBL_SII_HEN As String = "t_syousen_siire_henkou"      '���i�I��DB(�d���ύX���)
Private Const TBL_CEN_HEN As String = "t_syousen_center_cd_henkou"  '���i�I��DB(�Z���^�[�ύX���)
'********** 2008/05/27 ������ �ǉ� **********��
Private Const TBL_JYO_TEN As String = "t_syousen_jyogai_tenpo"      '���i�I��WF(�ύX���O�X��)
'********** 2008/05/27 ������ �ǉ� **********��

' �e�t�B�[���h�̌���
Private Const LENGTH_JANCODE1 As Long = 8                   'JAN�R�[�h�����P
Private Const LENGTH_JANCODE2 As Long = 13                  'JAN�R�[�h�����Q
Private Const LENGTH_TENCODE As Long = 4                    '�X�R�[�h�̌���
Private Const LENGTH_CENTERCODE As Long = 4                 '�Z���^�[�R�[�h�̌���
Private Const LENGTH_BAIKA As Long = 7                      '�����̌���
Private Const LENGTH_HATTYU_TANI As Long = 5                '�����P�ʌ���
Private Const LENGTH_GENKA_SEISU As Long = 7                '����(������)����
Private Const LENGTH_GENKA_SYOUSUU As Long = 2              '����(������)����
Private Const LENGTH_HENKOU_DATE As Long = 8                '�ύX���̌���
Private Const LENGTH_BUNRUI As Long = 6                     '���ރR�[�h�̌���
Private Const LENGTH_JYUYOU As Long = 6                     '���v�R�[�h�̌���
Private Const LENGTH_SYOKANA As Long = 40                   '�J�i���i��(TA�`�[�p)�̌���
Private Const LENGTH_SYOKANJI As Long = 30                  '�������i��(SA/��ײ����ޗp)�̌���
Private Const LENGTH_RESKANA As Long = 14                   '���V�[�g��(��)�̌���
Private Const LENGTH_RESKANJI As Long = 14                  '���V�[�g��(����)�̌���
Private Const LENGTH_SYOHANSOKU As Long = 20                '���i��(�̑��p)�̌���
Private Const LENGTH_ATUPTN As Long = 3                     '�����p�^�[���̌���

'�G���[�̌����������萔
Private Const PARAM_EMPTY = 10              ' �p�����[�^�����͂���Ă��Ȃ�
Private Const PARAM_NOTDISIT = 20           ' �p�����[�^���񐔒l�ł���
Private Const PARAM_NOTFINDMASTER = 30      ' �}�X�^�ɑ��݂��Ȃ��f�[�^
Private Const PARAM_WRONGCHECKDIGIT = 40    ' �`�F�b�N�f�B�W�b�g�s��
Private Const PARAM_WRONGDATE = 50          ' ���t�s��
Private Const PARAM_SHORTDISIT = 60         ' �����s��
Private Const PARAM_NOTEXISTDB = 70         ' �f�[�^�x�[�X(�}�X�^)�ɑ��݂��Ȃ�
Private Const PARAM_SAMECODEDB = 80         ' ���łɃf�[�^�x�[�X��ɓo�^����Ă���
Private Const PARAM_TORIHIKISTOP = 90       ' �����Ƃ̎���͒�~����Ă���
Private Const PARAM_REPDATEBAIKA = 100      ' ����DB(�����ύX���)�Ɋ��ɓ����ύX���̃f�[�^�����݂���
Private Const PARAM_REPDATESIIRE = 110      ' ����DB(�d���ύX���)�Ɋ��ɓ����ύX���̃f�[�^�����݂���
Private Const PARAM_REPDATETENPO = 120      ' ����DB(�X�܃Z���^�[�R�[�h�ύX���)�Ɋ��ɓ����ύX���̃f�[�^�����݂���
'********** 2005/04/25 ��� �ǉ� **********��
Private Const PARAM_NOTDEFINED = 130        ' ���̒l�̓p�����[�^�̒�`����O��Ă���
Private Const PARAM_EARLYCANCELDAY = 140    ' �����I��������������i�J�n�����O)
'********** 2005/04/25 ��� �ǉ� **********��
'********** 2005/05/07 �k�V **********��
Private Const PARAM_ZERO = 150              ' ���z�ނ����̒l�Ŗ���
Private Const PARAM_MISENTAKU = 160         ' �������~���R��00(���I��)
'********** 2005/05/07 �k�V **********��
'********** 2005/06/16 �k�V **********��
Private Const PARAM_BAIKARIYUU_HISSU = 170  ' �X�R�[�h=0000�̏ꍇ�A���������͂���Ă�����A�����ύX���R�͓��͕K�{
Private Const PARAM_BAIKARIYUU_FUKA = 180   ' �����������͂̏ꍇ�A���͓X�R�[�h<>0000�̏ꍇ�A�����ύX���R�͓��͕s��
Private Const PARAM_NONARROW = 190          ' �p�����[�^�[���񔼊p�����ł���
Private Const PARAM_BUNRUIDB = 200          ' ���ޏ�񂪑��݂��Ȃ�
Private Const PARAM_URIBADB = 210           ' ����̌n��񂪑��݂��Ȃ�
Private Const PARAM_ZENKAKU = 220           ' �S�p������ȊO�œ��͂���Ă���
Private Const PARAM_SAMEKEYDB = 230         ' ���O�̓����R�[�h�Əd�����Ă���
Private Const PARAM_HANEISUU = 240          ' ���p�p�����ȊO�̕��������͂���Ă���B
Private Const PARAM_KINSIMOJI = 250         ' ���p�̋֎~���������͂���Ă���B
'2014/05/02 add wadak7
Private Const PARAM_KINSIMOJI_RECEIPT = 255 ' �֎~���������͂���Ă���(���V�[�g��(����))�B
'********** 2005/06/16 �k�V **********��
Private Const PARAM_GROUP = 260             ' �O���[�v�i�����͂���Ă���B
Private Const PARAM_KINSI = 270             ' ���͂ł��Ȃ��B
'********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
Private Const PARAM_HONBU_NAI = 280         ' �Z���^�[�敪��2�FPDC�̏ꍇ�A0000�F�{���̃f�[�^�͕K�{���͂ł�
Private Const PARAM_HATTYUU_TANI = 290      ' �Z���^�[�敪��2�FPDC�̔����P�ʂ͑S�ē����l�ł͂Ȃ�
Private Const PARAM_GENKA = 300             ' �Z���^�[�敪��2�FPDC�̌����͑S�ē����l�ł͂Ȃ�
Private Const PARAM_GENKA_HATTYUUTANI = 310 ' �X�V�X�܈ȊO��PDC�敪�ɂȂ��Ă���X�܂�����܂��B
'********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
'�G���[��\�����郂�W���[���ϐ�
    Private Structure typ_Error_Structure
        Dim intErrorLine As Integer
        Dim intErrorField As Integer
        Dim intErrorType As Integer
        Dim strErrorParameter As String
    End Structure

' �G���[�\���̔z��A�������ɃG���[��񂪊i�[�����
Private typErrorArray() As typ_Error_Structure
'********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
Private typErrorArrayPDC() As typ_Error_Structure
'********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��

' ���������G���[�̐�
Private intErrorCount As Integer
'********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
Private intErrorCountPDC As Integer
'********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��

'���W���[���ϐ�
Private pstrTantousyaId As String   '�S����ID
Private pstrFilePath    As String   '�t�@�C���p�X
Private pstrErrMsg As String        '�G���[���b�Z�[�W
Private pintSyouhizei As Integer    '����ŗ�(%)

'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************
Private plngSaibanNo2 As Long       '�̔ԇ�
Private plngDataCnt As Long         '�f�[�^��
Private pstrFileSyurui As String    '�t�@�C�����
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************

Private pstrKikakuSyurui As String      '�����
Private pstrKikakuNenrenban As String   '���N�A��
Private pstrKikakuEdaban As String      '���}��

Private pflgBaikaToukbn As Boolean   '�����ύX���o�^�敪�t���O�i1:�ǉ��̐�t��񂠂�A0:�Ȃ��j
Private pflgSiireToukbn As Boolean   '�d���ύX���o�^�敪�t���O�i1:�ǉ��̐�t��񂠂�A0:�Ȃ��j
Private pflgCenterToukbn As Boolean  '�Z���^�[�R�[�h�ύX���o�^�敪�t���O�i1:�ǉ��̐�t��񂠂�A0:�Ȃ��j
'********** 2005/04/25 ���� �ǉ� **********��
Private pflgJigJyoToukbn As Boolean  '���ƕʏ��o�^�敪�t���O�i1:�ǉ��̐�t��񂠂�A0:�Ȃ��j
'********** 2005/04/25 ���� �ǉ� **********��

'********** 2005/05/11 �y�c �ǉ� **********��
'DB�Q�ƃG���[�`�F�b�N�p�̐ؒf���R�[�h�Z�b�g

'�X�ʁA�������~�����g�p
Private pChkSameJANCodeRS As Object     '�d��JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g
'********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
Private pChkSameGenkaHattyuuTaniRS As Object  '�����P�ʁA��������v�`�F�b�N�p�ؒf���R�[�h�Z�b�g
Private pChkHonbuGenkaHattyuuTaniRS As Object  '�����P�ʁA��������v�`�F�b�N�p�ؒf���R�[�h�Z�b�g�i�{���j
'********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
Private pChkvk1JANCodeRS As Object      '�����}�X�^�o�^�ς�JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g

'���ޏ��g�p
Private pChkBunruiCodeRS As Object      '���ރ}�X�^KEY�`�F�b�N�p�ؒf���R�[�h�Z�b�g
Private pCsvBunrui1CodeList As String   'CSV�t�@�C������1���ރR�[�h�̃��X�g(��1���ރR�[�h�̏d������)

'���v���g�p
Private pChkJyuyouCodeRS As Object      '����̌n�\���p�}�X�^���v�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g
Private pCsvJyuyouCodeList As String    'CSV�t�@�C�������v�R�[�h�̃��X�g(��1���ރR�[�h�̏d������)

'�����p�^�[�����g�p
Private pChkAtukaiPtnCdRS As Object      '�����p�^�[���R�[�h�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g

'���v�Ǘ����g�p
'Private pChktsrkJANCodeRS As Object
'WF�Ǘ����g�p
'Private pChktswkJANCodeRS As Object

'********** 2005/05/18 �y�c �������~�ł��g�p����ׁA�ϐ����ύX **********��
Private pCsvJanCodeList As String      'CSV�t�@�C����Jan�R�[�h�̃��X�g(Jan�R�[�h�̏d������)
'********** 2005/05/18 �y�c �ϐ����ύX **********��
'2009.06.08 Add by Supattra T.(Tostem Thai)Start.
Private pCsvHenkouDate As String
Private pCsvhenkouDateKeizoku As String
Private pTTenCode As String
Private pTJanCode As String
'2009.06.08 Add by Supattra T.(Tostem Thai)End.

'�X�ʗp
Private pTChkJigyouRS As Object         '�Y�����ƃR�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g
'2009.07.10 Add by Supattra T.(Tostem Thai)Start.
Private pTChkkikakuRS As Object
Private pTChkCenterCdRS As Object       '�Z���^�[�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g

'2009.07.10 Add by Supattra T.(Tostem Thai)End
Private pTChkBaikaHenkouRS As Object    '�����ύX���R�`�F�b�N�p�ؒf���R�[�h�Z�b�g
Private pTChkTorihikisakiRS As Object   '�����`�F�b�N�p�ؒf���R�[�h�Z�b�g
Private pTChkCenterRS As Object         '�Z���^�[�敪�`�F�b�N�p�ؒf���R�[�h�Z�b�g

Private pTTenCodeList As String         '�W�J�X�R�[�h�̃��X�g(�S���ƁA�S�X�j

'********** 2005/05/19 �y�c �ǉ� **********��
Private pTKeyCDList As String      '�����}�X�^�o�^�ς�JAN�R�[�h�R�t����L�[�̃��X�g(��L�[�d������)
'********** 2005/05/19 �y�c �ǉ� **********��

'********** 2005/05/18 �y�c �ǉ� **********��
Private pstrMinDate As String      'CSV�t�@�C�����ŏ����t
Private pstrMaxDate As String      'CSV�t�@�C�����ő���t
'********** 2005/05/18 �y�c �ǉ� **********��

'�������~�p
Private pHChkJigyouRS As Object         '�Y�����ƃR�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g
Private pHChkHacchuRS As Object         '�������~���R�敪�`�F�b�N�p�ؒf���R�[�h�Z�b�g

'********** 2005/05/18 �y�c �ǉ� **********��
Private pHChkHattyuuKaisiDateRS As Object   '�����J�n���`�F�b�N�p�ؒf���R�[�h�Z�b�g
Private pHChkJHTYKaisiDateRS As Object      '�����J�n�������ƕʏ��e�[�u������擾�p�ؒf���R�[�h�Z�b�g

Private pHKeyCDList As String      '�����}�X�^�o�^�ς�JAN�R�[�h�R�t����L�[�̃��X�g(��L�[�d������)
'********** 2005/05/18 �y�c �ǉ� **********��

'********** 2005/05/11 �y�c �ǉ� **********��

'********** 2005/05/13 �c�� �ǉ� **********��
'DB�Q�ƃG���[�`�F�b�N�p�̐ؒf���R�[�h�Z�b�g�i�ǉ��j
Private pTChkMstBaikaRS As Object       '�����}�X�^�E�����ύX���̕ύX���`�F�b�N�p�ؒf���R�[�h�Z�b�g
Private pTChkMstSiireRS As Object       '�����}�X�^�E�d���ύX���̕ύX���`�F�b�N�p�ؒf���R�[�h�Z�b�g
Private pTChkMstCenterRS As Object       '�Z���^�[�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g
'********** 2005/05/13 �c�� �ǉ� **********��

    '********** 2005/06/15 �k�V **********��
    Dim arrTourokuKbn2(0, 0) As String      '�o�^�敪=2�̓X�R�[�h�̕ۊǗp�z��
    Private Const conColBaika As Integer = 0 '��L�z���2�����ڂ̗�ԍ�(�����ύX���)
    Private Const conColSiire As Integer = 1 '��L�z���2�����ڂ̗�ԍ�(�d���ύX���)
    Private Const conColCenter As Integer = 2 '��L�z���2�����ڂ̗�ԍ�(�Z���^�[�ύX���)
'********** 2005/06/15 �k�V **********��


'�萔
Private Const conLngStartRow As Integer = 2  'CSV�t�@�C����1�s�ڂ̃f�[�^�ʒu

'CSV�f�[�^�i�[�z�� (����)(���ڔԍ�)
Private paryCsvData() As Object

'********** 2018/12/17 ������ **********��
    Private pStrRenban As String


    'Property Let letFilePath(strPath As String)                     '�t�@�C���p�X
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


    'Property Let letKikakuSyurui(strKikakuSyurui As String)         '�����
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


    'Property Let letKikakuNenrenban(strKikakuNenrenban As String)   '���N�A�� (YYYY + �A��)
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


    'Property Let letKikakuEdaban(strKikakuEdaban As String)         '���}��
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


    'Property Let letTantousyaId(strTantousyaId As String)           '�S����ID
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
    'Property Let letSaibanNo2(lngSaibanNo2 As Long)                 '���O�L�ڗp_�̔ԇ�  2018/11/19 �R�c �\�����ʃG���[�Ή�
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

'*�y�v���V�[�W�����z
'*!    FncCSVIkkatsuTouroku

'*�y�����T�v�z
'*?    CSV�t�@�C���̃f�[�^�����i�I��WF��DB�Ɏ�荞��

'*�y�p�����[�^�z
'*P    strErrMsg;        string; �G���[���b�Z�[�W;

'*�y�߂�l�z
'*R    int; RET_FLD_HACH:   �������~���p�f�[�^;
'*R         RET_FLD_TENB:   �X�ʏ��p�f�[�^;
'*R         RET_NO_DATA:    �f�[�^����;
'*R         ��L�ȊO:        �G���[;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/18; ����(���j�V�X�e���������); �V�K�쐬;
'*M    2005/08/31; ����(IC)                 ; ���ޏ��ꊇ�ύX�ǉ�;
'*M    2009/06/06; Supattra T.(Tostem Thai) ; Add case CSV�t�@�C��;
'*< Comment End >**************************************************'
    Public Function FncCSVIkkatsuTouroku(ByRef strErrMsg _
                                        , ByRef aryCSVData As String()) As Integer
        Dim intRecNum
        Dim intRet

        plngDataCnt = 0                 '2018/11/19 �R�c �\�����ʃG���[�Ή�
        pstrFileSyurui = ""             '2018/11/19 �R�c �\�����ʃG���[�Ή�
        intRecNum = 0
        FncCSVIkkatsuTouroku = -1

        '�� �ꊇ�o�^���O�ǉ� 2018/12/17
        Call SetIkkatuLog("1", "CSV�t�@�C���̃f�[�^�����i�I��WF��DB�Ɏ�荞��", "FncCSVIkkatsuTouroku", "", "", pstrTantousyaId, "INS")

        '*******************************************************************************************************************************
        '���O�o�́i�@�A�ԁA�A�����t�@�C���A�B�f�[�^���A�C�������\�b�h�A�D�����ڍׁA�E���[�U�[ID�A�F�f�[�^�X�V�敪�j
        Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "FncCSVToArray_�����J�n�O", pstrTantousyaId, "1")
        '*******************************************************************************************************************************

        'CSV�t�@�C���̃f�[�^���擾���āA�z��ɓ����B
        '�����ŁA�������~��or�X�ʏ��ɕ�����B
        intRet = FncCSVToArray(strErrMsg)

        'CSV�f�[�^��ASP�ɕԂ�
        aryCSVData = paryCsvData

        If intRet = RET_FLD_HACH Then       '�������~���ꊇ�o�^�̏ꍇ
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
            pstrFileSyurui = "�������~��"
            '���O�o�́i�@�A�ԁA�A�����t�@�C���A�B�f�[�^���A�C�������\�b�h�A�D�����ڍׁA�E���[�U�[ID�A�F�f�[�^�X�V�敪�j
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "�������~���ꊇ�o�^����_�����J�n�O", pstrTantousyaId, "2")
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

            '********** 2005/04/26 ���� �ύX **********��
            '�������~���ꊇ�o�^�����̒ǉ�
            If FncHacTouroku(strErrMsg) = True Then
                FncCSVIkkatsuTouroku = RET_FLD_HACH
            Else
                FncCSVIkkatsuTouroku = -1
            End If
            '********** 2005/04/26 ���� �ύX **********��

            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
            '���O�o�́i�@�A�ԁA�A�����t�@�C���A�B�f�[�^���A�C�������\�b�h�A�D�����ڍׁA�E���[�U�[ID�A�F�f�[�^�X�V�敪�j
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "�������~���ꊇ�o�^����_�����I��", pstrTantousyaId, "3")
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

        ElseIf intRet = RET_FLD_TENB Then   '�X�ʏ��ꊇ�ύX�o�^�̏ꍇ
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
            pstrFileSyurui = "�X�ʏ��"
            '���O�o�́i�@�A�ԁA�A�����t�@�C���A�B�f�[�^���A�C�������\�b�h�A�D�����ڍׁA�E���[�U�[ID�A�F�f�[�^�X�V�敪�j
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "�X�ʏ��ꊇ�o�^����_�����J�n�O", pstrTantousyaId, "2")
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

            If FncTenTouroku(strErrMsg) = True Then
                FncCSVIkkatsuTouroku = RET_FLD_TENB
            Else
                FncCSVIkkatsuTouroku = -1
            End If

            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
            '���O�o�́i�@�A�ԁA�A�����t�@�C���A�B�f�[�^���A�C�������\�b�h�A�D�����ڍׁA�E���[�U�[ID�A�F�f�[�^�X�V�敪�j
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "�X�ʏ��ꊇ�o�^����_�����I��", pstrTantousyaId, "3")
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

            '2009.06.06 Add by Supattra T.(Tostem Thai)Start.
        ElseIf intRet = RET_FLD_MODO Then   'Add �X�ʏ��ꊇ�ύX�o�^�[
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
            pstrFileSyurui = "�߂����i"
            '���O�o�́i�@�A�ԁA�A�����t�@�C���A�B�f�[�^���A�C�������\�b�h�A�D�����ڍׁA�E���[�U�[ID�A�F�f�[�^�X�V�敪�j
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "�߂����i�ꊇ�o�^����_�����J�n�O", pstrTantousyaId, "2")
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

            If FncModosiKakaku(strErrMsg) = True Then
                FncCSVIkkatsuTouroku = RET_FLD_MODO
                aryCSVData = paryCsvData
            Else
                FncCSVIkkatsuTouroku = -1
            End If

            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
            '���O�o�́i�@�A�ԁA�A�����t�@�C���A�B�f�[�^���A�C�������\�b�h�A�D�����ڍׁA�E���[�U�[ID�A�F�f�[�^�X�V�敪�j
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "�߂����i�ꊇ�o�^����_�����I��", pstrTantousyaId, "3")
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
            '2009.06.06 Add by Supattra T.(Tostem Thai)End.
            '********** 2005/08/31 ���� �ǉ� **********��
        ElseIf intRet = RET_FLD_SYO Then   '���i��{���ꊇ�ύX�o�^�̏ꍇ
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
            pstrFileSyurui = "���i��{���"
            '���O�o�́i�@�A�ԁA�A�����t�@�C���A�B�f�[�^���A�C�������\�b�h�A�D�����ڍׁA�E���[�U�[ID�A�F�f�[�^�X�V�敪�j
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "���i��{���ꊇ�o�^����_�����J�n�O", pstrTantousyaId, "2")
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

            If FncSyoTouroku(strErrMsg) = True Then
                FncCSVIkkatsuTouroku = RET_FLD_SYO
            Else
                FncCSVIkkatsuTouroku = -1
            End If

            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
            '���O�o�́i�@�A�ԁA�A�����t�@�C���A�B�f�[�^���A�C�������\�b�h�A�D�����ڍׁA�E���[�U�[ID�A�F�f�[�^�X�V�敪�j
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "���i��{���ꊇ�o�^����_�����I��", pstrTantousyaId, "3")
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

        ElseIf intRet = RET_FLD_BUN Then   '���ޏ��ꊇ�ύX�o�^�̏ꍇ
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
            pstrFileSyurui = "���ޏ��"
            '���O�o�́i�@�A�ԁA�A�����t�@�C���A�B�f�[�^���A�C�������\�b�h�A�D�����ڍׁA�E���[�U�[ID�A�F�f�[�^�X�V�敪�j
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "���ޏ��ꊇ�o�^����_�����J�n�O", pstrTantousyaId, "2")
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

            If FncBunTouroku(strErrMsg) = True Then
                FncCSVIkkatsuTouroku = RET_FLD_BUN
            Else
                FncCSVIkkatsuTouroku = -1
            End If

            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
            '���O�o�́i�@�A�ԁA�A�����t�@�C���A�B�f�[�^���A�C�������\�b�h�A�D�����ڍׁA�E���[�U�[ID�A�F�f�[�^�X�V�敪�j
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "���ޏ��ꊇ�o�^����_�����I��", pstrTantousyaId, "3")
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

            '********** 2005/08/31 ���� �ǉ� **********��
        ElseIf intRet = RET_NO_DATA Then              '�f�[�^���P���������ꍇ
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
            pstrFileSyurui = "�f�[�^�Ȃ�"
            '���O�o�́i�@�A�ԁA�A�����t�@�C���A�B�f�[�^���A�C�������\�b�h�A�D�����ڍׁA�E���[�U�[ID�A�F�f�[�^�X�V�敪�j
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "�f�[�^���P�����i�V����", pstrTantousyaId, "3")
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

            FncCSVIkkatsuTouroku = RET_NO_DATA
        ElseIf intRet = -2 Then             '�u�t�@�C�����J���܂���B�v�G���[�̏ꍇ
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
            pstrFileSyurui = "OPEN�G���["
            '���O�o�́i�@�A�ԁA�A�����t�@�C���A�B�f�[�^���A�C�������\�b�h�A�D�����ڍׁA�E���[�U�[ID�A�F�f�[�^�X�V�敪�j
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "�t�@�C�����J���܂��񔻒�", pstrTantousyaId, "3")
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

            FncCSVIkkatsuTouroku = -2
        Else
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
            pstrFileSyurui = "���̑��G���["
            '���O�o�́i�@�A�ԁA�A�����t�@�C���A�B�f�[�^���A�C�������\�b�h�A�D�����ڍׁA�E���[�U�[ID�A�F�f�[�^�X�V�敪�j
            Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "���̑��G���[����", pstrTantousyaId, "3")
            '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

            FncCSVIkkatsuTouroku = -1       '����ȊO�i�G���[�j�̏ꍇ
        End If

        Call SetIkkatuLog("1", "", "", "", "OK", "", "UPD")

        '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
        '���O�o�́i�@�A�ԁA�A�����t�@�C���A�B�f�[�^���A�C�������\�b�h�A�D�����ڍׁA�E���[�U�[ID�A�F�f�[�^�X�V�敪�j
        Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncCSVIkkatsuTouroku", "FncCSVIkkatsuTouroku_�����I��", pstrTantousyaId, "4")
        '*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    End Function

'*< UNISYSTEM >****************************************************'

'*�y�v���V�[�W�����z
'*!    GetCsvLineData

'*�y�����T�v�z
'*?    CSV�t�@�C���̃f�[�^��1�����z��ɓ]������
'*�y�p�����[�^�z
'*P    arrayData;      �z��; CSV�t�@�C���̃f�[�^;

'*�y�߂�l�z
'*R    �Ȃ�
'******************************************************************'
'*�y�ύX�����z
'*M    2008/08/26; ����(��A); �@�@�@�@�@�V�K�쐬;
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

'*�y�v���V�[�W�����z
'*!    FncCSVToArray

'*�y�����T�v�z
'*?    CSV�t�@�C���̃f�[�^��z��ɓ����B
'*�y�p�����[�^�z
'*P    strErrMsg;      string; �G���[���b�Z�[�W;

'*�y�߂�l�z
'*R    int; RET_FLD_HACH:   �������~���p�f�[�^;
'*R         RET_FLD_TENB:   �X�ʏ��p�f�[�^;
'*R         RET_NO_DATA:    �f�[�^����;
'*R         -1:             �G���[;
'*R         -2:             �G���[(�t�@�C�����J���܂���B);

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/18; ����(���j�V�X�e���������); �V�K�쐬;
'*M    2009/06/06; Supattra T.(Tostem Thai);  Add case CSV�t�@�C�� ;
'*M    2009/07/31; �a�c                     ;  �܂链�̃`�F�b�N�ǉ� ;
'*M    2011/02/10; �a�c                     ;  �Z���^�[�R�[�h�ǉ� ;
'*< Comment End >**************************************************'
Private Function FncCSVToArray(ByRef strErrMsg) As Integer
    Dim intRecNum, intIndex
    Dim lngCount As Long
    Dim intFile
    Dim varData As Object      'CSV�f�[�^(1�s)
    Dim conIntKoumoku           '�C���f�b�N�X�ԍ��̍ő�l�i�v�f���m�F�p�j

    FncCSVToArray = -1
    
    intFile = FreeFile          '�g�p�\�ȃt�@�C���i���o�[���擾�i�t�@�C������ɕK�v�j

    'CSV�t�@�C������f�[�^���擾����B
        On Error Resume Next

    '�t�@�C�����J��
        'Open pstrFilePath For Input Access Read Lock Read As #intFile
        'If Err.Number <> 0 Then
        '    FncCSVToArray = -2
        '    strErrMsg = "�t�@�C�����J���܂���B(" & Err.Description & ")"
        '    Exit Function
        '    End If



        On Error GoTo 0

        'CSV�t�@�C���̃f�[�^��z��paryCsvData�Ɋi�[
        lngCount = 0
        intRecNum = 0

        For Each varData In File.ReadAllLines(pstrFilePath)
            '�z��̊g��
            ReDim Preserve paryCsvData(intRecNum)

            '�t�B�[���h�����ł́A�ǂ���̈ꊇ�o�^���𔻕ʂ���B�i�������~��or�X�ʏ��j
            If lngCount = 0 Then
                If pstrKikakuSyurui = "012" Or pstrKikakuSyurui = "015" Then
                    '2009.06.06 Add by Supattra  T.(Tostem Thai)Start.
                    If varData = FLD_MODOSI Then       '�X�ʏ��
                        FncCSVToArray = RET_FLD_MODO
                        conIntKoumoku = enmModColNo.PopSize
                    Else
                        FncCSVToArray = -1
                        strErrMsg = "�ꊇ�o�^�p(�܂链)��CSV�t�@�C���ł͂���܂���B"
                        Exit Function
                    End If
                Else
                    '2009.06.06 Add by Supattra  T.(Tostem Thai)End.
                    If varData = FLD_HACCHUU Then           '�������~��
                        FncCSVToArray = RET_FLD_HACH
                        conIntKoumoku = enmHacColNo.DaiJanCd
                    ElseIf varData = FLD_TENBETU Then  '�X�ʏ��
                        'FncCSVToArray = RET_FLD_TENB
                        'conIntKoumoku = enmTenColNo.CenterKbn

                        FncCSVToArray = -1
                        strErrMsg = "���o�[�W�����̓o�^�[�ł��B�ŐV�̓o�^�[���_�E�����[�h���Ă��������B"
                        Exit Function

                    ElseIf varData = FLD_SYOUHIN Then       '���i��{���
                        FncCSVToArray = RET_FLD_SYO
                        conIntKoumoku = enmSyoColNo.SvhJyuyouCd
                    ElseIf varData = FLD_BUNRUI Then        '���ޏ��
                        FncCSVToArray = RET_FLD_BUN
                        conIntKoumoku = enmBunColNo.Bunrui3
                    ElseIf varData = FLD_MODOSI Then        '�߂�����
                        FncCSVToArray = RET_FLD_MODO
                        conIntKoumoku = enmModColNo.PopSize
                    Else                                    '�G���[
                        FncCSVToArray = -1
                        strErrMsg = "�ꊇ�o�^�p��CSV�t�@�C���ł͂���܂���B"

                        Exit Function
                    End If
                End If

                '�f�[�^�����́A�J���}�ŕ������Ĕz��Ɋi�[
            ElseIf lngCount >= 1 Then
                paryCsvData(intRecNum) = Split(varData, ",")

                '�G���[�`�F�b�N
                If IsArray(paryCsvData(intRecNum)) = False Then       '�t�@�C���`���̊m�F
                    FncCSVToArray = -1
                    strErrMsg = "CSV�t�@�C���̌`�����s���ł��B"

                    Exit Function
                End If
                If UBound(paryCsvData(intRecNum)) <> conIntKoumoku Then  '�v�f���̊m�F
                    FncCSVToArray = -1
                    strErrMsg = "CSV�t�@�C���̍��ڐ��ɉߕs��������܂��B"
                    Exit Function
                End If

                '�e���ڂ̗��[�̋󔒂���菜��
                intIndex = 0
                For intIndex = 0 To conIntKoumoku Step 1
                    paryCsvData(intRecNum)(intIndex) = Trim(paryCsvData(intRecNum)(intIndex))
                Next

                intRecNum = intRecNum + 1
            End If
            varData = ""
            lngCount = lngCount + 1
        Next
        Do While Not EOF(intFile)           '�t�@�C���̏I�[�܂Ń��[�v���J��Ԃ��܂��B
        
        Loop


        '*** ��2018/11/20 �R�c �\�����ʃG���[�Ή� ************************************************************************
        '�f�[�^�������擾
        plngDataCnt = lngCount
        '*** ��2018/11/20 �R�c �\�����ʃG���[�Ή� ************************************************************************

        '�Y���f�[�^�������ꍇ
        If intRecNum = 0 Then
            FncCSVToArray = RET_NO_DATA
        End If

        Dim blnNotAllBlank As Boolean

        '********** 2005/05/31 �k�V **********��
        '�X��CSV�ŃL�[�ȊO���S�ău�����N��������Y���f�[�^�������b�Z�[�W��\������(���ƕʂ͕K�{���ڂ�����̂ł��蓾�Ȃ�)
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
        '********** 2005/05/31 �k�V **********��

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

        '********** 2005/06/15 �k�V **********��
        '�o�^�敪=2�̓X�R�[�h�̕ۊǗp�z����Ē�`����
        If intRecNum > 0 Then
            ReDim arrTourokuKbn2(intRecNum - 1, 2)
        End If
        '********** 2005/06/15 �k�V **********��

    End Function

'*< UNISYSTEM >****************************************************'

'*�y�v���V�[�W�����z
'*!    FncHacTouroku

'*�y�����T�v�z
'*?    �������~���̃f�[�^�����i�I��WF��DB�Ɏ�荞��

'*�y�p�����[�^�z
'*P    strErrMsg;        string; �G���[���b�Z�[�W;

'*�y�߂�l�z
'*R    Boolean; TRUE:�����AFALSE:���s;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/25; ����(���j�V�X�e���������); �V�K�쐬;
'*M    2005/05/12; �y�c�@�S��(USC); �p�t�H�[�}���X����̈פDB�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h��p�ӂ��鏈���ǉ�;
'*M    2006/09/13; �a�c(IC);        ���i�I��v�e�⍇���o�ŉ��P;
'*< Comment End >**************************************************'
Private Function FncHacTouroku(ByRef strErrMsg) As Boolean
    Dim intN As Integer
    Dim blnRet As Boolean
    Dim strJigCD As String
    Dim intNowLine As Integer
    Dim strCSVRecord() As String
    Dim flgChkError As Boolean      '�`�F�b�N���̃G���[�t���O�iTRUE:�G���[����j
    
    FncHacTouroku = False
    
'********** 2005/05/18 �y�c �ǉ� **********��
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(�Y�����ƃR�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����)"
'********** 2005/05/18 �y�c �ǉ� **********��

    '�� �ꊇ�o�^���O�ǉ� 2018/12/17
    Call SetIkkatuLog("1-1", "�������~���̃f�[�^�����i�I��WF��DB�Ɏ�荞��", "FncHacTouroku", "", "", pstrTantousyaId, "INS")
   
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    '���O�o��
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

'********** 2005/05/12 �y�c �ǉ� **********��
'DB�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h�Z�b�g�̐���
    Dim strSQL          As String
    Dim objCDBAccess    As Object
        
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    '�ꊇ�o�^���A�������[�U�[�̂ݎ��s�ł���
    Call LockTableIkkatu(objCDBAccess)

'�Y�����ƃR�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    'SQL���쐬�i�����F�ݒ���e��4���ڂ�1�AKEY�敪=W02�j
    Call SetIkkatuLog("1-1-1", "�Y�����ƃR�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")
    strSQL = "SELECT key_cd "
    strSQL = strSQL & "FROM m_wf_meisyou WITH(readcommitted) "
    strSQL = strSQL & "WHERE Substring(settei_naiyou,4,1) = '1' " & vbCrLf
    strSQL = strSQL & "AND key_kbn = '" & KEYCD_KIKAKU_DOUNYU_TAISYOU & "' " & vbCrLf
    
        pHChkJigyouRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pHChkJigyouRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
        Call SetIkkatuLog("1-1-1", "", "", "", "OK", "", "UPD")
        Call SetIkkatuLog("1-1-2", "MST_KEY�}�X�^1��KEY�R�[�h�𒊏o����i�����R�[�h�̏d�����`�F�b�N����ׁj", "SQL", "", "", pstrTantousyaId, "INS")
        
'�d��JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
'********** 2005/05/18 �y�c �ǉ� **********��
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(�d��JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����)"
'********** 2005/05/18 �y�c �ǉ� **********��

'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    '���O�o��
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

'********** 2005/09/12 ���� �ǉ� **********��
'VMST_KEY�}�X�^1��KEY�R�[�h�𒊏o����i�����R�[�h�̏d�����`�F�b�N����ׁj
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
    pChkSameJANCodeRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    pChkSameJANCodeRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
'********** 2005/09/12 ���� �ǉ� **********��
'CSV�t�@�C����Jan�R�[�h�̃��X�g(Jan�R�[�h�̏d������)�쐬
'********** 2005/05/18 �y�c �ǉ� **********��
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(CSV�t�@�C����Jan�R�[�h�̃��X�g(Jan�R�[�h�̏d������)�쐬)"
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    '���O�o��
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

    intNowLine = 0
    Do Until UBound(paryCsvData) < intNowLine
        
        strCSVRecord = paryCsvData(intNowLine)
         
        If InStr(pCsvJanCodeList, "'" & strCSVRecord(enmHacColNo.janCD) & "'") = 0 Then
            pCsvJanCodeList = pCsvJanCodeList & "'" & strCSVRecord(enmHacColNo.janCD) & "',"
        End If

'********** 2005/05/19 �y�c �ǉ� **********��
'���Jan�R�[�h�����X�g�ɉ�����
        If InStr(pCsvJanCodeList, "'" & strCSVRecord(enmHacColNo.DaiJanCd) & "'") = 0 Then
            pCsvJanCodeList = pCsvJanCodeList & "'" & strCSVRecord(enmHacColNo.DaiJanCd) & "',"
        End If
'********** 2005/05/19 �y�c �ǉ� **********��

        intNowLine = intNowLine + 1
    Loop
    
    If pCsvJanCodeList <> "" Then
        pCsvJanCodeList = Left(pCsvJanCodeList, Len(pCsvJanCodeList) - 1)
    Else
        pCsvJanCodeList = "''"
    End If
    intNowLine = 0
'********** 2005/05/18 �y�c �ǉ� **********��
    Call SetIkkatuLog("1-1-2", "", "", "", "OK", "", "UPD")
    Call SetIkkatuLog("1-1-3", " �����}�X�^�o�^�ς�JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")
    
'�����}�X�^�o�^�ς�JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
'********** 2005/05/18 �y�c �ύX **********��
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(�����}�X�^�o�^�ς�JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����)"
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    '���O�o��
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

'    strSQL = "SELECT key_cd , jan_cd FROM v_key_1 WITH(readcommitted)"
    strSQL = "SELECT jan_cd, key_cd FROM v_key_1 WITH(readcommitted)"
    strSQL = strSQL & " WHERE jan_cd IN(" & pCsvJanCodeList & ") "
    

        pChkvk1JANCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)

'    pChkvk1JANCodeRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    pChkvk1JANCodeRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
'********** 2005/05/18 �y�c �ύX **********��

'********** 2005/05/18 �y�c �ǉ� **********��
'�����}�X�^�o�^�ς�JAN�R�[�h�R�t����L�[�̃��X�g(��L�[�d������)�쐬
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(CSV�t�@�C����JAN�R�[�h�R�t����L�[�̃��X�g(��L�[�d������)�쐬)"
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    '���O�o��
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    
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
'********** 2005/05/18 �y�c �ǉ� **********��

    
'********** 2005/05/18 �y�c �ǉ� **********��
    Call SetIkkatuLog("1-1-4", "�����J�n���`�F�b�N�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")

'�����J�n���`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(�����J�n���`�F�b�N�p�ؒf���R�[�h�Z�b�g����)"
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    '���O�o��
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

    strSQL = "SELECT jan_cd, jigyou_cd, hattyuu_kaisi_date"
    strSQL = strSQL & " FROM t_syousen_jigyoubetu_settei WITH(readcommitted)"
    strSQL = strSQL & " WHERE kikaku_syurui_cd = '" & pstrKikakuSyurui & "' "
    strSQL = strSQL & " AND nen_ren_no = '" & pstrKikakuNenrenban & "' "
    strSQL = strSQL & " AND eda_no = '" & pstrKikakuEdaban & "'"
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "
    

        pHChkHattyuuKaisiDateRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)

    pHChkHattyuuKaisiDateRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
'(���i�I��WF(���ƕʐݒ��񍀖�)�Ƀf�[�^�������ꍇ)�����J�n�������ƕʏ��e�[�u������擾�p�ؒf���R�[�h�Z�b�g����
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(�����J�n�������ƕʏ��e�[�u������擾�p�ؒf���R�[�h�Z�b�g����)"
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    '���O�o��
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    Call SetIkkatuLog("1-1-4", "", "", "", "OK", "", "UPD")
   
    Call SetIkkatuLog("1-1-5", "(���i�I��WF(���ƕʐݒ��񍀖�)�Ƀf�[�^�������ꍇ)�����J�n�������ƕʏ��e�[�u������擾�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")

    strSQL = "SELECT key_cd, jigyou_cd, hattyuu_kaisi_date"
    strSQL = strSQL & " FROM v_jigyou_betu WITH(readcommitted)"
    strSQL = strSQL & " WHERE key_cd IN(" & pHKeyCDList & ") "
    
        pHChkJHTYKaisiDateRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pHChkJHTYKaisiDateRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U

    Call SetIkkatuLog("1-1-5", "", "", "", "OK", "", "UPD")
'********** 2005/05/18 �y�c �ǉ� **********��
    
'�������~���R�敪�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    Call SetIkkatuLog("1-1-6", "�������~���R�敪�`�F�b�N�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")
'********** 2005/05/18 �y�c �ǉ� **********��
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(�������~���R�敪�`�F�b�N�p�ؒf���R�[�h�Z�b�g����)"
'********** 2005/05/18 �y�c �ǉ� **********��
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    '���O�o��
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    
    strSQL = "SELECT key_cd "
    strSQL = strSQL & " FROM m_wf_meisyou WITH(readcommitted)"
    strSQL = strSQL & " WHERE key_kbn = '" & KEYCD_HATTYUU_TYUUSI_RIYU_KBN & "'"

        pHChkHacchuRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pHChkHacchuRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U

    Call SetIkkatuLog("1-1-6", "", "", "", "OK", "", "UPD")
'********** 2005/05/12 �y�c �ǉ� **********��

    Call SetIkkatuLog("1-1-7", "�S�Ă̔z��f�[�^�̃`�F�b�N���s���B", "SQL", "", "", pstrTantousyaId, "INS")
    '�S�Ă̔z��f�[�^�̃`�F�b�N���s���B
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(�S�Ă̔z��f�[�^�̃`�F�b�N���s���B)"
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    '���O�o��
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    
    flgChkError = False
    intNowLine = 0
    strErrMsg = ""
    Do Until UBound(paryCsvData) < intNowLine
        
        strCSVRecord = paryCsvData(intNowLine)
        
        '(1)DB�Ɉˑ����Ȃ��`�F�b�N�FFncCheckParameter_CorrectForm_Hac
        '(2)DB���Q�Ƃ��ă`�F�b�N:FncCheckParameter_DataBase_Hac
        If FncCheckParameter_CorrectForm_Hac(strCSVRecord, intNowLine + 1) = False Or _
            FncCheckParameter_DataBase_Hac(strCSVRecord, intNowLine + 1) = False Then
            flgChkError = True
        End If
        
        intNowLine = intNowLine + 1
    Loop
    Call SetIkkatuLog("1-1-7", "", "", "", "OK", "", "UPD")
    
    Call SetIkkatuLog("1-1-8", "B�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h�Z�b�g�̊J��", "SQL", "", "", pstrTantousyaId, "INS")
'********** 2005/05/12 �y�c �ǉ� **********��
'DB�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h�Z�b�g�̊J��
'********** 2005/05/18 �y�c �ǉ� **********��
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(DB�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h�Z�b�g�̊J��)"
'********** 2005/05/18 �y�c �ǉ� **********��
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    '���O�o��
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
        
    pHChkJigyouRS.Close
        pHChkJigyouRS = Nothing
        
    pChkSameJANCodeRS.Close
        pChkSameJANCodeRS = Nothing
    
    pChkvk1JANCodeRS.Close
        pChkvk1JANCodeRS = Nothing
    
    pHChkHacchuRS.Close
        pHChkHacchuRS = Nothing
        
    'DB�N���[�Y
    objCDBAccess = Nothing
    'GetObjectContext.SetComplete
        Call SetIkkatuLog("1-1-8", "", "", "", "OK", "", "UPD")
        
'********** 2005/05/12 �y�c �ǉ� **********��
    
    '�f�[�^�ɂP�ł���肪����ꍇ�́A�����ŏI���I�I�I
    If flgChkError = True Then
        '�G���[������΃G���[���b�Z�[�W���擾����B
        strErrMsg = FncGetErrorMsg_Hac()
        
'********** 2005/05/19 �y�c �ǉ� **********��
        If Not pHChkHattyuuKaisiDateRS Is Nothing Then
                pHChkHattyuuKaisiDateRS = Nothing
        End If
        If Not pHChkJHTYKaisiDateRS Is Nothing Then
                pHChkJHTYKaisiDateRS = Nothing
        End If
 '********** 2005/05/19 �y�c �ǉ� **********��
        
        Exit Function
    End If
    
'********** 2005/05/19 �y�c �ǉ� **********��
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTouroku(�f�[�^�X�V�����J�n)"
'********** 2005/05/19 �y�c �ǉ� **********��
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    '���O�o��
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", strMethodCallInfo, pstrTantousyaId, "2")
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************

    Call SetIkkatuLog("1-2", "�f�[�^�X�V�����J�n", "SQL", "", "", pstrTantousyaId, "INS")
    
    intN = 0
    '�X�ʏ��p�t���O���ꉞ�N���A����B�i�����ł͎g�p���Ȃ��j
    pflgBaikaToukbn = False
    pflgSiireToukbn = False
    pflgCenterToukbn = False
    
    '�z��f�[�^��DB�Ɏ�荞��
    Do Until UBound(paryCsvData) < intN         '�z��̏I�[�܂Ń��[�v���J��Ԃ��܂��B
        '���i�I��DB(WF�Ǘ�����)�̐�t�ύX�̓o�^�敪�X�V�m�F�p�t���O
        '���ƕʏ��o�^�敪�͏�ɍX�V����ׁA�t���O�𗧂Ă�B
        pflgJigJyoToukbn = True
    
        '���i�I��DB(���ƕʐݒ��񍀖�)�̒ǉ�or�X�V
        strJigCD = paryCsvData(intN)(enmHacColNo.jigCD)      '�Y�����ƃR�[�h���擾
        If strJigCD = ALL_JIG_CD Then
            blnRet = FncHacTourokuAllJig(strErrMsg, intN)
        Else
            blnRet = FncHacTourokuJig(strErrMsg, intN, strJigCD)
        End If

        If blnRet = False Then      '�G���[����
'********** 2005/05/19 �y�c �C�� **********��
'            Exit Function           '�G���[����
            GoTo DefectErrorHandler:
'********** 2005/05/19 �y�c �C�� **********��
        End If

'********** 2005/05/07 �k�V **********��
'����P���Ƃ̃��R�[�h���쐬����ہA���̑S�Ă̎��Ƃ̃��R�[�h���쐬
        If strJigCD <> ALL_JIG_CD Then
            blnRet = Insert_t_syousen_jigyoubetu_settei(strErrMsg, intN _
                   , pstrKikakuSyurui, pstrKikakuNenrenban, pstrKikakuEdaban _
                   , paryCsvData(intN)(enmTenColNo.janCD) _
                   , strJigCD _
                   , pstrTantousyaId)
            If blnRet = False Then
'********** 2005/05/19 �y�c �C�� **********��
'              Exit Function           '�G���[����
              GoTo DefectErrorHandler:
'********** 2005/05/19 �y�c �C�� **********��
            End If
        End If
'********** 2005/05/07 �k�V **********��
        
'********** 2005/05/02 �k�V **********��
'VSWF�\���d�l�e�[�u���Ƀ��R�[�h�����݂��Ȃ��ꍇ��VMST��ʊ�{�e�[�u����������擾�����R�[�h��ǉ�����
'(�Ɖ�Ȃǂŏ��i�����\������Ȃ���)
        blnRet = Insert_t_syousen_hyouji_siyou(strErrMsg, intN _
                , pstrKikakuSyurui, pstrKikakuNenrenban, pstrKikakuEdaban _
                , paryCsvData(intN)(enmTenColNo.janCD) _
                , pstrTantousyaId)
        If blnRet = False Then
'********** 2005/05/19 �y�c �C�� **********��
'            Exit Function           '�G���[����
            GoTo DefectErrorHandler:
'********** 2005/05/19 �y�c �C�� **********��
        End If
'********** 2005/05/02 �k�V **********��
        
        '���i�I��DB(WF�Ǘ�����)�̍X�V
        If FncSetWFKanri(strErrMsg, intN, RET_FLD_HACH) = False Then
'********** 2005/05/19 �y�c �C�� **********��
'            Exit Function           '�G���[����
            GoTo DefectErrorHandler:
'********** 2005/05/19 �y�c �C�� **********��
        End If
        
        intN = intN + 1
    Loop
    
    FncHacTouroku = True

'********** 2005/05/19 �y�c �ǉ� **********��
    If Not pHChkHattyuuKaisiDateRS Is Nothing Then
            pHChkHattyuuKaisiDateRS = Nothing
    End If
    If Not pHChkJHTYKaisiDateRS Is Nothing Then
            pHChkJHTYKaisiDateRS = Nothing
    End If
 '********** 2005/05/19 �y�c �ǉ� **********��

'********** 2005/05/18 �y�c �ǉ� **********��
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
    '���O�o��
    Call SetLog(plngSaibanNo2, pstrFileSyurui, plngDataCnt, "FncHacTouroku", "FncHacTouroku_�����I��", pstrTantousyaId, "3")
'*** ��2018/11/19 �R�c �\�����ʃG���[�Ή� ****************************************************************************************
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
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description
    
    
'********** 2005/05/18 �y�c �ǉ� **********��

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckParameter_CorrectForm_Hac
'*�y�����T�v�z
'*?    CSV�t�@�C���̃p�����[�^�ɑ΂��āADB�Ɉˑ����Ȃ��`���`�F�b�N���s���܂��B
'*?    �i�������~���p�`�F�b�N�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckParameter_CorrectForm; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/25; ���@����(USC); �V�K�쐬;
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
'********** 2005/05/07 �k�V **********��
    Dim blnIsOKNotZero
'********** 2005/05/07 �k�V **********��
    
    'a.�K�{���ځi�Y�����ƃR�[�h�A�����I�����AJAN�R�[�h�j�������Ă��邩�m�F����B
        blnIsOKRequired = FncCheck_RequiredParameter_Hac(arryCSVRecord, intNowLine)
    
    'b.�S�Ă̍��ڂɂ����āA���p�����ł��邩�m�F����B
        blnIsOKNumeric = FncCheck_IsParameterNumeric_Hac(arryCSVRecord, intNowLine)
    
    'c.�����������𒴂��Ă��Ȃ����m�F����B
        blnIsOKLimit = FncCheck_Limit_Hac(arryCSVRecord, intNowLine)

    'd.�����I�����̓��t���L���ł��邩�m�F����B
    blnIsOKHacEndDay = FncCheck_HacEndDay_Hac(arryCSVRecord(enmHacColNo.SyuryouDate), intNowLine)
        
    'e.JAN�R�[�h�ɐ������`�F�b�N�f�W�b�g�̒l���Z�b�g����Ă��邩�𒲂ׂ�
    blnIsOKCheckDigit = FncCheck_CheckDigit(arryCSVRecord(enmHacColNo.janCD), intNowLine, 2)
    
    '���JAN�R�[�h�ɐ������`�F�b�N�f�B�W�b�g�̒l���Z�b�g����Ă��邩�𒲂ׂ�
    blnIsOKCheckDigit_DaiJAN = True
    If arryCSVRecord(enmHacColNo.DaiJanCd) <> "" Then
        blnIsOKCheckDigit_DaiJAN = FncCheck_CheckDigit(arryCSVRecord(enmHacColNo.DaiJanCd), intNowLine, 3)
    End If
    
    'f.�������@��"0"�i���ݒ�j�A"1"�i�ԕi�j�A"2"�i�ꎞ��~�j�A"3"�i�戵�I���i�X�U��Ώہj�j��
    '�����ꂩ�ł��邩�m�F����B
    blnIsOKSyobun = True
    If arryCSVRecord(enmHacColNo.Syobun) <> "" Then
        blnIsOKSyobun = FncCheck_CheckSyobun(arryCSVRecord(enmHacColNo.Syobun), intNowLine)
    End If
    
'********** 2005/05/07 �k�V **********��
'���z�ނ����̒l�����`�F�b�N����
        blnIsOKNotZero = FncCheck_NotZero(arryCSVRecord, intNowLine, "HAC")
'********** 2005/05/07 �k�V **********��
    
'********** 2005/05/07 �k�V **********��
'    If blnIsOKRequired And blnIsOKNumeric And blnIsOKLimit And _
'       blnIsOKHacEndDay And blnIsOKCheckDigit And blnIsOKCheckDigit_DaiJAN And _
'       blnIsOKSyobun Then
    If blnIsOKRequired And blnIsOKNumeric And blnIsOKLimit And _
       blnIsOKHacEndDay And blnIsOKCheckDigit And blnIsOKCheckDigit_DaiJAN And _
       blnIsOKSyobun And blnIsOKNotZero Then
'********** 2005/05/07 �k�V **********��
        FncCheckParameter_CorrectForm_Hac = True
    End If
    
End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_RequiredParameter_Hac
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�K�{���ڂ��󔒂łȂ����ǂ����`�F�b�N���܂��B
'*?    �i�������~���p�`�F�b�N�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_RequiredParameter; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/25; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'

Private Function FncCheck_RequiredParameter_Hac(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_RequiredParameter_Hac = True
    
    If arryCSVRecord(enmHacColNo.jigCD) = "" Then       '�Y�����ƃR�[�h
            FncSetError(intNowLine, enmTenColNo.TenCd, PARAM_EMPTY)
        FncCheck_RequiredParameter_Hac = False
    End If
    If arryCSVRecord(enmHacColNo.SyuryouDate) = "" Then  '�����I����
            FncSetError(intNowLine, enmTenColNo.HenkouDate, PARAM_EMPTY)
        FncCheck_RequiredParameter_Hac = False
    End If
    If arryCSVRecord(enmHacColNo.janCD) = "" Then       'JAN�R�[�h
            FncSetError(intNowLine, enmTenColNo.janCD, PARAM_EMPTY)
        FncCheck_RequiredParameter_Hac = False
    End If

'********** 2005/05/07 �k�V **********��
'�������~���R�͕K�{
    If arryCSVRecord(enmHacColNo.ChushiKbn) = "" Then
            FncSetError(intNowLine, enmHacColNo.ChushiKbn, PARAM_EMPTY)
        FncCheck_RequiredParameter_Hac = False
    End If
'�������~���R��00(���I��)�͕s��
    If arryCSVRecord(enmHacColNo.ChushiKbn) = "00" Then
            FncSetError(intNowLine, enmHacColNo.ChushiKbn, PARAM_MISENTAKU)
        FncCheck_RequiredParameter_Hac = False
    End If
'********** 2005/05/07 �k�V **********��

'********** 2005/05/09 �k�V **********��
''********** 2005/05/07 �k�V **********��
''�������@���ꎞ�I�ɓ��͕K�{�ɂ���(�������@����}�X�^����擾���郍�W�b�N������������폜����)
'    If arryCSVRecord(enmHacColNo.Syobun) = "" Then
'        FncSetError (intNowLine, enmHacColNo.Syobun, PARAM_EMPTY
'        FncCheck_RequiredParameter_Hac = False
'    End If
''��֏��i�R�[�h���ꎞ�I�ɓ��͕K�{�ɂ���(��֏��i�R�[�h����}�X�^����擾���郍�W�b�N������������폜����)
'    If arryCSVRecord(enmHacColNo.DaiJanCd) = "" Then
'        FncSetError (intNowLine, enmHacColNo.DaiJanCd, PARAM_EMPTY
'        FncCheck_RequiredParameter_Hac = False
'    End If
''********** 2005/05/07 �k�V **********��
'********** 2005/05/09 �k�V **********��

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_IsParameterNumeric_Hac
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�S�t�B�[���h�����p���l���ǂ����`�F�b�N���܂��B
'*?    �i�������~���p�`�F�b�N�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_IsParameterNumeric_Hac; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/25; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterNumeric_Hac(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_IsParameterNumeric_Hac = True

    If FncChkHankakuNumber(arryCSVRecord(enmHacColNo.jigCD)) = False Then       '�Y�����ƃR�[�h
        FncSetError (intNowLine, enmTenColNo.TenCd, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Hac = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmHacColNo.SyuryouDate)) = False Then '�����I����
        FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Hac = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmHacColNo.janCD)) = False Then       'JAN�R�[�h
        FncSetError (intNowLine, enmTenColNo.janCD, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Hac = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmHacColNo.Syobun)) = False Then      ' �������@
        FncSetError (intNowLine, enmHacColNo.Syobun, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Hac = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmHacColNo.ChushiKbn)) = False Then   ' �������~���R�敪
    
        FncSetError (intNowLine, enmHacColNo.ChushiKbn, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Hac = False
        
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmHacColNo.DaiJanCd)) = False Then    ' ���JAN�R�[�h
    
        FncSetError (intNowLine, enmHacColNo.DaiJanCd, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Hac = False
    End If

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_Limit_Hac
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�������������z���Ă��Ȃ������m�F���܂��B
'*?    �i�������~���p�`�F�b�N�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_Limit_Hac; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/25; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
Function FncCheck_Limit_Hac(arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_Limit_Hac = True
    
    '(b)JAN�R�[�h
    If FncCheck_JANCode(arryCSVRecord(enmHacColNo.janCD)) = False Then
        FncSetError (intNowLine, enmHacColNo.janCD, PARAM_SHORTDISIT)
        FncCheck_Limit_Hac = False
    End If

    '(c)���JAN�R�[�h
    If arryCSVRecord(enmHacColNo.DaiJanCd) <> "" Then
    
        If FncCheck_JANCode(arryCSVRecord(enmHacColNo.DaiJanCd)) = False Then
            FncSetError (intNowLine, enmHacColNo.DaiJanCd, PARAM_SHORTDISIT)
            FncCheck_Limit_Hac = False
        End If
    End If

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_HacEndDay_Hac
'*�y�����T�v�z
'*?    �������~���̓��t���L���ł��邩���m�F���܂��B
'*?    �i�������~���p�`�F�b�N�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strHacEndDay ; String ; �����I����������킷������
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_CheckSyobun; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/25; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
Function FncCheck_HacEndDay_Hac(strHacEndDay As String, intNowLine As Integer)

    FncCheck_HacEndDay_Hac = True
    ' �������W���ł��鎖�ƁA�ύX���̓��t���L���ł��邩���m�F����B
    If Len(strHacEndDay) <> 8 Or FncCheckDate(strHacEndDay) = False Then
        FncSetError (intNowLine, enmHacColNo.SyuryouDate, PARAM_WRONGDATE)
        FncCheck_HacEndDay_Hac = False
    End If

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_CheckSyobun
'*�y�����T�v�z
'*?    �������@���K�肳��Ă���l�Ɋ܂܂�邱�Ƃ��m�F���܂��B
'*?    �i�������~���p�`�F�b�N�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strSyobun ; String ; �������@������킷������
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_CheckSyobun; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/25; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'

Function FncCheck_CheckSyobun(strSyobun As String, intNowLine As Integer) As Boolean
    
    ' �������@��0,1,2,3�̂����ꂩ�ł��邱�Ƃ��m���߂�
    
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
'*�y�v���V�[�W�����z
'*!    FncCheckParameter_DataBase_Hac
'*�y�����T�v�z
'*?    CSV�t�@�C���̃p�����[�^�ɑ΂��āADB�ɃA�N�Z�X���ă`�F�b�N���s���܂��B
'*?    (�������~��)
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVRecord() ; String CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckParameter_DataBase_Hac; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/25; ���@����(USC); �V�K�쐬;
'*M    2005/09/12; �����@����(IC);  �ύX �����R�[�h�̏d���`�F�b�N
'*< Comment End >**************************************************'
Private Function FncCheckParameter_DataBase_Hac(ByRef arryCSVRecord() As String, intNowLine As Integer)

    Dim blnIsOKJigyouCode
    Dim blnIsOKSyuryouDate
    Dim blnIsOKJANCode
    Dim blnIsOKNaibuCode
    Dim blnIsOKChushiKbn
    Dim blnIsOKDaiJANCode
    
    Dim strKeyCode As String
    
    'a.�Y�����ƃR�[�h��WF���̃}�X�^�e�[�u���ɓo�^����Ă��邩�m�F����B
    blnIsOKJigyouCode = FncCheckDB_JigyouCode(arryCSVRecord(enmHacColNo.jigCD), intNowLine)

    'b1.JAN�R�[�h���L�����m�F����B
    blnIsOKJANCode = FncCheckDB_JANCode(arryCSVRecord(enmHacColNo.janCD), strKeyCode, intNowLine, RET_FLD_HACH)
    
    'b2.KEY�R�[�h���L�����m�F����B
    blnIsOKNaibuCode = FncCheckDB_NaibuCode(strKeyCode, intNowLine, RET_FLD_HACH)
    
    'c.�����I�����������J�n������ł��邩�m�F����B
    '(JAN�R�[�h���L�[�Ƃ��Č������邽�߁AJAN�R�[�h�̗L�������Ɋm�F����F�L���Ȃ�Ό������s���j
    blnIsOKSyuryouDate = True
    If blnIsOKJANCode = True Then
        blnIsOKSyuryouDate = FncCheckDB_SyuryouDate(arryCSVRecord, strKeyCode, intNowLine)
    End If
    
    'd.�������~���R�敪���L�����m�F����B
    If arryCSVRecord(enmHacColNo.ChushiKbn) <> "" Then
        blnIsOKChushiKbn = FncCheckDB_ChushiKbn(arryCSVRecord(enmHacColNo.ChushiKbn), intNowLine)
    Else
        blnIsOKChushiKbn = True
    End If
    
    'e.���JAN�R�[�h�������}�X�^�ɓo�^����Ă��邩�m�F����B
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
'*�y�v���V�[�W�����z
'*!     FncCheckDB_JigyouCode
'*�y�����T�v�z
'*?     ���ƃR�[�h��WF���̃}�X�^�ɓo�^����Ă��邱�Ƃ��m�F����B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P     strJigyouCode ; String ; ���ƃR�[�h��\��������
'*P     intNowLine ; Integer ; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R     FncCheckDB_JigyouCode ; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/25; ���@����(USC); �V�K�쐬;
'*M    2005/05/12; �y�c�@�S��(USC); �p�t�H�[�}���X����̈פ���̊֐����ł̐ؒf���R�[�h�擾����ߑ��ϐ��̐ؒf���R�[�h�𗘗p����l�C��;
'*< Comment End >**************************************************'
'
Private Function FncCheckDB_JigyouCode(strJigyouCode As String, intNowLine As Integer)

'********** 2005/05/12�y�c �폜 **********��
'���̊֐����ł̐ؒf���R�[�h�Z�b�g�擾�����폜
'    Dim strSql            As String
'
'    '�@�W�J���Ă��鎖�ƃR�[�h�́AWF���̃}�X�^����擾����B
'    '   �@DB�F      WF���̃}�X�^�im_wf_meisyou�j
'    '   �@�L�[�F        KEY�敪�ikey_kbn�j                          �uW02�v
'    '           KEY�R�[�h�ikey_cd�j                         CSV�t�@�C������擾
'    '           �ݒ���e�isettei_naiyou�j                           4���ڂ�"1"�������Ă���
'    '�@�S���Ɓi�Y�����ƃR�[�h���Ɂu99�v�w��j�̏ꍇ�́AKEY�R�[�h���w�肹���Ɍ���
'    '�@���鎖�ɂ��A�W�J���Ă��鎖�ƃR�[�h�S�Ă��擾�ł���B
'    '�@�����ɂP�����q�b�g���Ȃ��ꍇ�́A�G���[�Ƃ���B
'
'    strSql = "SELECT key_cd "
'    strSql = strSql & "FROM m_wf_meisyou WITH(readcommitted) "
'    strSql = strSql & "WHERE Substring(settei_naiyou,4,1) = '1' " & vbCrLf
'    strSql = strSql & "AND key_kbn = '" & KEYCD_KIKAKU_DOUNYU_TAISYOU & "' " & vbCrLf
'
'    '�S���ƂłȂ��ꍇ�́A���ƃR�[�h���w�肷��B
'    If strJigyouCode <> ALL_JIG_CD Then
'        strSql = strSql & "AND key_cd = '" & strJigyouCode & "'"      '�e����
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
'********** 2005/05/12�y�c �폜 **********��
    
'********** 2005/05/12�y�c �ǉ� **********��
'���ϐ��̐ؒf���R�[�h�Z�b�g���g�p���G���[�`�F�b�N
    With pHChkJigyouRS
'********** 2005/05/17 �y�c �ύX **********��
'********** 2005/05/19 �y�c �ύX **********��
'       If Not .EOF Then
       .Filter = adFilterNone
       If Not .BOF Then
'********** 2005/05/19 �y�c �ύX **********��
            .MoveFirst
            '�S���ƂłȂ��ꍇ�́A���ƃR�[�h���w�肷��B
            If strJigyouCode <> ALL_JIG_CD Then
                    .Find("key_cd = '" & strJigyouCode & "'")
            End If
        End If
'********** 2005/05/17 �y�c �ύX **********��
    End With
    
    If pHChkJigyouRS.EOF Then
            FncSetError(intNowLine, enmHacColNo.jigCD, PARAM_NOTFINDMASTER)
    End If

    FncCheckDB_JigyouCode = Not pHChkJigyouRS.EOF
'********** 2005/05/12�y�c �ǉ� **********��
    
End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!     FncCheckDB_SyuryouDate
'*�y�����T�v�z
'*?     �����I�����������J�n��������̓��t�ł��邩�ǂ������m�F���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P     arryCSVRecord() ; String CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P     strKeyCode ; String �����}�X�^����擾�����AJAN�R�[�h�ɕR�t����ꂽKeyCode
'*P     intNowLine ; Integer CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R     FncCheckDB_SyuryouDate ; Boolean �`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/25; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'

Private Function FncCheckDB_SyuryouDate(ByRef arryCSVRecord() As String, strKeyCode As String, intNowLine As Integer)
    
    Dim strKaisiDate As String
    Dim IsOKGetKaisiDate As Boolean
    
    FncCheckDB_SyuryouDate = False
    
    ' �����J�n�����擾����B
    
'********** 2005/05/18 �y�c �ύX **********��
'    If arryCSVRecord(enmHacColNo.jigCD) = ALL_JIG_CD Then
        IsOKGetKaisiDate = FncCheckDB_GetKaisiDateAllJig(arryCSVRecord, strKaisiDate, strKeyCode)
'    Else
'        IsOKGetKaisiDate = FncCheckDB_GetKaisiDate(arryCSVRecord, strKaisiDate, strKeyCode)
'    End If
'********** 2005/05/18 �y�c �ǉ� **********��
    
    ' �����J�n���������I�������ǂ����`�F�b�N����B
    If IsOKGetKaisiDate = True Then
        
        ' �����J�n�����������I�����̏ꍇ�̓G���[
        If StrComp(strKaisiDate, arryCSVRecord(enmHacColNo.SyuryouDate)) >= 0 Then
            FncCheckDB_SyuryouDate = False
                FncSetError(intNowLine, enmHacColNo.SyuryouDate, PARAM_EARLYCANCELDAY)
            Exit Function
        End If
        
        ' ���ׂĂ̔����J�n���͔����I�������O�̓��t
        FncCheckDB_SyuryouDate = True
    
    Else
        ' �����J�n����������Ȃ��ꍇ
            FncSetError(intNowLine, enmHacColNo.SyuryouDate, PARAM_NOTEXISTDB)
        FncCheckDB_SyuryouDate = False
    
    End If
        
End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!     FncCheckDB_GetKaisiDateAllJig
'*�y�����T�v�z
'*?     ���i�I��WF(���ƕʐݒ��񍀖�)���甭���J�n�����擾���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P     arryCSVRecord() ; String CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P     strKaisiDate ; String �����J�n�����i�[���镶����
'*P     strKeyCode ; String ; �����}�X�^����擾�����AJAN�R�[�h�ɕR�t����ꂽKeyCode
'*�y�߂�l�z
'*    < Name Type Comment >
'*R     FncCheckDB_GetKaisiDateAllJig ; Boolean �J�n���擾����:True,���s:False
'*R     strKaisiDate ; �����J�n��()���i�[
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/27; ����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheckDB_GetKaisiDateAllJig(ByRef arryCSVRecord() As String, ByRef strKaisiDate As String, strKeyCode As String) As Boolean

'********** 2005/05/18 �y�c �폜 **********��
'    Dim objCDBAccess As Object
'    Dim objRSJig As Object
'    Dim objRS As Object
'
'    Dim strSQL As String
'    Dim strJigCd As String
'    Dim strTmpKaisiDate As String       '�e���Ɩ��̔����J�n��
'    Dim dtKaisiDate As Date
'
'    Dim strMethodCallInfo As String
'    on error GoTo DefectErrorHandler
'    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncCheckDB_GetKaisiDateAllJig()"
'
'    'CDBAccess�C���X�^���X����
'    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
'
'    '�W�J���Ă��鎖�ƃR�[�h���擾����B
'    strSQL = "SELECT key_cd "
'    strSQL = strSQL & "FROM m_wf_meisyou WITH(readcommitted) "
'    strSQL = strSQL & "WHERE Substring(settei_naiyou,4,1) = '1' " & vbCrLf
'    strSQL = strSQL & "AND key_kbn = '" & KEYCD_KIKAKU_DOUNYU_TAISYOU & "' " & vbCrLf
'    'SQL���s
'        objRSJig = objCDBAccess.GetDCRecordset("VSWF", strSQL)
'
'    If objRSJig.EOF = True Then
'        '�W�J���Ă��鎖�ƃR�[�h�������ꍇ�́A�����ŏI���B
'        FncCheckDB_GetKaisiDateAllJig = False
'        GoTo Proc_End:
'    Else
'        Do Until objRSJig.EOF
'            strJigCd = CStr(objRSJig.Fields("key_cd"))
'
'            '�e���Ɩ��ɔ����J�n�����擾����B
'            '�@�����J�n���ƏI�������A�����J�n���������I�����ƂȂ��Ă��邩�m�F����B
'            '   �@DB�F      ���i�I��DB�i���ƕʐݒ��񍀖ځj�it_syousen_jigyoubetu_settei�j
'            '   �@�L�[�F        ����޺���(kikaku_syurui_cd)                          �O��ʂ���擾
'            '           �N�A�ԁinen_ren_no�j                            �O��ʂ���擾
'            '           �}�ԁieda_no�j                          �O��ʂ���擾
'            '           JAN�R�[�h�ijan_cd�j                         CSV�t�@�C������擾
'            '           ���ƃR�[�h�ijigyou_cd�j                         CSV�t�@�C������擾
'
'            strSQL = "SELECT hattyuu_kaisi_date"
'            strSQL = strSQL & " FROM t_syousen_jigyoubetu_settei WITH(readcommitted)"
'            strSQL = strSQL & " WHERE kikaku_syurui_cd = '" & pstrKikakuSyurui & "' "
'            strSQL = strSQL & " AND nen_ren_no = '" & pstrKikakuNenrenban & "' "
'            strSQL = strSQL & " AND eda_no = '" & pstrKikakuEdaban & "'"
'            strSQL = strSQL & " AND jan_cd = '" & arryCSVRecord(enmHacColNo.JanCd) & "'"
'            strSQL = strSQL & " AND jigyou_cd = '" & strJigCd & "'"
'
'            'SQL���s
'                objRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
'
'            If objRS.EOF = True Then
'                '���i�I��WF(���ƕʐݒ��񍀖�)�Ƀf�[�^�������ꍇ�́A�����}�X�^(���ƕʏ��)����擾����B
'                If FncCheckDB_GetKaisiDateMas(strKaisiDate, strKeyCode, strJigCd) = False Then
'                    FncCheckDB_GetKaisiDateAllJig = False
'                    GoTo Proc_End:
'                End If
'            Else    '�����J�n���f�[�^��NULLor�󔒂ł���ꍇ�́A�����}�X�^(���ƕʏ��)����擾����B
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
'                    ' ��t�ύXDB�e�[�u�����甭���J�n�����擾
'                    ' ��t�ύXDB�̓��t�������قȂ邽�߁A�����łW���ɒ���
'                    dtKaisiDate = CDate(Trim(objRS.Fields("hattyuu_kaisi_date")))
'                    strKaisiDate = Format(dtKaisiDate, "yyyymmdd")
'                End If
'            End If
'
'            '�ł��V���������J�n���݂̂�ϐ��Ɋi�[����B
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
'    'DB�N���[�Y
''    objRS.Close
'        objRSJig = Nothing
'        objRS = Nothing
'    objCDBAccess = Nothing
'
'    'GetObjectContext.SetComplete
'********** 2005/05/18 �y�c �폜 **********��

'********** 2005/05/18 �y�c �ǉ� **********��

    Dim strTmpKaisiDate As String       '�e���Ɩ��̔����J�n��
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    
    '�`�F�b�N�Ώێ��ƃR�[�h�̎擾
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.FncCheckDB_GetKaisiDateAllJig(�`�F�b�N�Ώێ��ƃR�[�h�̎擾)"
    
    '���̃��R�[�h�̎��ƃR�[�h
    Dim jigCD As String
    jigCD = arryCSVRecord(enmHacColNo.jigCD)
    
    '�`�F�b�N�Ώێ��ƃR�[�h�̎擾�ɊY�����ƃR�[�h�`�F�b�N�pRS���ė��p����
    With pHChkJigyouRS

        '***  �`�F�b�N�Ώێ��ƃR�[�h�̍i���� ***
        
        '�t�B���^�[���������Ă���
        .Filter = adFilterNone
        
        If jigCD = ALL_JIG_CD Then
            '�S���Ƃ̏ꍇ�t�B���^�����O����
        Else
            '�e���Ƃ̏ꍇ
            '�Y�����ƃR�[�h�`�F�b�N�pRS�����ƃR�[�h�Ńt�B���^�����O����
            .Filter = "key_cd = '" & jigCD & "'"
        End If
        
        If .EOF Then
            '�W�J���Ă��鎖�ƃR�[�h�������ꍇ�́A�����ŏI���B
            FncCheckDB_GetKaisiDateAllJig = False
            Exit Function
        Else
            '***  �`�F�b�N�Ώێ��ƃR�[�h�̃��[�v ***
            Do Until .EOF
                Dim strJigCD As String
                strJigCD = CStr(.Fields("key_cd"))
                
                With pHChkHattyuuKaisiDateRS
                    '�����J�n�����擾����B
                    .Filter = "jan_cd = '" & arryCSVRecord(enmHacColNo.janCD) & "' AND jigyou_cd = '" & strJigCD & "'"
                    
                    '���i�I��WF(���ƕʐݒ��񍀖�)�Ƀf�[�^�������ꍇ�A����
                    '�����J�n���f�[�^��NULL�A���͋󔒂ł���ꍇ�A�����}�X�^(���ƕʏ��)����擾����B
                    'strMethodCallInfo = App.Title & ".VswfIkatuTou.FncCheckDB_GetKaisiDateAllJig(�`�F�b�N�Ώێ��ƃR�[�h�̎擾:�����}�X�^(���ƕʏ��)����擾����B)"
                    
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
                            ' ��t�ύXDB�e�[�u�����甭���J�n�����擾
                            ' ��t�ύXDB�̓��t�������قȂ邽�߁A�����łW���ɒ���
'                            strKaisiDate = Format(CDate(Trim(.Fields("hattyuu_kaisi_date"))), "yyyymmdd")
                            strTmpKaisiDate = Format(CDate(Trim(.Fields("hattyuu_kaisi_date"))), "yyyymmdd")
                        End If
                    End If
                    
                    '�ł��V���������J�n���݂̂�ϐ��Ɋi�[����B
                    If StrComp(strTmpKaisiDate, strKaisiDate) > 0 Then
                        strKaisiDate = strTmpKaisiDate
                    End If
                End With
                .MoveNext
            Loop
        End If
    End With
    FncCheckDB_GetKaisiDateAllJig = True
'********** 2005/05/18 �y�c �ǉ� **********��

Exit Function

'********** 2005/05/18 �y�c �폜 **********��
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
'********** 2005/05/18 �y�c �폜 **********��
DefectErrorHandler:
'********** 2005/05/18 �y�c �폜 **********��
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
'********** 2005/05/18 �y�c �폜 **********��
'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!     FncCheckDB_GetKaisiDate
'*�y�����T�v�z
'*?     ���i�I��WF(���ƕʐݒ��񍀖�)���甭���J�n�����擾���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P     arryCSVRecord() ; String CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P     strKaisiDate ; String �����J�n�����i�[���镶����
'*P     strKeyCode ; String ; �����}�X�^����擾�����AJAN�R�[�h�ɕR�t����ꂽKeyCode
'*�y�߂�l�z
'*    < Name Type Comment >
'*R     FncCheckDB_GetKaisiDate ; Boolean �J�n���擾����:True,���s:False
'*R     strKaisiDate ; �����J�n��()���i�[
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/25; ���@����(USC); �V�K�쐬;
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

    '�@�����J�n���ƏI�������A�����J�n���������I�����ƂȂ��Ă��邩�m�F����B
    '   �@DB�F      ���i�I��DB�i���ƕʐݒ��񍀖ځj�it_syousen_jigyoubetu_settei�j
    '   �@�L�[�F        ����޺���(kikaku_syurui_cd)                          �O��ʂ���擾
    '           �N�A�ԁinen_ren_no�j                            �O��ʂ���擾
    '           �}�ԁieda_no�j                          �O��ʂ���擾
    '           JAN�R�[�h�ijan_cd�j                         CSV�t�@�C������擾
    '           ���ƃR�[�h�ijigyou_cd�j                         CSV�t�@�C������擾

    strSQL = "SELECT hattyuu_kaisi_date"
    strSQL = strSQL & " FROM t_syousen_jigyoubetu_settei WITH(readcommitted)"
    strSQL = strSQL & " WHERE kikaku_syurui_cd = '" & pstrKikakuSyurui & "' "
    strSQL = strSQL & " AND nen_ren_no = '" & pstrKikakuNenrenban & "' "
    strSQL = strSQL & " AND eda_no = '" & pstrKikakuEdaban & "'"
    strSQL = strSQL & " AND jan_cd = '" & arryCSVRecord(enmHacColNo.janCD) & "'"
    strSQL = strSQL & " AND jigyou_cd = '" & arryCSVRecord(enmHacColNo.jigCD) & "'"
        
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL���s
        objRs = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    
    If objRs.EOF = True Then
        '���i�I��WF(���ƕʐݒ��񍀖�)�Ƀf�[�^�������ꍇ�́A�����}�X�^(���ƕʏ��)����擾����B
        If FncCheckDB_GetKaisiDateMas(strKaisiDate, strKeyCode, arryCSVRecord(enmHacColNo.jigCD)) = False Then
            FncCheckDB_GetKaisiDate = False
            GoTo Proc_End:
        End If
    Else    '�����J�n���f�[�^��NULLor�󔒂ł���ꍇ�́A�����}�X�^(���ƕʏ��)����擾����B
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
                ' ��t�ύXDB�e�[�u�����甭���J�n�����擾
                ' ��t�ύXDB�̓��t�������قȂ邽�߁A�����łW���ɒ���
                dtKaisiDate = CDate(Trim(objRs.Fields("hattyuu_kaisi_date")))
                strKaisiDate = Format(dtKaisiDate, "yyyymmdd")
            End If
    End If
    
    FncCheckDB_GetKaisiDate = True
    
    'DB�N���[�Y
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
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!     FncCheckDB_GetKaisiDateMas
'*�y�����T�v�z
'*?     �����}�X�^(���ƕʏ��)���甭���J�n�����擾���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P     strKaisiDate ; String �����J�n�����i�[���镶����
'*P     strKeyCode ; String ; �����}�X�^����擾�����AJAN�R�[�h�ɕR�t����ꂽKeyCode
'*P     strJigCd ; String ; ���ƃR�[�h
'*�y�߂�l�z
'*    < Name Type Comment >
'*R     FncCheckDB_GetKaisiDateMas ; Boolean �J�n���擾����:True,���s:False
'*R     strKaisiDate ; �����J�n��()���i�[
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/27; ���@����(USC); �V�K�쐬;
'*M    2005/05/18; �y�c�@�S��(USC); �p�t�H�[�}���X����̈פ���̊֐����ł̐ؒf���R�[�h�擾����ߑ��ϐ��̐ؒf���R�[�h�𗘗p����l�C��;
'*< Comment End >**************************************************'
Private Function FncCheckDB_GetKaisiDateMas(ByRef strKaisiDate As String, _
                                      ByVal strKeyCode As String, _
                                      ByVal strJigCD As String) As Boolean
                    
'********** 2005/05/18 �y�c �폜 **********��
'    Dim objCDBAccess As Object
'    Dim objRS As Object
'    Dim strSQL As String
'    Dim strTmp As String
'********** 2005/05/18 �y�c �폜 **********��
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncCheckDB_GetKaisiDateMas()"
    
'********** 2005/05/18 �y�c �폜 **********��
'    '�@��t�ύXDB�ɓo�^����Ă��Ȃ��ꍇ�́A�����J�n�������ƕʏ��e�[�u������
'    '�@�\�ߎ擾���Ă����B
'    '   �@DB�F      ���ƕʏ��e�[�u���im_jigyou_betu�j
'    '   �@�L�[�F        ��L�[�ikey_cd�j                            1.(2)c.(b) KEY���1����擾
'    '           ���ƃR�[�h�ijigyou_cd�j                         CSV�t�@�C������擾
'    strSQL = "SELECT hattyuu_kaisi_date"
'    strSQL = strSQL & " FROM v_jigyou_betu WITH(readcommitted)"
'    strSQL = strSQL & " WHERE key_cd = '" & strKeyCode & "'" & vbCrLf
'    strSQL = strSQL & " AND jigyou_cd = '" & strJigCd & "'"
'
'    'CDBAccess�C���X�^���X����
'    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
'
'    'SQL���s
'        objRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
'********** 2005/05/18 �y�c �폜 **********��
    

'********** 2005/05/18 �y�c �ύX **********��
    With pHChkJHTYKaisiDateRS
        .Filter = "key_cd = '" & strKeyCode & "' AND jigyou_cd = '" & strJigCD & "'"
    
    
'        If objRS.EOF = True Then
        If .EOF = True Then
            FncCheckDB_GetKaisiDateMas = False
'            GoTo Proc_End:
            Exit Function
        Else    '�����J�n���f�[�^��NULLor�󔒂ł���ꍇ�́A�G���[
'�����������ɂ܂Ƃ߂�
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
                ' ���ƕʏ��e�[�u�����甭���J�n�����擾
'                strKaisiDate = Trim(objRS.Fields("hattyuu_kaisi_date"))     'char�^�ׁ̈A���̂܂ܓ����B
                strKaisiDate = Trim(.Fields("hattyuu_kaisi_date"))     'char�^�ׁ̈A���̂܂ܓ����B
            End If
        End If
    End With
'********** 2005/05/18 �y�c �ύX **********��
    
    FncCheckDB_GetKaisiDateMas = True
    

'********** 2005/05/18 �y�c �폜 **********��
'    'DB�N���[�Y
'    objRS.Close
'        objRS = Nothing
'    objCDBAccess = Nothing
'
'    'GetObjectContext.SetComplete
'********** 2005/05/18 �y�c �폜 **********��
Exit Function

'********** 2005/05/18 �y�c �폜 **********��
'Proc_End:
'    If Not objRS Is Nothing Then
'            objRS = Nothing
'    End If
'    If Not objCDBAccess Is Nothing Then
'        objCDBAccess = Nothing
'    End If
'    Exit Function
'********** 2005/05/18 �y�c �폜 **********��
DefectErrorHandler:
'********** 2005/05/18 �y�c �폜 **********��
'    If Not objRS Is Nothing Then
'            objRS = Nothing
'    End If
'    If Not objCDBAccess Is Nothing Then
'        objCDBAccess = Nothing
'    End If
'    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
'********** 2005/05/18 �y�c �폜 **********��
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł�a�B", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!     FncCheckDB_ChushiKbn
'*�y�����T�v�z
'*?     �������~���R�敪�����������m�F����B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P     strChushiKbn ; �������~�敪��\��������
'*P     intNowLine ; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R     FncCheckDB_ChushiKbn ; Boolean �`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/25; ���@����(USC); �V�K�쐬;
'*M    2005/05/12; �y�c�@�S��(USC); �p�t�H�[�}���X����̈פ���̊֐����ł̐ؒf���R�[�h�擾����ߑ��ϐ��̐ؒf���R�[�h�𗘗p����l�C��;
'*< Comment End >**************************************************'
'
Private Function FncCheckDB_ChushiKbn(strChushiKbn As String, intNowLine As Integer) As Boolean

    '�@�������~���R�敪��WF���̃}�X�^�e�[�u���ɓo�^����Ă��邩�m�F����B
    '   �@DB�F      WF���̃}�X�^�im_wf_meisyou�j
    '   �@�L�[�F        KEY�敪�ikey_kbn�j                          �uM08�v
    '           KEY�R�[�h�ikey_cd�j                         CSV�t�@�C������擾
    '
    '
'********** 2005/05/12�y�c �폜 **********��
'���̊֐����ł̐ؒf���R�[�h�Z�b�g�擾�����폜
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
'********** 2005/05/12�y�c �폜 **********��

'********** 2005/05/12�y�c �ǉ� **********��
'���ϐ��̐ؒf���R�[�h�Z�b�g���g�p���G���[�`�F�b�N
    With pHChkHacchuRS
'********** 2005/05/17 �y�c �ύX **********��
'********** 2005/05/18 �y�c �ύX **********��
'        If Not .EOF Then
       .Filter = adFilterNone
        If Not .BOF Then
'********** 2005/05/18 �y�c �ύX **********��
            .MoveFirst
                .Find("key_cd = '" & strChushiKbn & "'")
        End If
'********** 2005/05/17 �y�c �ύX **********��
    End With
    
    If pHChkHacchuRS.EOF Then
            FncSetError(intNowLine, enmHacColNo.ChushiKbn, PARAM_NOTFINDMASTER)
    End If

    FncCheckDB_ChushiKbn = Not pHChkHacchuRS.EOF
'********** 2005/05/12�y�c �ǉ� **********��


End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!     FncCheckDB_DaiJANCode
'*�y�����T�v�z
'*?     ���JAN�R�[�h���L�������m�F����B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P     strDaiJANCode ; String ; ���JAN�R�[�h��\��������
'*P     intNowLine As Integer ; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R     FncCheckDB_DaiJANCode ; Boolean ;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/25; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
'
Private Function FncCheckDB_DaiJANCode(strDaiJANCode As String, intNowLine As Integer) As Boolean

    '�@�}�X�^�o�^����Ă��邩�́A�����}�X�^��KEY���1�e�[�u���Ŋm�F����B
    '   �@DB�F      �����}�X�^ KEY���1�iv_key_1�j
    '   �@�L�[�F        JAN�R�[�h�ijan_cd�j                         CSV�t�@�C������擾
    
    Dim strKeyCode As String
    If FncCheckDB_EntryJANCode(strDaiJANCode, strKeyCode) = True Then
        FncCheckDB_DaiJANCode = True
    Else
            FncSetError(intNowLine, enmHacColNo.DaiJanCd, PARAM_NOTEXISTDB, strDaiJANCode)
        FncCheckDB_DaiJANCode = False
    End If
    
End Function

'*******************************************************************
' �������~���p�`�F�b�N���[�`�������܂�
'*******************************************************************

'*< UNISYSTEM >****************************************************'

'*�y�v���V�[�W�����z
'*!    FncHacTourokuAllJig

'*�y�����T�v�z
'*?    ���Ə��}�X�^���玖�ƃR�[�h���擾���āA���ƃR�[�h����DB�Ɏ�荞��

'*�y�p�����[�^�z
'*P    strErrMsg;        string; �G���[���b�Z�[�W;
'*P    intN;             int; ���Ԗڂ̃��R�[�h��;

'*�y�߂�l�z
'*R    Boolean; TRUE:�����AFALSE:���s;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/25; ����(���j�V�X�e���������); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncHacTourokuAllJig(ByRef strErrMsg, ByVal intN) As Boolean
    Dim sql, objCDB, intSelNum, blnRet
    Dim rsSel As ADODB.Recordset
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTourokuAllJig(" & CStr(intN) + 1 & "�s�ځF�y���ƃR�[�h�F" & ALL_JIG_CD & "�z)"
    
    FncHacTourokuAllJig = False

    '���Ə��}�X�^���玖�ƃR�[�h���擾����B
        objCDB = CreateObject("COMDBUtil.CDBAccess")    'CDBAccess���C���X�^���X��

    'SQL���쐬�i�����F�X����NULLor�ꊇ�o�^���ȍ~�AKEY�敪=W02�A�ݒ���e�̂S���ڂ�1�j
    sql = "SELECT DISTINCT A.jigyou_cd AS jigcd " & vbCrLf
    sql = sql & "FROM v_jigyousyo AS A WITH(readcommitted) " & vbCrLf
    sql = sql & "INNER JOIN m_wf_meisyou AS B WITH(readcommitted) ON A.jigyou_cd = B.key_cd " & vbCrLf
    sql = sql & "WHERE " & vbCrLf
    sql = sql & " Substring(B.settei_naiyou,4,1) = '1' AND " & vbCrLf
    sql = sql & " B.key_kbn = '" & KEYCD_KIKAKU_DOUNYU_TAISYOU & "' " & vbCrLf
    
        rsSel = objCDB.GetDCRecordset("VSWF", sql)
        objCDB = Nothing
    If rsSel.EOF = True Then    '�Ώۂ̎��ƃR�[�h�������ꍇ�̓G���[
        strErrMsg = CStr(intN + 1) & "�s�ځF�y���ƃR�[�h�F" & ALL_JIG_CD & "�z�W�J���Ă��鎖�ƃR�[�h�͂P��������܂���ł����B"
        rsSel.Close
            rsSel = Nothing
        Exit Function
    End If
    
    '���ƺ��ޖ��Ƀf�[�^�����i�I��WF��DB�Ɏ�荞�ށiFncHacTourokuJig�֐����Ăԁj�B
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
    
    'DB�N���[�Y
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
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'

'*�y�v���V�[�W�����z
'*!    FncHacTourokuJig

'*�y�����T�v�z
'*?    �w�肵�����ƃR�[�h�̃f�[�^���ADB�Ɏ�荞��

'*�y�p�����[�^�z
'*P    strErrMsg;        string; �G���[���b�Z�[�W;
'*P    intN;             int; ���Ԗڂ̃��R�[�h��;
'*P    strJigCd;         string; ���ƃR�[�h;

'*�y�߂�l�z
'*R    Boolean; TRUE:�����AFALSE:���s;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/25; ����(���j�V�X�e���������); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncHacTourokuJig(ByRef strErrMsg, ByVal intN, ByVal strJigCD) As Boolean
    Dim sql, objCDB, intNum, blnRet
    Dim strSelSql, strInsSql, strUpdSql
    Dim arrParam() As Object
    Dim lngP_Cnt
    Dim rsSel
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncHacTourokuJig(" & CStr(intN) + 1 & "�s�ځF�y���ƃR�[�h�F" & strJigCD & "�z)"
    
    FncHacTourokuJig = False

    '�w�肳�ꂽ���ƃR�[�h�̃f�[�^��ǉ�or�X�V����B
        objCDB = CreateObject("COMDBUtil.CDBAccess")    'CDBAccess���C���X�^���X��

    lngP_Cnt = 0
    Erase arrParam
    'SQL���쐬�i�����FKEY�敪=W02�A�X����NULLor�ꊇ�o�^���ȍ~�A�ݒ���e�̂S���ڂ�1�j
    Call SubHacSqlSelJig(intN, strJigCD, strSelSql, arrParam, lngP_Cnt)
    Call SubHacSqlInsJig(intN, strJigCD, strInsSql, arrParam, lngP_Cnt)
    Call SubHacSqlUpdJig(intN, strJigCD, strUpdSql, arrParam, lngP_Cnt)
    sql = "If (" & strSelSql & ") = 0 " & vbCrLf        '�f�[�^�����ɑ��݂��邩�`�F�b�N
    sql = sql & "BEGIN " & strInsSql & " END " & vbCrLf  '���݂��Ȃ��ꍇ�́A�ǉ�
    sql = sql & "ELSE " & vbCrLf
    sql = sql & "BEGIN " & strUpdSql & " END " & vbCrLf  '���݂���ꍇ�́A�X�V
    
    intNum = objCDB.ExecuteSql("VSWF", sql, arrParam)
    If intNum <= 0 Then    '�Ώۂ̎��ƃR�[�h�̓o�^�Ɏ��s�����ꍇ�A�G���[
        strErrMsg = CStr(intN + 1) & "�s�ځF�y���ƃR�[�h�F" & strJigCD & "�z�̃f�[�^�̓o�^�Ɏ��s���܂����B"
            objCDB = Nothing
        Exit Function
    End If
    
    FncHacTourokuJig = True
    
    'DB�N���[�Y
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
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'

'*�y�v���V�[�W�����z
'*!    FncTenTouroku

'*�y�����T�v�z
'*?    �z��Ɋi�[���ꂽ�X�ʏ��̃f�[�^��DB�Ɏ�荞��

'*�y�p�����[�^�z
'*P    strErrMsg;        string; �G���[���b�Z�[�W;

'*�y�߂�l�z
'*R    Boolean; TRUE:�����AFALSE:���s;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/20; ����(���j�V�X�e���������); �V�K�쐬;
'*M    2005/05/11; �y�c�@�S��(USC); �p�t�H�[�}���X����̈פDB�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h��p�ӂ��鏈���ǉ�;
'*M    2005/05/13; �c���@���l(USC); �p�t�H�[�}���X����΍�i�ǉ��j;
'*M    2006/09/13; �a�c(IC);        ���i�I��v�e�⍇���o�ŉ��P;
'*< Comment End >**************************************************'
Private Function FncTenTouroku(ByRef strErrMsg) As Boolean
    Dim intN As Integer
    Dim blnRet As Boolean
    Dim strTenCd As String
    Dim intNowLine As Integer
    Dim strCSVRecord() As String
    Dim flgChkError As Boolean      '�`�F�b�N���̃G���[�t���O�iTRUE:�G���[����j
    
    FncTenTouroku = False
    
    '�S�Ă̔z��f�[�^�̃`�F�b�N���s���B
    flgChkError = False
    intNowLine = 0
    strErrMsg = ""
    
'********** 2005/05/13 �c�� �ǉ� **********��
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(�Y�����ƃR�[�h�`�F�b�N)"
'********** 2005/05/13 �c�� �ǉ� **********��
    
    
'********** 2005/05/11 �y�c �ǉ� **********��
'DB�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h�Z�b�g�̐���
    Dim strSQL          As String
    Dim objCDBAccess    As Object
        
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    '�ꊇ�o�^���A�������[�U�[�̂ݎ��s�ł���
    Call LockTableIkkatu(objCDBAccess)
    
    Call SetIkkatuLog("1-1", "�X�ʏ��̃f�[�^��DB�Ɏ�荞��", "FncTenTouroku", "", "", pstrTantousyaId, "INS")
    
    
    Call SetIkkatuLog("1-1-1", "�Y�����ƃR�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")
'�Y�����ƃR�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    'SQL���쐬�i�����F���ʋ敪=5�A�X����NULLor�ꊇ�o�^���ȍ~�AKEY�敪=W02�A�ݒ���e��4���ڂ�1�A�X�R�[�h�j
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
    
    
    
    '�Y�����ƃR�[�h�`�F�b�N�p�@�ؒf���R�[�h�Z�b�g�擾
        pTChkJigyouRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkJigyouRS.Fields("jigyousyo_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
'�W�J�X�R�[�h���X�g�̐���
    
    With pTChkJigyouRS
'********** 2005/05/17 �y�c �ύX **********��
'********** 2005/05/18 �y�c �ύX **********��
'        If Not .EOF Then
        If Not .BOF Then
'********** 2005/05/18 �y�c �ύX **********��
            .MoveFirst
            Do Until .EOF
                pTTenCodeList = pTTenCodeList & "'" & .Fields("jigyousyo_cd") & "',"
                .MoveNext
            Loop
            .MoveFirst
        End If
'********** 2005/05/17 �y�c �ύX **********��
    End With
    
    pTTenCodeList = Left(pTTenCodeList, Len(pTTenCodeList) - 1)

        Call SetIkkatuLog("1-1-1", "", "", "", "OK", "", "UPD")



'********** 2005/05/19 �y�c �폜 **********��
''********** 2005/05/19 �y�c �ǉ� **********��
''�����}�X�^�o�^�ς�JAN�R�[�h�R�t����L�[�̃��X�g(��L�[�d������)�쐬
'    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(CSV�t�@�C����JAN�R�[�h�R�t����L�[�̃��X�g(��L�[�d������)�쐬)"
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
''********** 2005/05/19 �y�c �ǉ� **********��
'********** 2005/05/19 �y�c �폜 **********��


'�����ύX���R�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
 Call SetIkkatuLog("1-1-2", "�����ύX���R�`�F�b�N�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")
 
'********** 2005/05/13 �c�� �ǉ� **********��
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(�����ύX���R�`�F�b�N)"
'********** 2005/05/13 �c�� �ǉ� **********��
    
   'SQL���쐬�i�����FKEY�敪=025�A�����ύX���R�R�[�h�j
    strSQL = "SELECT key_cd FROM v_meisyou WITH(readcommitted) "
    strSQL = strSQL & "WHERE key_kbn = '" & KEYCD_BAIKARIYUU & "'"

        pTChkBaikaHenkouRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkBaikaHenkouRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
Call SetIkkatuLog("1-1-2", "", "", "", "OK", "", "UPD")
'�����`�F�b�N�p�ؒf���R�[�h�Z�b�g����
 
  Call SetIkkatuLog("1-1-3", "�����`�F�b�N�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")
'********** 2005/05/13 �c�� �ǉ� **********��
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(�����`�F�b�N)"
'********** 2005/05/13 �c�� �ǉ� **********��
    
   strSQL = "SELECT torihikisaki_cd FROM v_torihikisaki WITH(readcommitted)"
    strSQL = strSQL & "WHERE torikesi_date LIKE ' ' OR torikesi_date IS NULL"
    
        pTChkTorihikisakiRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkTorihikisakiRS.Fields("torihikisaki_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
Call SetIkkatuLog("1-1-3", "", "", "", "OK", "", "UPD")
'�Z���^�[�敪�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
 
'********** 2005/05/13 �c�� �ǉ� **********��
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(�Z���^�[�敪�`�F�b�N)"
'********** 2005/05/13 �c�� �ǉ� **********��
  Call SetIkkatuLog("1-1-4", "�Z���^�[�敪�`�F�b�N�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")
  
   strSQL = "SELECT key_cd FROM m_wf_meisyou WITH(readcommitted) "
    strSQL = strSQL & "WHERE key_kbn = '" & KEYCD_CENTER_KBN & "' "
    
        pTChkCenterRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkCenterRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
Call SetIkkatuLog("1-1-4", "", "", "", "OK", "", "UPD")
'********** 2005/05/11 �y�c �ǉ� **********��


'********** 2005/05/13 �c�� �ǉ� **********��
'��t�ύX���̕ύX���`�F�b�N�p�ؒf���R�[�h�Z�b�g�����i�����}�X�^�E�����ύX���j
  Call SetIkkatuLog("1-1-5", "��t�ύX���̕ύX���i�����}�X�^�E�����ύX���j", "SQL", "", "", pstrTantousyaId, "INS")
  
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(��t�ύX���̕ύX���i�����}�X�^�E�����ύX���j)"
    strSQL = "SELECT key_cd, ten_cd, henkou_date "
    strSQL = strSQL & "FROM v_baika_henkou WITH(readcommitted) "
    'strSQL = strSQL & "where 1=0 "    '���e�X�g
'********** 2005/05/19 �y�c �ǉ� �� �폜**********��
'    strSQL = strSQL & " WHERE key_cd IN(" & pTKeyCDList & ") "
'********** 2005/05/19 �y�c �ǉ� �� �폜**********��
    
        pTChkMstBaikaRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkMstBaikaRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    pTChkMstBaikaRS.Fields("henkou_date").Properties("Optimize") = True
Call SetIkkatuLog("1-1-5", "", "", "", "OK", "", "UPD")
  Call SetIkkatuLog("1-1-6", "��t�ύX���̕ύX���`�F�b�N�p�ؒf���R�[�h�Z�b�g�����i�����}�X�^�E�d���ύX���j", "SQL", "", "", pstrTantousyaId, "INS")
  
'��t�ύX���̕ύX���`�F�b�N�p�ؒf���R�[�h�Z�b�g�����i�����}�X�^�E�d���ύX���j
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(��t�ύX���̕ύX���i�����}�X�^�E�d���ύX���j)"
    strSQL = "SELECT key_cd, ten_cd, henkou_date "
    strSQL = strSQL & "FROM v_siire_henkou WITH(readcommitted) "
    'strSQL = strSQL & "where 1=0 "    '���e�X�g
'********** 2005/05/19 �y�c �ǉ� �� �폜**********��
'    strSQL = strSQL & " WHERE key_cd IN(" & pTKeyCDList & ") "
'********** 2005/05/19 �y�c �ǉ� �� �폜**********��

        pTChkMstSiireRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkMstSiireRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    pTChkMstSiireRS.Fields("henkou_date").Properties("Optimize") = True
Call SetIkkatuLog("1-1-6", "", "", "", "OK", "", "UPD")

  Call SetIkkatuLog("1-1-7", "��t�ύX���̕ύX���i�����}�X�^�E�Z���^�[�R�[�h�ύX���j", "SQL", "", "", pstrTantousyaId, "INS")
  
'��t�ύX���̕ύX���`�F�b�N�p�ؒf���R�[�h�Z�b�g�����i�����}�X�^�E�Z���^�[�R�[�h�ύX���j
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(��t�ύX���̕ύX���i�����}�X�^�E�Z���^�[�R�[�h�ύX���j)"
    strSQL = "SELECT key_cd, ten_cd, henkou_date "
    strSQL = strSQL & "FROM v_tenpo_center_cd_henkou WITH(readcommitted) "
    'strSQL = strSQL & "where 1=0 "    '���e�X�g
'********** 2005/05/19 �y�c �ǉ� �� �폜**********��
'    strSQL = strSQL & " WHERE key_cd IN(" & pTKeyCDList & ") "
'********** 2005/05/19 �y�c �ǉ� �� �폜**********��

        pTChkMstCenterRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkMstCenterRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    pTChkMstCenterRS.Fields("henkou_date").Properties("Optimize") = True
Call SetIkkatuLog("1-1-7", "", "", "", "OK", "", "UPD")
'********** 2005/05/13 �c�� �ǉ� **********��

'********** 2005/05/18 �y�c �ǉ� **********��
'CSV�t�@�C�����ő�A�ŏ����t�̏�����
    pstrMinDate = paryCsvData(intNowLine)(enmTenColNo.HenkouDate)
    pstrMaxDate = paryCsvData(intNowLine)(enmTenColNo.HenkouDate)
'********** 2005/05/18 �y�c �ǉ� **********��


'********** 2005/05/19 �y�c �ړ� **********��
'�d��JAN�R�[�h�`�F�b�N���R�[�h�Z�b�g�̔r�����b�N�̎��Ԃ�Z������׈ړ�

    '�d��JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    
    '********** 2005/05/13 �c�� �ǉ� **********��
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(�d��JAN�R�[�h�`�F�b�N)"
    '********** 2005/05/13 �c�� �ǉ� **********��
        
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
'        pChkSameJANCodeRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
'********** 2005/05/20�y�c �ύX **********��

'********** 2005/09/12 ���� �ύX **********��
  Call SetIkkatuLog("1-1-8", "�d��JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")
  
    '�d��JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(�d��JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����)"
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
    pChkSameJANCodeRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    pChkSameJANCodeRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
'********** 2005/09/12 ���� �ύX **********��
    
    '********** 2005/05/18 �y�c �ǉ� **********��
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(CSV�t�@�C����Jan�R�[�h�̃��X�g(Jan�R�[�h�̏d������)�쐬)"
    'CSV�t�@�C����Jan�R�[�h�̃��X�g(Jan�R�[�h�̏d������)���쐬����
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

Call SetIkkatuLog("1-1-9", "�����}�X�^�o�^�ς�JAN�R�[�h�`�F�b�N", "SQL", "", "", pstrTantousyaId, "INS")
        intNowLine = 0
    '********** 2005/05/18 �y�c �ǉ� **********��
        
    '�����}�X�^�o�^�ς�JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    '********** 2005/05/13 �c�� �ǉ� **********��
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(�����}�X�^�o�^�ς�JAN�R�[�h�`�F�b�N)"
    '********** 2005/05/13 �c�� �ǉ� **********��
        
'********** 2005/05/20�y�c �ύX **********��
'       strSQL = "SELECT jan_cd, key_cd  FROM v_key_1 WITH(readcommitted)"
       strSQL = "SELECT jan_cd, key_cd  FROM v_key_1 WITH(UPDLOCK)"
'********** 2005/05/20�y�c �ύX **********��
    '********** 2005/05/18 �y�c �ǉ� **********��
        strSQL = strSQL & " WHERE jan_cd IN(" & pCsvJanCodeList & ") "
'********** 2005/05/20�y�c �ǉ� **********��
        strSQL = strSQL & " SET LOCK_TIMEOUT -1"
'********** 2005/05/20�y�c �ǉ� **********��
    '********** 2005/05/18 �y�c �ǉ� **********��
        
            pChkvk1JANCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
        pChkvk1JANCodeRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U

'���d��JAN�R�[�h�`�F�b�N���R�[�h�Z�b�g�̔r�����b�N�̎��Ԃ�Z������׈ړ�
'********** 2005/05/19 �y�c �ړ� **********��


    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(�G���[�`�F�b�N)"
    
    Do Until UBound(paryCsvData) < intNowLine
        
        strCSVRecord = paryCsvData(intNowLine)
        
        '(1)DB�Ɉˑ����Ȃ��`�F�b�N�FFncCheckParameter_CorrectForm
        '(2)DB���Q�Ƃ��ă`�F�b�N:FncCheckParameter_DataBase
'********** 2005/06/17 �k�V **********��
'        If FncCheckParameter_CorrectForm(strCSVRecord, intNowLine + 1) = False Or _
'           FncCheckParameter_DataBase(strCSVRecord, intNowLine + 1) = False Then
'            flgChkError = True
'        End If
'�A�x���h����̈עDB�Ɉˑ����Ȃ��`�F�b�N��ŃG���[�������ꍇ�̂ݢDB���Q�Ƃ��ă`�F�b�N����s��
        If FncCheckParameter_CorrectForm(strCSVRecord, intNowLine + 1) = False Then
            flgChkError = True
        Else
            If FncCheckParameter_DataBase(strCSVRecord, intNowLine + 1) = False Then
                flgChkError = True
            End If
        End If
'********** 2005/06/17 �k�V **********��
        
'********** 2005/05/17 �y�c �ǉ� **********��
'CSV�t�@�C����Jan�R�[�h�̃��X�g(Jan�R�[�h�̏d������)���쐬����
'        If InStr(pTCsvJanCodeList, strCSVRecord(enmTenColNo.JanCd)) = 0 Then
'            pTCsvJanCodeList = pTCsvJanCodeList & "'" & strCSVRecord(enmTenColNo.JanCd) & "',"
'        End If
'********** 2005/05/17 �y�c �ǉ� **********��

        intNowLine = intNowLine + 1
    Loop
Call SetIkkatuLog("1-1-9", "", "", "", "OK", "", "UPD")
'********** 2005/05/17 �y�c �ǉ� **********��
'CSV�t�@�C����Jan�R�[�h�̃��X�g(Jan�R�[�h�̏d������)���쐬����
'    If pTCsvJanCodeList <> "" Then
'        pTCsvJanCodeList = Left(pTCsvJanCodeList, Len(pTCsvJanCodeList) - 1)
'    End If
'
'********** 2005/05/17 �y�c �ǉ� **********��
    
'********** 2005/05/11 �y�c �ǉ� **********��
'DB�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h�Z�b�g�̊J��
    
'********** 2005/05/11 �y�c �폜 **********��
'�f�[�^�o�^�����p����̂ŊJ�����Ȃ�
'        pTChkJigyouRS.Close
'            pTChkJigyouRS = Nothing
'********** 2005/05/11 �y�c �폜 **********��
        
'********** 2005/05/19 �y�c �폜 **********��
'�f�[�^�o�^���ς܂���܂Ŕr�����b�N��������̂ł����ł͊J�����Ȃ�
'        pChkSameJANCodeRS.Close
'            pChkSameJANCodeRS = Nothing
'
'********** 2005/05/19 �y�c �폜 **********��
       
'********** 2005/05/17 �y�c �폜 **********��
'�f�[�^�o�^�����p����̂ŊJ�����Ȃ�
'        pChkvk1JANCodeRS.Close
'            pChkvk1JANCodeRS = Nothing
'********** 2005/05/17 �y�c �폜 **********��
       
        pTChkBaikaHenkouRS.Close
            pTChkBaikaHenkouRS = Nothing
       
        pTChkTorihikisakiRS.Close
            pTChkTorihikisakiRS = Nothing
       
        pTChkCenterRS.Close
            pTChkCenterRS = Nothing
       
'********** 2005/05/13 �c�� �ǉ� **********��
'DB�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h�Z�b�g�̊J���i�ǉ��j
        pTChkMstBaikaRS.Close
            pTChkMstBaikaRS = Nothing
       
        pTChkMstSiireRS.Close
            pTChkMstSiireRS = Nothing
       
        pTChkMstCenterRS.Close
            pTChkMstCenterRS = Nothing
   
'********** 2005/05/13 �c�� �ǉ� **********��

        'DB�N���[�Y
        objCDBAccess = Nothing
        'GetObjectContext.SetComplete
    
'********** 2005/05/11 �y�c �ǉ� **********��
   
   
    '�f�[�^�ɂP�ł���肪����ꍇ�́A�����ŏI���I�I�I
    If flgChkError = True Then
        '�G���[������΃G���[���b�Z�[�W���擾����B�i�I�����Ȃ��j
        strErrMsg = FncGetErrorMsg()
        Exit Function
    End If
     
     'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(�f�[�^�o�^)"
        Call SetIkkatuLog("1-2", "�f�[�^�o�^", "FncTenTouroku", "", "", pstrTantousyaId, "INS")
'********** 2005/05/16 �y�c ���ǉ� **********��
    subIns_SetAllTables (paryCsvData)
'********** 2005/05/16 �y�c ���ǉ� **********��
        Call SetIkkatuLog("1-2", "", "", "", "OK", "", "UPD")

'********** 2005/05/20 �y�c �ǉ� **********��
'�r�����b�N����
        pChkSameJANCodeRS.Close
            pChkSameJANCodeRS = Nothing
        objCDBAccess = Nothing
        'GetObjectContext.SetComplete
        
        Call SetIkkatuLog("1-1", "", "", "", "OK", "", "UPD")
                
'********** 2005/05/20 �y�c �ǉ� **********��
    
      'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTouroku(�f�[�^�o�^�I��)"

'********** 2005/05/16 �y�c ���폜 **********��
'    '�z��f�[�^��DB�Ɏ�荞��
'    intN = 0
''********** 2005/04/25 ���� �ǉ� **********��
'    pflgJigJyoToukbn = False            '�������~���p�t���O���ꉞ�N���A����B�i�����ł͎g�p���Ȃ��j
''********** 2005/04/25 ���� �ǉ� **********��
'
'
'    Do Until UBound(paryCsvData) < intN         '�z��̏I�[�܂Ń��[�v���J��Ԃ��܂��B
'        '���i�I��DB(WF�Ǘ�����)�̐�t�ύX�̓o�^�敪�X�V�m�F�p�t���O���N���A
'        pflgBaikaToukbn = False
'        pflgSiireToukbn = False
'        pflgCenterToukbn = False
'
'        strTenCd = paryCsvData(intN)(enmTenColNo.TenCd)      '�Y���X�R�[�h���擾
'
'        '�Y���X�R�[�h�ʂɏ����𕪂���B
'        If strTenCd = ALL_TEN_CD Then           '�S�X
'            blnRet = FncTenTourokuJig(strErrMsg, intN, strTenCd)
'        ElseIf strTenCd = ALL_JIG_CD Then       '�S����
'            blnRet = FncTenTourokuJig(strErrMsg, intN, strTenCd)
'        ElseIf Len(strTenCd) = 2 Then           '����
'            blnRet = FncTenTourokuJig(strErrMsg, intN, strTenCd)
'        ElseIf Len(strTenCd) = 4 Then           '�X
'            blnRet = FncTenTourokuTen(strErrMsg, intN, strTenCd)
'        End If
'        If blnRet = False Then      '�G���[����
'            Exit Function
'        End If
'
''********** 2005/04/27 �k�V **********��
''VSWF�\���d�l�e�[�u���Ƀ��R�[�h�����݂��Ȃ��ꍇ��VMST��ʊ�{�e�[�u����������擾�����R�[�h��ǉ�����
''(�Ɖ�Ȃǂŏ��i�����\������Ȃ���)
'        blnRet = Insert_t_syousen_hyouji_siyou(strErrMsg, intN _
'                , pstrKikakuSyurui, pstrKikakuNenrenban, pstrKikakuEdaban _
'                , paryCsvData(intN)(enmTenColNo.JanCd) _
'                , pstrTantousyaId)
'        If blnRet = False Then
'            Exit Function
'        End If
''********** 2005/04/27 �k�V **********��
'
'        '���i�I��DB(WF�Ǘ�����)�̍X�V
''********** 2005/04/26 ���� �ύX **********��
'        If FncSetWFKanri(strErrMsg, intN, RET_FLD_TENB) = False Then
''        If FncUpdWFKanri(strErrMsg, intN) = False Then
''********** 2005/04/26 ���� �ύX **********��
'            Exit Function           '�G���[����
'        End If
'
'        intN = intN + 1
'    Loop
'
'********** 2005/05/16 �y�c ���폜 **********��

    FncTenTouroku = True
    
'********** 2005/05/13 �c�� �ǉ� **********��

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
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description
'********** 2005/05/13 �c�� �ǉ� **********��

End Function

'*< UNISYSTEM >****************************************************'

'*�y�v���V�[�W�����z
'*!    FncModosiKakaku

'*�y�����T�v�z
'*?    �z��Ɋi�[���ꂽ�X�ʏ��̃f�[�^��DB�Ɏ�荞��

'*�y�p�����[�^�z
'*P    strErrMsg;        string; �G���[���b�Z�[�W;

'*�y�߂�l�z
'*R    Boolean; TRUE:�����AFALSE:���s;

'******************************************************************'
'*�y�ύX�����z
'*M    2009/06/06; Supattra T.(Tostem Thai); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncModosiKakaku(ByRef strErrMsg) As Boolean
    Dim intN As Integer
    Dim blnRet As Boolean
    Dim strTenCd As String
    Dim intNowLine As Integer
    Dim strCSVRecord() As String
    Dim paryCheckData() As Object  '�u�ύX���v�A�uJAN�R�[�h�v������̍s�P�ʂŃ`�F�b�N���郌�R�[�h
    Dim flgChkError As Boolean      '�`�F�b�N���̃G���[�t���O�iTRUE:�G���[����j
    
    FncModosiKakaku = False
    
    '�S�Ă̔z��f�[�^�̃`�F�b�N���s���B
    flgChkError = False
    intNowLine = 0
    strErrMsg = ""
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(�Y�����ƃR�[�h�`�F�b�N)"
    
    Dim strSQL          As String
    Dim objCDBAccess    As Object
        
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    '�ꊇ�o�^���A�������[�U�[�̂ݎ��s�ł���
    Call LockTableIkkatu(objCDBAccess)
    
    Call SetIkkatuLog("1-1", "�X�ʏ��̃f�[�^��DB�Ɏ�荞��", "FncModosiKakaku", "", "", pstrTantousyaId, "INS")
    Call SetIkkatuLog("1-1-1", "�Y�����ƃR�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")
    
    '�Y�����ƃR�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    'SQL���쐬�i�����F���ʋ敪=5�A�X����NULLor�ꊇ�o�^���ȍ~�AKEY�敪=W02�A�ݒ���e��4���ڂ�1�A�X�R�[�h�j
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
    
    
    
    '�Y�����ƃR�[�h�`�F�b�N�p�@�ؒf���R�[�h�Z�b�g�擾
        pTChkJigyouRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkJigyouRS.Fields("jigyousyo_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
'�W�J�X�R�[�h���X�g�̐���
    
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


        Call SetIkkatuLog("1-1-2", "�����ύX���R�`�F�b�N", "SQL", "", "", pstrTantousyaId, "INS")
    '�����ύX���R�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(�����ύX���R�`�F�b�N)"
    
   'SQL���쐬�i�����FKEY�敪=025�A�����ύX���R�R�[�h�j
    strSQL = "SELECT key_cd FROM v_meisyou WITH(readcommitted) "
    strSQL = strSQL & "WHERE key_kbn = '" & KEYCD_BAIKARIYUU & "'"

        pTChkBaikaHenkouRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkBaikaHenkouRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
        Call SetIkkatuLog("1-1-2", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-1-3", "�����`�F�b�N", "SQL", "", "", pstrTantousyaId, "INS")
    '�����`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(�����`�F�b�N)"
    
    strSQL = "SELECT torihikisaki_cd FROM v_torihikisaki WITH(readcommitted)"
    strSQL = strSQL & "WHERE torikesi_date LIKE ' ' OR torikesi_date IS NULL"
    
        pTChkTorihikisakiRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkTorihikisakiRS.Fields("torihikisaki_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
        Call SetIkkatuLog("1-1-3", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-1-4", "�Z���^�[�敪�`�F�b�N", "SQL", "", "", pstrTantousyaId, "INS")
   '�Z���^�[�敪�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(�Z���^�[�敪�`�F�b�N)"
    
   strSQL = "SELECT key_cd FROM m_wf_meisyou WITH(readcommitted) "
    strSQL = strSQL & "WHERE key_kbn = '" & KEYCD_CENTER_KBN & "' "
    
        pTChkCenterRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkCenterRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
        Call SetIkkatuLog("1-1-4", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-1-5", "��t�ύX���̕ύX���i�����}�X�^�E�����ύX���j", "SQL", "", "", pstrTantousyaId, "INS")
   '��t�ύX���̕ύX���`�F�b�N�p�ؒf���R�[�h�Z�b�g�����i�����}�X�^�E�����ύX���j
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(��t�ύX���̕ύX���i�����}�X�^�E�����ύX���j)"
    'strSQL = "SELECT key_cd, ten_cd, henkou_date "
    strSQL = "SELECT key_cd, henkou_date, ten_cd "
    strSQL = strSQL & "FROM v_baika_henkou WITH(readcommitted) "
    
        pTChkMstBaikaRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkMstBaikaRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    'pTChkMstBaikaRS.Fields("henkou_date").Properties("Optimize") = True
        Call SetIkkatuLog("1-1-5", "", "", "", "OK", "", "UPD")
        Call SetIkkatuLog("1-1-6", "��t�ύX���̕ύX���i�����}�X�^�E�d���ύX���)", "SQL", "", "", pstrTantousyaId, "INS")
   '��t�ύX���̕ύX���`�F�b�N�p�ؒf���R�[�h�Z�b�g�����i�����}�X�^�E�d���ύX���j
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(��t�ύX���̕ύX���i�����}�X�^�E�d���ύX���j)"
    'strSQL = "SELECT key_cd, ten_cd, henkou_date "
    strSQL = "SELECT key_cd, henkou_date, ten_cd "
    strSQL = strSQL & "FROM v_siire_henkou WITH(readcommitted) "

        pTChkMstSiireRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkMstSiireRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    'pTChkMstSiireRS.Fields("henkou_date").Properties("Optimize") = True
        Call SetIkkatuLog("1-1-6", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-1-7", "��t�ύX���̕ύX���i�����}�X�^�E�Z���^�[�R�[�h�ύX���j", "SQL", "", "", pstrTantousyaId, "INS")
   '��t�ύX���̕ύX���`�F�b�N�p�ؒf���R�[�h�Z�b�g�����i�����}�X�^�E�Z���^�[�R�[�h�ύX���j
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(��t�ύX���̕ύX���i�����}�X�^�E�Z���^�[�R�[�h�ύX���j)"
    'strSQL = "SELECT key_cd, ten_cd, henkou_date "
    strSQL = "SELECT key_cd, henkou_date, ten_cd "
    strSQL = strSQL & "FROM v_tenpo_center_cd_henkou WITH(readcommitted) "

        pTChkMstCenterRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pTChkMstCenterRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    'pTChkMstCenterRS.Fields("henkou_date").Properties("Optimize") = True

   'CSV�t�@�C�����ő�A�ŏ����t�̏�����
    pstrMinDate = paryCsvData(intNowLine)(enmTenColNo.HenkouDate)
    pstrMaxDate = paryCsvData(intNowLine)(enmTenColNo.HenkouDate)
    Call SetIkkatuLog("1-1-7", "", "", "", "OK", "", "UPD")


        Call SetIkkatuLog("1-1-8", "��JAN�R�[�h�`�F�b�N", "SQL", "", "", pstrTantousyaId, "INS")
    '�d��JAN�R�[�h�`�F�b�N���R�[�h�Z�b�g�̔r�����b�N�̎��Ԃ�Z������׈ړ�
    '�d��JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(�d��JAN�R�[�h�`�F�b�N)"

    '�d��JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(�d��JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����)"
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
    pChkSameJANCodeRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    pChkSameJANCodeRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
        
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(CSV�t�@�C����Jan�R�[�h�̃��X�g(Jan�R�[�h�̏d������)�쐬)"
    'CSV�t�@�C����Jan�R�[�h�̃��X�g(Jan�R�[�h�̏d������)���쐬����
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

Call SetIkkatuLog("1-1-9", "�����}�X�^�o�^�ς�JAN�R�[�h�`�F�b�N", "SQL", "", "", pstrTantousyaId, "INS")
    '   �����}�X�^�o�^�ς�JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(�����}�X�^�o�^�ς�JAN�R�[�h�`�F�b�N)"
        
       strSQL = "SELECT jan_cd, key_cd  FROM v_key_1 WITH(UPDLOCK)"
        strSQL = strSQL & " WHERE jan_cd IN(" & pCsvJanCodeList & ") "
        strSQL = strSQL & " SET LOCK_TIMEOUT -1"
        
            pChkvk1JANCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
        pChkvk1JANCodeRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U

    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(�G���[�`�F�b�N)"
    
    '********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
    Dim flgFirst As Boolean
    flgFirst = False
    '********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
    Do Until UBound(paryCsvData) < intNowLine
        
        strCSVRecord = paryCsvData(intNowLine)
        
        If FncCheckParameter_CorrectForm_Modosi(strCSVRecord, intNowLine + 1) = False Then
            flgChkError = True
        Else
            If FncCheckParameter_DataBase(strCSVRecord, intNowLine + 1) = False Then
                flgChkError = True
            End If
            '********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
            If FncCheckParameter_PDC_Modosi(strCSVRecord, intNowLine + 1, paryCheckData, flgFirst) = False Then
                flgChkError = True
            End If
            '********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
        End If

        intNowLine = intNowLine + 1
    Loop

    '********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��

    Dim intLoopCnt As Integer
    
   If flgFirst = True Then
    For intLoopCnt = 0 To UBound(paryCheckData, 1)
        If InStr(paryCheckData(intLoopCnt)(4), "'0000'") = 0 Then
            '�G���[�̏ꍇ
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
            '�z��̈��`
            ReDim aryAllLineInfo(0)
            '���[�N�ϐ�
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
                '���[�N�ϐ�
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
Call SetIkkatuLog("1-1-10", "(�����P�ʁA��������v�`�F�b�N)", "SQL", "", "", pstrTantousyaId, "INS")

    intNowLine = 0
        
    Dim strHattyuuTaniHonbu As String
    Dim strGenkaHonbu As String
    Dim strHattyuuTani As String
    Dim strGenka As String

    For intLoopCnt = 0 To UBound(aryAllLineInfo, 1)
    '�u�����v���Ƃ��ɓ��͂��Ȃ��ꍇ�͈ȉ��̃`�F�b�N�͍s��Ȃ��B
    If aryAllLineInfo(intLoopCnt)(2) = "1" Then
    
        strHattyuuTaniHonbu = ""
        strGenkaHonbu = ""
    
        '�ꊇ�o�^�[�ȊO�̓X�܂œX�ʏ��(v_mise_betu)�ɃZ���^�[�敪��2�FPDC�̓X�܂̔����P�ʁA����
        '��
        '�ꊇ�o�^�[�̔����P�ʁA��������v���Ȃ��ꍇ
        '�����P�ʁA��������v�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(�����P�ʁA��������v�`�F�b�N)"
            
        '�����P�ʁA��������v�`�F�b�N�p�ؒf���R�[�h�Z�b�g����(�{��)
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(�����P�ʁA��������v�`�F�b�N�p�ؒf���R�[�h�Z�b�g����(�{��))"
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
                '�{���̔����P�ʁA�����̎擾
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
                
                    '�{���̔����P�ʁA�����̃Z�b�g
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
    
    
        '�����P�ʁA��������v�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(�����P�ʁA��������v�`�F�b�N�p�ؒf���R�[�h�Z�b�g����)"
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
        pChkSameGenkaHattyuuTaniRS.Fields("ten_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
        
        With pChkSameGenkaHattyuuTaniRS
            If Not .BOF Then
                .MoveFirst
                '�X�ʁA���Ə��̏ꍇ
                Do Until .EOF
                    intNowLine = 0
                    Do Until UBound(paryCsvData) < intNowLine
    
                        strCSVRecord = paryCsvData(intNowLine)
                        
                        If strCSVRecord(enmModColNo.HenkouDate) = aryAllLineInfo(intLoopCnt)(0) _
                            And strCSVRecord(enmModColNo.janCD) = aryAllLineInfo(intLoopCnt)(1) Then
                        
                            If Trim(.Fields("ten_cd")) = strCSVRecord(enmModColNo.TenCd) Or Trim(.Fields("jigyou_cd")) = strCSVRecord(enmModColNo.TenCd) Then
                                
                                If strCSVRecord(enmModColNo.CenterKbn) = "" Then
                                    '�Z���^�[���ύX���Ȃ��ꍇ
                                    If strCSVRecord(enmModColNo.Hacchuutani) <> "" Then
                                        .Fields("hattyuu_tani") = strCSVRecord(enmModColNo.Hacchuutani)
                                    End If
                                    If strCSVRecord(enmModColNo.Genka) <> "" Then
                                        .Fields("genka") = strCSVRecord(enmModColNo.Genka)
                                    End If
                                Else
                                    '�Z���^�[��ύX����ꍇ
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
                
                '9999�̏ꍇ
                intNowLine = 0
                Do Until UBound(paryCsvData) < intNowLine
                    
                    strCSVRecord = paryCsvData(intNowLine)
                    
                    If strCSVRecord(enmModColNo.HenkouDate) = aryAllLineInfo(intLoopCnt)(0) _
                        And strCSVRecord(enmModColNo.janCD) = aryAllLineInfo(intLoopCnt)(1) Then
                    
                        If strCSVRecord(enmModColNo.TenCd) = "9999" Then
                            
                            If strCSVRecord(enmModColNo.CenterKbn) = "" Then
                                .MoveFirst
                                Do Until .EOF
                                    '�Z���^�[���ύX���Ȃ��ꍇ
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
                                    '�Z���^�[��ύX����ꍇ
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
                            '�G���[�̏ꍇ
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

    '********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
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
        
        '********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
        If Not pChkSameGenkaHattyuuTaniRS Is Nothing Then
            pChkSameGenkaHattyuuTaniRS.Close
                pChkSameGenkaHattyuuTaniRS = Nothing
        End If
        
        If Not pChkHonbuGenkaHattyuuTaniRS Is Nothing Then
            pChkHonbuGenkaHattyuuTaniRS.Close
                pChkHonbuGenkaHattyuuTaniRS = Nothing
        End If
        '********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
        
        'DB�N���[�Y
        objCDBAccess = Nothing
        'GetObjectContext.SetComplete
   
    '�f�[�^�ɂP�ł���肪����ꍇ�́A�����ŏI���I�I�I
    If flgChkError = True Then
        '�G���[������΃G���[���b�Z�[�W���擾����B�i�I�����Ȃ��j
        strErrMsg = FncGetErrorMsg_Modosi()
        '********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
        If strErrMsg = "" Then
            strErrMsg = FncGetErrorMsg_PDC()
        Else
            strErrMsg = strErrMsg & FncGetErrorMsg_PDC()
        End If
        '********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
        Exit Function
    End If
     
     'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(�f�[�^�o�^)"


    '�����
    If pstrKikakuSyurui = "012" Or pstrKikakuSyurui = "015" Then
    
                Call SetIkkatuLog("1-2", "���i�I��WF(�߂����i)", "InsertModosiKakaku", "", "", pstrTantousyaId, "INS")
        InsertModosiKakaku (paryCsvData)
        Call SetIkkatuLog("1-2", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-3", "���i�I��WF(�߂����i)", "subIns_SetAllTablesMosi", "", "", pstrTantousyaId, "INS")
        subIns_SetAllTablesMosi (paryCsvData)
        Call SetIkkatuLog("1-3", "", "", "", "OK", "", "UPD")
        
    Else
        Call SetIkkatuLog("1-2", "���i�I��WF(�߂����i)", "subIns_SetAllTablesMosi", "", "", pstrTantousyaId, "INS")
        subIns_SetAllTables (paryCsvData)
        Call SetIkkatuLog("1-2", "", "", "", "OK", "", "UPD")
        Call SetIkkatuLog("1-3", "�Ȃ�", "", "", "", pstrTantousyaId, "INS")
        Call SetIkkatuLog("1-3", "", "", "", "OK", "", "UPD")
    End If

    
    '�ꊇ�o�^
    If pstrKikakuSyurui = "015" Then
        Call SetIkkatuLog("1-4", "�ꊇ�o�^", "UpdGenka_HattuTani", "", "", pstrTantousyaId, "INS")
        UpdGenka_HattuTani (paryCsvData)
        Call SetIkkatuLog("1-4", "", "", "", "OK", "", "UPD")
    End If

    '�r�����b�N����
    pChkSameJANCodeRS.Close
        pChkSameJANCodeRS = Nothing
    objCDBAccess = Nothing
    'GetObjectContext.SetComplete
    Call SetIkkatuLog("1-1", "", "", "", "OK", "", "UPD")
    
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncModosiKakaku(�f�[�^�o�^�I��)"

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
    '********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
    If Not pChkSameGenkaHattyuuTaniRS Is Nothing Then
            pChkSameGenkaHattyuuTaniRS = Nothing
    End If
    
    If Not pChkHonbuGenkaHattyuuTaniRS Is Nothing Then
            pChkHonbuGenkaHattyuuTaniRS = Nothing
    End If
    '********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��

    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
    
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    Call SetIkkatuLog("1-1", "", "", Left(Err.Description, 400), "NG", "", "UPD")
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description

End Function
'******************************************************************'
'*�y�v���V�[�W�����z
'*!    InsertModosiKakaku
'*�y�����T�v�z
'*?    INSERT���쐬�i���i�I��WF(�߂����i)�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*�y�߂�l�z
'*    < Name Type Comment >
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2009/06/08; Supattra T(Tostem Thai); �V�K�쐬;
'*M    2010/03/18; �c�ƃV�X�e�� �a�c; ���f�|�Ή�;
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
    
    '���Ə��}�X�^(v_jigyousyo)���A�X�܂��擾����B
    Dim strSQL  As String
    
    '2009.07.02 Add by Supattra T.(Tostem Thai)Start.
    
    Dim strTenkai_kaisi_ym_date As String
            
    '�����DB(���ڍ׏��).�W�J�J�n�N�����̔N�����擾����B
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
    
    'tenkaiDateRS.Fields("tenkai_kaisi_ym_date").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    strTenkai_kaisi_ym_date = tenkaiDateRS.Fields("tenkai_kaisi_ym_date").Value
    
    If Not tenkaiDateRS Is Nothing Then
            tenkaiDateRS = Nothing
    End If
    
    
    '���Ə��}�X�^(v_jigyousyo)���A�X�܂��擾����B
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
    
        'baikaRS.Fields("jigyousyo_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
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

    '�ύX���O�X�܂�o�^�ΏۊO�ɂ���B
    '���i�I��DB(�ύX���O�X��)�����No�Ō�������B
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
    'tenCdRS.Fields("baika_jyogai_kbn").Properties("Optimize") = True '�I�v�e�B�}�C�U
        'tenCdRS.Fields("ten_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    '2009.07.02 Add by Supattra T.(Tostem Thai)End.
    
        Dim nengetuRS As New ADODB.Recordset

    For iRecord = LBound(arryData) To UBound(arryData)

        janCD = arryData(iRecord)(enmTenColNo.janCD)
        strGaitouMiseCD = arryData(iRecord)(enmTenColNo.TenCd)
        
        '�X�܃R�[�h�istrGaitouMiseCD�j���A"9999�F�S�X"�A"01�FVH"�A"03�FVCS"�A"06�FSVH"�̏ꍇ�A�X�܃R�[�h�ɓW�J����B
        '"99�F�S����"�A"07�F���f�|"��ǉ� 2010/03/18
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
        ' �擾�����X�܃R�[�h�����ŌJ��Ԃ��B
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
            
            'nengetuRS.Fields("ten_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
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
    

    '�u�ύX���v���A�܂链���S����.�ύX��(�V�K)���́A�ύX��(�p��)�ɕύX����B
    
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
    pTChkkikakuRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
                
    '�z��̕ύX�����ڂ���L���e�ɂ���āA�ύX����B
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
        '�p�����i�̏ꍇ�A�ύX��(�p��)�ikikaku_mei_hosoku�j�Ŕz��̕ύX�����ڂ�ύX����B
         pCsvhenkouDateKeizoku = marutokuRS.Fields("henkou_date_keizoku").Value
            
        '�V�K���i�̏ꍇ�A�ύX��(�V�K)�i henkou_date_sinki �j�Ŕz��̕ύX�����ڂ�ύX����B
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

    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
            'Err.Number, Err.Source, Err.Description
End Sub
  '********** 2005/05/16 �y�c ���쐬 **********��
Private Sub subIns_SetAllTables(ByVal arryData As Object)
        Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim strMethodCallInfo As String
    Dim lngCount As Long
    Dim aryJanCd() As String
 
    Dim strNVHFilter As String
    Dim rsNVH As ADODB.Recordset
 
'********** 2005/05/18 �y�c �e�X�g�ǉ� **********��
'���ϐ��̒l�����[�J���ɃR�s�[���g�p����

    Dim lstrKikakuSyurui As String
    lstrKikakuSyurui = pstrKikakuSyurui
    
    Dim lstrKikakuNenrenban As String
    lstrKikakuNenrenban = pstrKikakuNenrenban
    
    Dim lstrKikakuEdaban As String
    lstrKikakuEdaban = pstrKikakuEdaban

    Dim lstrTantousyaId As String
    lstrTantousyaId = pstrTantousyaId

'���t���n�߂Ɏ擾
    Dim lstrNow As Object
    lstrNow = Format(Now(), "yyyy/mm/dd hh:nn:ss")
    
'********** 2005/05/18 �y�c �e�X�g�ǉ� **********��

    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables()"

    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    '---------------
    '   ��������
    '---------------
    On Error GoTo DefectErrorHandler

    'RS�쐬�pSQL�̏��������i����ރR�[�h�A�N�A�ԁA�}�Ԃ̎w��j
    Dim strWhereSQL As String
    strWhereSQL = "WHERE kikaku_syurui_cd = '" & lstrKikakuSyurui & "' AND nen_ren_no ='" & lstrKikakuNenrenban & "' AND eda_no ='" & lstrKikakuEdaban & "' "

'********** 2005/05/18 �y�c �ǉ� **********��
     'RS�쐬�pSQL�̏��������i�ύX���t�͈͎̔w��j
    Dim strHenkouSQL As String
    strHenkouSQL = " AND henkou_date BETWEEN '" & pstrMinDate & "' AND '" & pstrMaxDate & "' "
'********** 2005/05/18 �y�c �ǉ� **********��

    'RS�쐬�pSQL
    Dim strSQL  As String

    '�����ύX���RS�쐬
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�����ύX���RS�쐬)"
    strSQL = "SELECT * FROM " & TBL_BAI_HEN & " with (UPDLOCK) " & strWhereSQL
'********** 2005/05/18 �y�c �ǉ� **********��
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") " & strHenkouSQL
'********** 2005/05/18 �y�c �ǉ� **********��

    Dim baikaRS As ADODB.Recordset
        baikaRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
        'baikaRS.Fields("ten_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    '�d���ύX���RS�쐬
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�d���ύX���RS�쐬)"
    strSQL = "SELECT * FROM " & TBL_SII_HEN & " with (UPDLOCK) " & strWhereSQL
'********** 2005/05/18 �y�c �ǉ� **********��
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") " & strHenkouSQL
'********** 2005/05/18 �y�c �ǉ� **********��
    
    Dim siireRS As ADODB.Recordset
        siireRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    siireRS.Fields("ten_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    '�Z���^�[�R�[�h�ύX���RS�쐬
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�Z���^�[�R�[�h�ύX���RS�쐬)"
    strSQL = "SELECT * FROM " & TBL_CEN_HEN & " with (UPDLOCK) " & strWhereSQL
'********** 2005/05/18 �y�c �ǉ� **********��
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") " & strHenkouSQL
'********** 2005/05/18 �y�c �ǉ� **********��
    
    Dim centerRS As ADODB.Recordset
        centerRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    centerRS.Fields("ten_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    'WF�Ǘ�����RS�쐬
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(WF�Ǘ�����RS�쐬)"
    strSQL = "SELECT * FROM t_syousen_wf_kanri with (UPDLOCK) " & strWhereSQL
    
    Dim wfRS As ADODB.Recordset
        wfRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    wfRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
'********** 2005/05/17 �y�c �폜 **********��
'    '��L�[RS����
'    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(��L�[RS����)"
'    strSQL = "SELECT key_cd , jan_cd FROM v_key_1 WITH(readcommitted)"
'
'    Dim vk1RS As ADODB.Recordset
'        vk1RS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
'    vk1RS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
'********** 2005/05/17 �y�c �폜 **********��


'********** 2005/05/17 �y�c �ǉ� **********��
    'VSWF�\���d�l�e�[�u��RS
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(VSWF�\���d�l�e�[�u��RS����)"
    strSQL = _
    "SELECT *" & _
    "  FROM t_syousen_hyouji_siyou WITH(UPDLOCK)" & _
    strWhereSQL
    
    Dim hyoujiSiyouRS As ADODB.Recordset
        hyoujiSiyouRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    hyoujiSiyouRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
         
    'VMST��ʊ�{�e�[�u��RS
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(VMST��ʊ�{�e�[�u��RS����)"
    strSQL = _
    "SELECT v_key_1.jan_cd, v_ippan_kihon.*" & _
    "  FROM v_ippan_kihon WITH(READCOMMITTED)" & _
    " INNER JOIN v_key_1 WITH(READCOMMITTED) ON v_ippan_kihon.key_cd = v_key_1.key_cd" & _
    " AND v_key_1.jan_cd IN(" & pCsvJanCodeList & ") "
    '��jan�R�[�h�ł̍i���ݒǉ�
    
    Dim ippanKihonRS As ADODB.Recordset
        ippanKihonRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    ippanKihonRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
        
    '�Z���^�[�敪�����ɂ��āA�Z���^�[�R�[�h���擾����ׂ�RS
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�Z���^�[�敪�����ɂ��āA�Z���^�[�R�[�h���擾����ׂ�RS����)"
    strSQL = _
    "SELECT jigyousyo_cd, center_cd " & _
    "  FROM v_jigyousyo WITH(READCOMMITTED)"
    
    Dim centerCDRS As ADODB.Recordset
        centerCDRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    centerCDRS.Fields("jigyousyo_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    
'********** 2005/05/17 �y�c �ǉ� **********��

    'WF�X�V�p�t���O(WF�X�V���̌������̈�String��"0"����"1"������)
    Dim strBaikaAdded As String     '�����ύXRS���X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    Dim strSiireAdded As String     '�d���ύXRS���X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    Dim strCenterAdded As String    '�Z���^�[�R�[�h�ύXRS���X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    
    ' ************** 1���R�[�h����RS���X�V *******************
    
    '********** 2008/06/04 ������i��A�j �ǉ� **********��
    Dim strBaikaFlg As String       '�������X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    Dim strSiireFlg As String       '�d�����X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    Dim strCenterFlg As String    '�Z���^�[�R�[�h���X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    '********** 2008/06/04 ������i��A�j �ǉ� **********��
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(1���R�[�h����RS���X�V �J�n)"
    Dim iRecord As Long
'    For iRecord = 0 To UBound(arryData)
    For iRecord = LBound(arryData) To UBound(arryData)
    
        'WF�X�V�p�t���O�̏�����
        strBaikaAdded = "0"
        strSiireAdded = "0"
        strCenterAdded = "0"
        
        '********** 2008/05/27 ������i��A�j �ǉ� **********��
        '�����A�d�����A�Z���^�[�R�[�h���X�V(�V�K���͍X�V)���ꂽ���̃t���O
        strBaikaFlg = "0"
        strSiireFlg = "0"
        strCenterFlg = "0"
        '********** 2008/05/27 ������i��A�j �ǉ� **********��
        '�����A�d���A�Z���^�[�R�[�hRS��JanCD�ƕύX���Ńt�B���^�����O
        Dim strFilter As String
        strFilter = "jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "' AND henkou_date = '" & arryData(iRecord)(enmTenColNo.HenkouDate) & "'"
        
        baikaRS.Filter = strFilter
        siireRS.Filter = strFilter
        centerRS.Filter = strFilter
        
        '�f�[�^�X�V�ΏۓX�R�[�h�̎擾
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�f�[�^�X�V�ΏۓX�R�[�h�̎擾)"
        
        '���̃��R�[�h�̓X�R�[�h
        Dim TenCd As String
        TenCd = arryData(iRecord)(enmTenColNo.TenCd)
        
        '�f�[�^�X�V�ΏۓX�R�[�h�̎擾�ɊY�����ƃR�[�h�`�F�b�N�pRS���ė��p����
        With pTChkJigyouRS
            '�t�B���^�[���������Ă���
            .Filter = adFilterNone
'********** 2005/05/17 �y�c �폜 **********��
'            .MoveFirst
'********** 2005/05/17 �y�c �폜 **********��
            
'********** 2005/05/20 �y�c �ύX **********��
'�{��("0000")�̏ꍇ���t�B���^�����O���Ȃ�(����"�f�[�^�X�V�ΏۓX�R�[�h�̃��[�v"�ɓ��肽���ׁj
'            If tenCD = ALL_TEN_CD Or tenCD = ALL_JIG_CD Then
            If TenCd = ALL_TEN_CD Or TenCd = ALL_JIG_CD Or TenCd = HON_TEN_CD Then
                '�S�X�A�S���ƁA�{���̏ꍇ�Ȃɂ����Ȃ�
'********** 2005/05/20 �y�c �ύX **********��
            ElseIf Len(TenCd) = 2 Then
                '�e���Ƃ̏ꍇ
                '�Y�����ƃR�[�h�`�F�b�N�pRS�����ƃR�[�h�Ńt�B���^�����O����
                
                '***** 2008/09/19 �a�c(IC) �ύX *****��
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
                '***** 2008/09/19 �a�c(IC) �ύX *****��

            ElseIf Len(TenCd) = 4 Then

                '�e�X�̏ꍇ�A�Y�����ƃR�[�h�`�F�b�N�pRS��X�R�[�h�Ńt�B���^�����O����
                .Filter = "jigyousyo_cd = '" & TenCd & "'"
            End If
        End With
        
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�f�[�^�X�V�ΏۓX�R�[�h�̃��[�v �J�n) "
' ************** �f�[�^�X�V�ΏۓX�R�[�h�̃��[�v *******************
'        Dim iTenCD As Integer
'        For iTenCD = LBound(arryTenCD) To UBound(arryTenCD)

        Dim strTenCd As String
        Dim i As Integer
        Do Until pTChkJigyouRS.EOF
           
            
'********** 2005/05/20 �y�c �ύX **********��
'�{��("0000")�̏ꍇ�͓X�R�[�h��"0000"�𒼐ړ����(���R�[�h�Z�b�g����͎擾���Ȃ�)
'            strTenCD = pTChkJigyouRS.Fields("jigyousyo_cd")
            If (TenCd = HON_TEN_CD) Then
                strTenCd = HON_TEN_CD
            Else
                strTenCd = pTChkJigyouRS.Fields("jigyousyo_cd")
            
            End If
'********** 2005/05/20 �y�c �ύX **********��
                        
           '�����A�����ύX���R�̂����ꂩ������ꍇ�A�����ύX���RS�ɒǉ����X�V
            If arryData(iRecord)(enmTenColNo.Baika) <> "" Or arryData(iRecord)(enmTenColNo.BaikaRiyuu) <> "" Then
'                'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�����ύX���RS�ɒǉ����X�V) " & strTenCD
                '�����ύX���RS�ɒǉ����X�V
                With baikaRS
                    
'                    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�����ύX���RS�ɒǉ����X�V) movefirst: " & strTenCD
        '********** 2005/05/18 �y�c �ύX **********��
        '            If Not .EOF Then
                    If Not .BOF Then
        '********** 2005/05/18 �y�c �ύX **********��
                        .MoveFirst
'                        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�����ύX���RS�ɒǉ����X�V) find: " & strTenCD
                        .Find ("ten_cd = '" & strTenCd & "'")
                        
                    End If
                    
                    
                    If .EOF Then
'                        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�V�K�ǉ�)"
                        '�V�K�ǉ�
                        .AddNew
                        .Fields("kikaku_syurui_cd").Value = lstrKikakuSyurui
                        .Fields("nen_ren_no").Value = lstrKikakuNenrenban
                        .Fields("eda_no").Value = lstrKikakuEdaban
                        .Fields("jan_cd").Value = arryData(iRecord)(enmTenColNo.janCD)
'                        .Fields("ten_cd").Value = arryData(iRecord)(enmTenColNo.tenCD)
                        .Fields("ten_cd").Value = strTenCd
                        .Fields("henkou_date").Value = arryData(iRecord)(enmTenColNo.HenkouDate)
                      
                        '����
                        '********** 2008/05/27 ������i��A�j �ǉ� **********��
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�������O�敪��t�B�[���h�͢���O��ɂȂ��
                        'CSV����擾�������������o�^�E�X�V����Ȃ�(NULL�ɃZ�b�g)�B
                                                                                            
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�������O�敪��t�B�[���h�͢���O��ł͂Ȃ����
                        '���i�I��DB�i�����ύX���j�e�[�u����CSV����擾�������������o�^�E�X�V�����B
                       
                        If (strTenCd <> HON_TEN_CD) And _
                            (pTChkJigyouRS.Fields("baika_jyogai_kbn").Value = "1") Then
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�������O�敪��t�B�[���h�͢���O��ɂȂ��
                        'NULL�ɃZ�b�g
                            .Fields("baika").Value = dbNull.value
                        Else
                        
                        '********** 2008/05/27 ������i��A�j �ǉ� **********��
                            If arryData(iRecord)(enmTenColNo.Baika) <> "" Then
                                .Fields("baika").Value = CLng(arryData(iRecord)(enmTenColNo.Baika))
                                strBaikaFlg = "1"
    '********** 2005/05/19 �y�c �ǉ� **********��
                            Else
                                .Fields("baika").Value = dbNull.value
    '********** 2005/05/19 �y�c �ǉ� **********��
                            End If
                        End If
                        
                        
                        '�����ύX���R�R�[�h
'********** 2005/05/19 �y�c �ǉ� **********��
'                        .Fields("baika_henkou_riyuu_cd").Value = arryData(iRecord)(enmTenColNo.BaikaRiyuu)
                        If arryData(iRecord)(enmTenColNo.BaikaRiyuu) <> "" Then
                            .Fields("baika_henkou_riyuu_cd").Value = arryData(iRecord)(enmTenColNo.BaikaRiyuu)
                        Else
                            .Fields("baika_henkou_riyuu_cd").Value = dbNull.value
                        End If
'********** 2005/05/19 �y�c �ǉ� **********��
'********** 2005/06/15 �k�V **********��
'                        .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
'������t�ύX��񂪊DB�Ɋ����̏ꍇ�͓o�^�敪�2����Z�b�g����(���F�����Ɛ�t�ύX�����㏑����f�[�^�ƂȂ�)
                        If InStr(1, arrTourokuKbn2(iRecord, conColBaika), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
'********** 2005/06/15 �k�V **********��
                        .Fields("touroku_user_id").Value = lstrTantousyaId
                        .Fields("touroku_ymd_hms").Value = lstrNow
                        
                    Else
'                        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�X�V����)"
                        ' �X�V����
                        '����
                        '********** 2008/05/27 ������i��A�j �ǉ� **********��
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�������O�敪��t�B�[���h�͢���O��ɂȂ��
                        'CSV����擾�������������o�^�E�X�V����Ȃ�(NULL�ɃZ�b�g)�B
                                                                                            
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�������O�敪��t�B�[���h�͢���O��ł͂Ȃ����
                        '���i�I��DB�i�����ύX���j�e�[�u����CSV����擾�������������o�^�E�X�V�����B
                       
                       
                        If (strTenCd <> HON_TEN_CD) And _
                            (pTChkJigyouRS.Fields("baika_jyogai_kbn").Value = "1") Then
                       
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�������O�敪��t�B�[���h�͢���O��ɂȂ��
                        'NULL�ɃZ�b�g
                            .Fields("baika").Value = dbNull.value
                        Else
                    '********** 2008/05/27 ������i��A�j �ǉ� **********��
                            If arryData(iRecord)(enmTenColNo.Baika) <> "" Then
                                .Fields("baika").Value = CLng(arryData(iRecord)(enmTenColNo.Baika))
                                strBaikaFlg = "1"
                            End If
                        End If
                       
                        '�����ύX���R�R�[�h
                        If arryData(iRecord)(enmTenColNo.BaikaRiyuu) <> "" Then
                            .Fields("baika_henkou_riyuu_cd").Value = arryData(iRecord)(enmTenColNo.BaikaRiyuu)
                        End If
'********** 2005/06/15 �k�V **********��
'������t�ύX��񂪊DB�Ɋ����̏ꍇ�͓o�^�敪�2����Z�b�g����(���F�����Ɛ�t�ύX�����㏑����f�[�^�ƂȂ�)
                        If InStr(1, arrTourokuKbn2(iRecord, conColBaika), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
'********** 2005/06/15 �k�V **********��
                        .Fields("kousin_user_id").Value = lstrTantousyaId
                        .Fields("kousin_ymd_hms").Value = lstrNow
                    End If
                    .Update
                   
                End With
                strBaikaAdded = "1"

            End If
            
            
'            'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�����A�����A�����P�ʂ̂ǂꂩ�����邩�`�F�b�N����B)"
            '�����A�����A�����P�ʂ̂ǂꂩ�����邩�`�F�b�N����B
            If arryData(iRecord)(enmTenColNo.Genka) <> "" _
                Or arryData(iRecord)(enmTenColNo.Torihikisaki) <> "" _
                Or arryData(iRecord)(enmTenColNo.Hacchuutani) <> "" Then
                
'                'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�d���ύX���RS�X�V)"
                '�d���ύX���RS�X�V
                With siireRS
'********** 2005/05/18 �y�c �ύX **********��
'                   If Not .EOF Then
                   If Not .BOF Then
'********** 2005/05/18 �y�c �ύX **********��
                        .MoveFirst
                        .Find ("ten_cd = '" & strTenCd & "'")
                    End If
                    If .EOF Then
                         '�V�K�ǉ�
                        .AddNew
                        .Fields("kikaku_syurui_cd").Value = lstrKikakuSyurui
                        .Fields("nen_ren_no").Value = lstrKikakuNenrenban
                        .Fields("eda_no").Value = lstrKikakuEdaban
                        .Fields("jan_cd").Value = arryData(iRecord)(enmTenColNo.janCD)
'                        .Fields("ten_cd").Value = arryData(iRecord)(enmTenColNo.tenCD)
                        .Fields("ten_cd").Value = strTenCd
                        .Fields("henkou_date").Value = arryData(iRecord)(enmTenColNo.HenkouDate)
                        
                        '�����R�[�h
'********** 2005/05/19 �y�c �ύX **********��
'                        .Fields("torihikisaki_cd").Value = arryData(iRecord)(enmTenColNo.Torihikisaki)
                        '********** 2008/05/28 ������i��A�j �ǉ� **********��
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢����揜�O�敪��͢���O��ł͂Ȃ��ꍇ
                        '�����R�[�h�͓o�^�����B
                       
                        If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("torihikisaki_jyogai_kbn").Value = "1" Then
                       
                            .Fields("torihikisaki_cd").Value = dbNull.value
                            
                        Else
                    '********** 2008/05/28 ������i��A�j �ǉ� **********��
                            If arryData(iRecord)(enmTenColNo.Torihikisaki) <> "" Then
                                .Fields("torihikisaki_cd").Value = arryData(iRecord)(enmTenColNo.Torihikisaki)
                                strSiireFlg = "1"
                            Else
                                .Fields("torihikisaki_cd").Value = dbNull.value
                            End If
                        End If
                       
                        
                    
                        
'********** 2005/05/19 �y�c �ύX **********��
                        '�����P��
                         '********** 2008/05/28 ������i��A�j �ǉ� **********��
                         '���i�I��WF(�ύX���O�X��)�e�[�u���̢�����P�ʏ��O�敪��͢���O��ł͂Ȃ��ꍇ
                         '�����P�ʂ͓o�^�����B
                 
                        If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("hattyuu_tani_jyogai_kbn").Value = "1" Then
                       
                            .Fields("hattyuu_tani").Value = dbNull.value
                            
                        Else
                            '********** 2008/05/28 ������i��A�j �ǉ� **********��
                            If arryData(iRecord)(enmTenColNo.Hacchuutani) <> "" Then
                                .Fields("hattyuu_tani").Value = CLng(arryData(iRecord)(enmTenColNo.Hacchuutani))
     '********** 2005/05/19 �y�c �ǉ� **********��
                                strSiireFlg = "1"
                            Else
                                .Fields("hattyuu_tani").Value = dbNull.value
    '********** 2005/05/19 �y�c �ǉ� **********��
                            End If
                         
                        End If
                        
                        '����
                         '********** 2008/05/28 ������i��A�j �ǉ� **********��
                         '���i�I��WF(�ύX���O�X��)�e�[�u���̢�������O�敪��͢���O��ł͂Ȃ��ꍇ
                         '�����͓o�^�����B
                         
                        If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("genka_jyogai_kbn").Value = "1" Then
                            .Fields("genka").Value = dbNull.value
                        Else
                        '********** 2008/05/28 ������i��A�j �ǉ� **********��
                            If arryData(iRecord)(enmTenColNo.Genka) <> "" Then
    '********** 2005/05/19 �y�c �ύX **********��
    '                            .Fields("genka").Value = CLng(arryData(iRecord)(enmTenColNo.Genka))
                                .Fields("genka").Value = CDbl(arryData(iRecord)(enmTenColNo.Genka))
    '********** 2005/05/19 �y�c �ǉ� **********��
                                strSiireFlg = "1"
                            Else
                                .Fields("genka").Value = dbNull.value
    '********** 2005/05/19 �y�c �ǉ� **********��
    '********** 2005/05/19 �y�c �ύX **********��
                            End If
                        End If
                       
'********** 2005/06/15 �k�V **********��
'                        .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
'�d����t�ύX��񂪊DB�Ɋ����̏ꍇ�͓o�^�敪�2����Z�b�g����(���F�����Ɛ�t�ύX�����㏑����f�[�^�ƂȂ�)
                        If InStr(1, arrTourokuKbn2(iRecord, conColSiire), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
'********** 2005/06/15 �k�V **********��
                        .Fields("touroku_user_id").Value = lstrTantousyaId
                        .Fields("touroku_ymd_hms").Value = lstrNow
                    Else
                        '�X�V����
                        '�����R�[�h
                        '********** 2008/05/28 ������i��A�j �ǉ� **********��
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢����揜�O�敪��͢���O��ł͂Ȃ��ꍇ
                        '�����R�[�h�͍X�V�����B
                        If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("torihikisaki_jyogai_kbn").Value = "1" Then
                    
                            .Fields("torihikisaki_cd").Value = dbNull.value
                        Else
                        
                            If arryData(iRecord)(enmTenColNo.Torihikisaki) <> "" Then
    '********** 2005/05/19 �y�c �ύX **********��
    '                            .Fields("torihikisaki_cd").Value = CLng(arryData(iRecord)(enmTenColNo.Torihikisaki))
                                 .Fields("torihikisaki_cd").Value = arryData(iRecord)(enmTenColNo.Torihikisaki)
    '********** 2005/05/19 �y�c �ύX **********��
                                 strSiireFlg = "1"
                            End If
                        End If
                        
                        '********** 2008/05/28 ������i��A�j �ǉ� **********��
                        '�����P��
                        '********** 2008/05/28 ������i��A�j �ǉ� **********��
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�����P�ʏ��O�敪��͢���O��ł͂Ȃ��ꍇ
                        '�����P�ʂ͍X�V�����B
                        If (strTenCd <> HON_TEN_CD) And _
                                pTChkJigyouRS.Fields("hattyuu_tani_jyogai_kbn").Value = "1" Then
                            .Fields("hattyuu_tani").Value = dbNull.value
                        Else
                            If arryData(iRecord)(enmTenColNo.Hacchuutani) <> "" Then
                                .Fields("hattyuu_tani").Value = arryData(iRecord)(enmTenColNo.Hacchuutani)
                                strSiireFlg = "1"
                            End If
                        End If
                        
                        
                        '********** 2008/05/28 ������i��A�j �ǉ� **********��
                        
                        '����
                        '********** 2008/05/28 ������i��A�j �ǉ� **********��
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�������O�敪��͢���O��ł͂Ȃ��ꍇ
                        '�����͍X�V�����B
                         If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("genka_jyogai_kbn").Value = "1" Then
                            .Fields("genka").Value = dbNull.value
                                
                        Else
                            
                            If arryData(iRecord)(enmTenColNo.Genka) <> "" Then
    '********** 2005/05/19 �y�c �ύX **********��
    '                            .Fields("genka").Value = arryData(iRecord)(enmTenColNo.Genka)
                                .Fields("genka").Value = CDbl(arryData(iRecord)(enmTenColNo.Genka))
    '********** 2005/05/19 �y�c �ύX **********��
                                strSiireFlg = "1"
                            End If
                        End If
                        
                        
'********** 2005/06/15 �k�V **********��
'�d����t�ύX��񂪊DB�Ɋ����̏ꍇ�͓o�^�敪�2����Z�b�g����(���F�����Ɛ�t�ύX�����㏑����f�[�^�ƂȂ�)
                        If InStr(1, arrTourokuKbn2(iRecord, conColSiire), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
'********** 2005/06/15 �k�V **********��
                        .Fields("kousin_user_id").Value = lstrTantousyaId
                        .Fields("kousin_ymd_hms").Value = lstrNow
                   End If
                End With
                strSiireAdded = "1"
            End If
        
'            'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�Z���^�[�敪�����邩�`�F�b�N����)"
            '�Z���^�[�敪�����邩�`�F�b�N����B
            If arryData(iRecord)(enmTenColNo.CenterKbn) <> "" Then
'                'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�Z���^�[�R�[�h�ύX���RS�X�V)"
                '�Z���^�[�R�[�h�ύX���RS�X�V
                
                '�Z���^�[�R�[�h�����߂�F
                '�Z���^�[�敪��0�̏ꍇ��NULL�A1�̏ꍇ��'0151'�A2�̏ꍇ�͎��Ə��}�X�^����擾����
'********** 2005/05/19 �y�c �ύX **********��
'   �Z���^�[�敪���O�̏ꍇ��""�ł͂Ȃ�NULL������悤�C��
'                Dim strCenterCD As String
                Dim centerCD
                Select Case arryData(iRecord)(enmTenColNo.CenterKbn)
                    Case "0":       '�Z���^�[�敪��0�̏ꍇ�ANull���Z�b�g
'                        strCenterCD = ""
                        centerCD = dbNull.value
                    
                    '2015/09/15 wadak7 upd
                    'Case "1":       '�Z���^�[�敪��1�̏ꍇ�A0151(���ʃZ���^�[)���Z�b�g
                    Case "1", "3":      '�Z���^�[�敪��1,3�̏ꍇ
                        'centerCD = CENTERCD_KODAMA
                        '�o�^�[�̃Z���^�[�R�[�h���Z�b�g
                        centerCD = arryData(iRecord)(enmModColNo.centerCD)
                    Case "2":       '�Z���^�[�敪��2�̏ꍇ�A�X�R�[�h�����Ɏ��Ə��}�X�^����擾�����Z���^�[�R�[�h���Z�b�g
                        With centerCDRS
'********** 2005/05/19 �y�c �ύX **********��
'                            If Not .EOF Then
                            If Not .BOF Then
'********** 2005/05/19 �y�c �ύX **********��
                                .MoveFirst
                                .Find ("jigyousyo_cd = '" & strTenCd & "'")
                            End If
                        
                            '�f�[�^������������Null���Z�b�g����B
                            If .EOF Then
'                                strCenterCD = ""
                                centerCD = dbNull.value
                            Else
                                centerCD = centerCDRS.Fields("center_cd")
                            End If
                        End With
                        
                    Case Else:       '�Z���^�[�敪��0,1,2�ȊO�̏ꍇ�ANull���Z�b�g
'                       strCenterCD = ""
                        centerCD = dbNull.value
'********** 2005/05/19 �y�c �ύX **********��
                        
                End Select
                
                With centerRS
'********** 2005/05/18 �y�c �ύX **********��
'                    If Not .EOF Then
                    If Not .BOF Then
'********** 2005/05/18 �y�c �ύX **********��
                        .MoveFirst
                        .Find ("ten_cd = '" & strTenCd & "'")
                    End If
                    If .EOF Then
                         '�V�K�ǉ�
                        .AddNew
                        .Fields("kikaku_syurui_cd").Value = lstrKikakuSyurui
                        .Fields("nen_ren_no").Value = lstrKikakuNenrenban
                        .Fields("eda_no").Value = lstrKikakuEdaban
                        .Fields("jan_cd").Value = arryData(iRecord)(enmTenColNo.janCD)
'                        .Fields("ten_cd").Value = arryData(iRecord)(enmTenColNo.tenCD)
                        .Fields("ten_cd").Value = strTenCd
                        .Fields("henkou_date").Value = arryData(iRecord)(enmTenColNo.HenkouDate)
                      
                        '�Z���^�[�敪
                        '********** 2008/05/28 ������i��A�j �ǉ� **********��
                
               
                        If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("center_jyogai_kbn").Value = "1" Then
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�Z���^�[���O�敪��͢���O��ł���΁ANULL�ɃZ�b�g
                            .Fields("center_kbn").Value = dbNull.value
                        Else
                    '********** 2008/05/28 ������i��A�j �ǉ� **********��
                            .Fields("center_kbn").Value = arryData(iRecord)(enmTenColNo.CenterKbn)
                            '�Z���^�[�R�[�h�F�Z���^�[�敪��0�̏ꍇ��NULL�A1�̏ꍇ��'0151'�A2�̏ꍇ�͎��Ə��}�X�^����擾����
    '                        .Fields("center_cd").Value = FncGetCenterCd(arryData(iRecord)(enmTenColNo.CenterKbn), strTenCD)
    '********** 2005/05/18 �y�c �ύX **********��
    '                        .Fields("center_cd").Value = strCenterCD
                            .Fields("center_cd").Value = centerCD
    '********** 2005/05/18 �y�c �ύX **********��
                            strCenterFlg = "1"
                        End If
                       
'********** 2005/06/15 �k�V **********��
'                        .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
'�Z���^�[��t�ύX��񂪊DB�Ɋ����̏ꍇ�͓o�^�敪�2����Z�b�g����(���F�����Ɛ�t�ύX�����㏑����f�[�^�ƂȂ�)
                        If InStr(1, arrTourokuKbn2(iRecord, conColCenter), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
'********** 2005/06/15 �k�V **********��
                        .Fields("touroku_user_id").Value = lstrTantousyaId
                        .Fields("touroku_ymd_hms").Value = lstrNow
                    Else
                        ' �X�V����
                        '�Z���^�[�敪
                        '********** 2008/05/28 ������i��A�j �ǉ� **********��
                        If (strTenCd <> HON_TEN_CD) And _
                                pTChkJigyouRS.Fields("center_jyogai_kbn").Value = "1" Then
                            '���i�I��WF(�ύX���O�X��)�e�[�u���̢�Z���^�[���O�敪��͢���O��ł���΁ANULL�ɃZ�b�g
                                .Fields("center_kbn").Value = dbNull.value
                        Else
                    '********** 2008/05/28 ������i��A�j �ǉ� **********��
                            If arryData(iRecord)(enmTenColNo.CenterKbn) <> "" Then
                                .Fields("center_kbn").Value = arryData(iRecord)(enmTenColNo.CenterKbn)
                                '�Z���^�[�R�[�h�F�Z���^�[�敪��0�̏ꍇ��NULL�A1�̏ꍇ��'0151'�A2�̏ꍇ�͎��Ə��}�X�^����擾����
        '                        .Fields("center_cd").Value = FncGetCenterCd(arryData(iRecord)(enmTenColNo.CenterKbn), strTenCD)
    '********** 2005/05/18 �y�c �ύX **********��
        '                        .Fields("center_cd").Value = strCenterCD
                                .Fields("center_cd").Value = centerCD
    '********** 2005/05/18 �y�c �ύX **********��
                                strCenterFlg = "1"
                            End If
                        End If
                        
                        
'********** 2005/06/15 �k�V **********��
'�Z���^�[��t�ύX��񂪊DB�Ɋ����̏ꍇ�͓o�^�敪�2����Z�b�g����(���F�����Ɛ�t�ύX�����㏑����f�[�^�ƂȂ�)
                        If InStr(1, arrTourokuKbn2(iRecord, conColCenter), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
'********** 2005/06/15 �k�V **********��
                        .Fields("kousin_user_id").Value = lstrTantousyaId
                        .Fields("kousin_ymd_hms").Value = lstrNow
                    End If
                End With
                strCenterAdded = "1"
           End If
'********** 2005/05/20 �y�c �ǉ� **********��
'�{��("0000")�̏ꍇ�͈���Do���[�v�𔲂���
            If (strTenCd = HON_TEN_CD) Then Exit Do
'********** 2005/05/20 �y�c �ǉ� **********��
            pTChkJigyouRS.MoveNext
            
        Loop
            
        '�����A�d���A�Z���^�[�R�[�hRS�̂����ꂩ�Œǉ��������������ꍇ
        If (strBaikaAdded = "1" Or strCenterAdded = "1" Or strSiireAdded = "1") Then
'            'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�����A�d���A�Z���^�[�R�[�hRS�̂����ꂩ�Œǉ��������������ꍇ)"


            'JAN�R�[�h�ƕR�t����L�[���擾����
            Dim strKey_cd
'********** 2005/05/17 �y�c �ύX **********��
'            With vk1RS
            With pChkvk1JANCodeRS
'********** 2005/05/18 �y�c �ύX **********��
'                If Not .EOF Then
                If Not .BOF Then
'********** 2005/05/18 �y�c �ύX **********��
                    .MoveFirst
                    .Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
                    strKey_cd = .Fields("key_cd")
                End If
'********** 2005/05/17 �y�c �ύX **********��
            End With

            
            'WF�Ǘ����ڂ��ŐV���ɍX�V����
            With wfRS
'********** 2005/05/17 �y�c �ύX **********��
'********** 2005/05/18 �y�c �ύX **********��
'                If Not .EOF Then
                If Not .BOF Then
'********** 2005/05/18 �y�c �ύX **********��
                    .MoveFirst
                    .Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
                End If
'********** 2005/05/17 �y�c �ύX **********��
                'WF�Ǘ����ڂɖ����ꍇ�́A�V�K�ǉ�
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
                    '********** 2008/05/27 ������i��A�j �폜 **********��
                    '�����ύX���o�^�敪
                    '.Fields("baika_henkou_touroku_kbn").Value = strBaikaAdded
                    '�d���ύX���o�^�敪
                    '.Fields("siire_henkou_touroku_kbn").Value = strSiireAdded
                    '�������ޕύX���o�^�敪
                    '.Fields("center_cd_henkou_touroku_kbn").Value = strCenterAdded
                    '********** 2008/05/27 ������i��A�j �폜 **********��
                    
                    '�����ύX���o�^�敪
                    '********** 2008/05/27 ������i��A�j �ǉ� **********��
                    If strBaikaFlg = "1" Then
                        .Fields("baika_henkou_touroku_kbn").Value = strBaikaAdded
                    Else
                        .Fields("baika_henkou_touroku_kbn").Value = 0
                    End If
                    '********** 2008/05/27 ������i��A�j �ǉ� **********��
                    
                    '�d���ύX���o�^�敪
                    '********** 2008/05/27 ������i��A�j �ǉ� **********��
                    If strSiireFlg = "1" Then
                        .Fields("siire_henkou_touroku_kbn").Value = strSiireAdded
                    Else
                        .Fields("siire_henkou_touroku_kbn").Value = 0
                    End If
                    
                    '********** 2008/05/27 ������i��A�j �ǉ� **********��
                    
                    '�������ޕύX���o�^�敪
                    '********** 2008/05/27 ������i��A�j �ǉ� **********��
                    If strCenterFlg = "1" Then
                        .Fields("center_cd_henkou_touroku_kbn").Value = strCenterAdded
                    Else
                        .Fields("center_cd_henkou_touroku_kbn").Value = 0
                    End If
                    
                    '********** 2008/05/27 ������i��A�j �ǉ� **********��
                    .Fields("kowake_henkou_touroku_kbn").Value = 0
                    .Fields("hanki_henkou_touroku_kbn").Value = 0
                    .Fields("touroku_user_id").Value = lstrTantousyaId
                    .Fields("kousin_user_id").Value = dbNull.value
                    .Fields("touroku_ymd_hms").Value = lstrNow
                    .Fields("kousin_ymd_hms").Value = dbNull.value
        
                'WF�Ǘ����ڂɂ���ꍇ�́A�X�V
                Else
                    .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                    .Fields("touroku_tantousya").Value = lstrTantousyaId
                    .Fields("kikaku_no").Value = lstrKikakuSyurui & lstrKikakuNenrenban & "-" & lstrKikakuEdaban
                    .Fields("naibu_cd").Value = strKey_cd
                    .Fields("hyouji_siyou_touroku_kbn").Value = 1
'********** 2005/06/16 �k�V **********��
'                    .Fields("baika_henkou_touroku_kbn").Value = strBaikaAdded
'                    .Fields("siire_henkou_touroku_kbn").Value = strSiireAdded
'                    .Fields("center_cd_henkou_touroku_kbn").Value = strCenterAdded
'�X�V���́A�o�^�敪�0����Z�b�g���鎖�͂��蓾�Ȃ�
                    '********** 2008/05/27 ������i��A�j �폜 **********��
                    'If strBaikaAdded = "1" Then
                        '�����ύX���o�^�敪
                    '    .Fields("baika_henkou_touroku_kbn").Value = strBaikaAdded
                    'End If
                    'If strSiireAdded = "1" Then
                        '�d���ύX���o�^�敪
                    '    .Fields("siire_henkou_touroku_kbn").Value = strSiireAdded
                    'End If
                    'If strCenterAdded = "1" Then
                        '�������ޕύX���o�^�敪
                    '    .Fields("center_cd_henkou_touroku_kbn").Value = strCenterAdded
                    'End If
                    '********** 2008/05/27 ������i��A�j �폜 **********��
                    
                    '********** 2008/05/27 ������i��A�j �ǉ� **********��
                    If strBaikaFlg = "1" Then
                        If strBaikaAdded = "1" Then
                            '�����ύX���o�^�敪
                            .Fields("baika_henkou_touroku_kbn").Value = strBaikaAdded
                        End If
                    'Else
                    '    .Fields("baika_henkou_touroku_kbn").Value = dbNull.value
                        
                    End If
                    If strSiireFlg = "1" Then
                        If strSiireAdded = "1" Then
                            '�d���ύX���o�^�敪
                            .Fields("siire_henkou_touroku_kbn").Value = strSiireAdded
                        End If
                    'Else
                    '    .Fields("siire_henkou_touroku_kbn").Value = dbNull.value
                    End If
                    If strCenterFlg = "1" Then
                        If strCenterAdded = "1" Then
                            '�������ޕύX���o�^�敪
                            .Fields("center_cd_henkou_touroku_kbn").Value = strCenterAdded
                        End If
                    'Else
                    '    .Fields("center_cd_henkou_touroku_kbn").Value = dbNull.value
                    End If
                    '********** 2008/05/27 ������i��A�j �ǉ� **********��
'********** 2005/06/16 �k�V **********��
                    .Fields("kousin_user_id").Value = lstrTantousyaId
                    .Fields("kousin_ymd_hms").Value = lstrNow
                End If
                .Update
            End With
        End If
        
        '�eRS�̃t�B���^�[����
'        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�eRS�̃t�B���^�[����)"
            baikaRS .Filter = adFilterNone
        siireRS.Filter = adFilterNone
        centerRS.Filter = adFilterNone
        
'********** 2005/05/17 �y�c �ǉ� **********��
'VMST��ʊ�{�e�[�u����菤�i�����擾���AVSWF�\���d�l�e�[�u���Ƀ��R�[�h��ǉ�����
'        'strMethodCallInfo = CStr(iRecord) & "�s�ځFVMST��ʊ�{�e�[�u����菤�i�����擾���AVSWF�\���d�l�e�[�u���Ƀ��R�[�h��ǉ�����"
        With hyoujiSiyouRS
'********** 2005/05/18 �y�c �ύX **********��
'            If Not .EOF Then
            If Not .BOF Then
'********** 2005/05/18 �y�c �ύX **********��
                .MoveFirst
                .Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
            End If
'            'strMethodCallInfo = CStr(iRecord) & "�s�ځFhyoujiSiyouRS"
            
            'VSWF�\���d�l�e�[�u���Ƀ��R�[�h�����݂��Ȃ��ꍇ��VMST��ʊ�{�e�[�u����菤�i�����擾��VSWF�\���d�l�e�[�u���Ƀ��R�[�h��ǉ�����
            If .EOF Then
'                'strMethodCallInfo = CStr(iRecord) & "�s�ځFVMST��ʊ�{�e�[�u����菤�i�����擾��"
                With ippanKihonRS
'********** 2005/05/18 �y�c �ύX **********��
'                    If Not .EOF Then
                    If Not .BOF Then
'********** 2005/05/18 �y�c �ύX **********��
                        .MoveFirst
                        .Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
                    End If
'                    'strMethodCallInfo = CStr(iRecord) & "�s�ځFjan_cd�łe�h�m�c"
                    If .EOF Then
'                        'strMethodCallInfo = CStr(iRecord) & "�s�ځFVMST��ʊ�{�e�[�u���Ƀf�[�^�����݂��Ȃ���VSWF�\���d�l�e�[�u���Ƀ��R�[�h��ǉ��o���܂���ł����B"
                        GoTo DefectErrorHandler
                    Else
                        'VSWF�\���d�l�e�[�u��RS���X�V����
'                        'strMethodCallInfo = CStr(iRecord) & "�s�ځFVSWF�\���d�l�e�[�u��RS���X�V����"
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
'********** 2005/05/17 �y�c �ǉ� **********��
    Next iRecord

    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(RS�̃o�b�`�X�V0)"
    'RS�̃o�b�`�X�V
    Call objCDBAccess.UpdateDCRecordset("VSWF", baikaRS)
    Call objCDBAccess.UpdateDCRecordset("VSWF", siireRS)
    Call objCDBAccess.UpdateDCRecordset("VSWF", centerRS)
    Call objCDBAccess.UpdateDCRecordset("VSWF", wfRS)
'********** 2005/05/17 �y�c �ǉ� **********��
    Call objCDBAccess.UpdateDCRecordset("VSWF", hyoujiSiyouRS)
'********** 2005/05/17 �y�c �ǉ� **********��
    
    'GetObjectContext.SetComplete
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�I�u�W�F�N�g�J��)"
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

'********** 2005/05/17 �y�c �폜 **********��
'    If Not vk1RS Is Nothing Then
'            vk1RS = Nothing
'    End If
'********** 2005/05/17 �y�c �폜 **********��

'********** 2005/05/17 �y�c �ǉ� **********��
    If Not hyoujiSiyouRS Is Nothing Then
            hyoujiSiyouRS = Nothing
    End If

    If Not ippanKihonRS Is Nothing Then
            ippanKihonRS = Nothing
    End If
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�o�^�����p�����̂ł����ŊJ������)"
'�o�^�����p�����̂ł����ŊJ������(�쐬��FncTenTouroku�ōs���Ă���)
    If Not pTChkJigyouRS Is Nothing Then
        pTChkJigyouRS.Close
            pTChkJigyouRS = Nothing
    End If
    
'�o�^�����p�����̂ł����ŊJ������(�쐬��FncTenTouroku�ōs���Ă���)
    If Not pChkvk1JANCodeRS Is Nothing Then
        pChkvk1JANCodeRS.Close
            pChkvk1JANCodeRS = Nothing
    End If
    
    If Not centerCDRS Is Nothing Then
        centerCDRS.Close
            centerCDRS = Nothing
    End If
    
'********** 2005/05/17 �y�c �ǉ� **********��
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
    
'********** 2005/05/17 �y�c �폜 **********��
'    If Not vk1RS Is Nothing Then
'            vk1RS = Nothing
'    End If
'********** 2005/05/17 �y�c �폜 **********��

'********** 2005/05/17 �y�c �ǉ� **********��
    If Not hyoujiSiyouRS Is Nothing Then
            hyoujiSiyouRS = Nothing
    End If

    If Not ippanKihonRS Is Nothing Then
            ippanKihonRS = Nothing
    End If

'�o�^�����p�����̂ł����ŊJ������(�쐬���̂�FncTenTouroku�ōs���Ă���)
    If Not pTChkJigyouRS Is Nothing Then
       pTChkJigyouRS.Close
            pTChkJigyouRS = Nothing
    End If

'�o�^�����p�����̂ł����ŊJ������(�쐬��FncTenTouroku�ōs���Ă���)
    If Not pChkvk1JANCodeRS Is Nothing Then
        pChkvk1JANCodeRS.Close
            pChkvk1JANCodeRS = Nothing
    End If
    
    If Not centerCDRS Is Nothing Then
        centerCDRS.Close
            centerCDRS = Nothing
    End If
    
'********** 2005/05/17 �y�c �ǉ� **********��
    
    'If Not GetObjectContext Is Nothing Then
        'GetObjectContext.SetAbort
        'End If

    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
            'Err.Number, Err.Source, Err.Description
End Sub
'********** 2005/05/16 �y�c ���쐬 **********��

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    subIns_SetAllTablesMosi
'*�y�����T�v�z
'*?    INSERT���쐬�i���i�I��WF(�߂����i)�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*�y�߂�l�z
'*    < Name Type Comment >
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2009/06/08; Supattra T(Tostem Thai); �V�K�쐬;
'*M    2011/02/16; �c�ƃV�X�e��/�a�c; �Z���^�[�R�[�h�ǉ�;
'*M    2015/09/15; ���ʃV�X�e��G/�a�c; TC�V�X�e���\�z;
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

'���t���n�߂Ɏ擾
    Dim lstrNow As Object
    lstrNow = Format(Now(), "yyyy/mm/dd hh:nn:ss")

    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi()"

    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    '---------------
    '   ��������
    '---------------
    On Error GoTo DefectErrorHandler

    'RS�쐬�pSQL�̏��������i����ރR�[�h�A�N�A�ԁA�}�Ԃ̎w��j
    Dim strWhereSQL As String
    strWhereSQL = "WHERE kikaku_syurui_cd = '" & lstrKikakuSyurui & "' AND nen_ren_no ='" & lstrKikakuNenrenban & "' AND eda_no ='" & lstrKikakuEdaban & "' "

     'RS�쐬�pSQL�̏��������i�ύX���t�͈͎̔w��j
    Dim strHenkouSQL As String
    If pCsvHenkouDate <> "" And pCsvhenkouDateKeizoku <> "" Then
        strHenkouSQL = " AND henkou_date BETWEEN '" & pCsvHenkouDate & "' AND '" & pCsvhenkouDateKeizoku & "' "
    Else
        strHenkouSQL = " AND henkou_date BETWEEN '" & pstrMinDate & "' AND '" & pstrMaxDate & "' "
    End If

    'RS�쐬�pSQL
    Dim strSQL  As String

    '�����ύX���RS�쐬
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(�����ύX���RS�쐬)"
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
    baikaRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    'baikaRS.Fields("ten_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    '�d���ύX���RS�쐬
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(�d���ύX���RS�쐬)"
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
    siireRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    'siireRS.Fields("ten_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    '�Z���^�[�R�[�h�ύX���RS�쐬
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(�Z���^�[�R�[�h�ύX���RS�쐬)"
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
    centerRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    'centerRS.Fields("ten_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    'WF�Ǘ�����RS�쐬
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(WF�Ǘ�����RS�쐬)"
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
    wfRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    'VSWF�\���d�l�e�[�u��RS
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(VSWF�\���d�l�e�[�u��RS����)"
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
    hyoujiSiyouRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
         
    'VMST��ʊ�{�e�[�u��RS
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(VMST��ʊ�{�e�[�u��RS����)"
    strSQL = _
    "SELECT v_key_1.jan_cd, v_ippan_kihon.*" & _
    "  FROM v_ippan_kihon WITH(READCOMMITTED)" & _
    " INNER JOIN v_key_1 WITH(READCOMMITTED) ON v_ippan_kihon.key_cd = v_key_1.key_cd" & _
    " AND v_key_1.jan_cd IN(" & pCsvJanCodeList & ") "
    '��jan�R�[�h�ł̍i���ݒǉ�
    
    Dim ippanKihonRS As ADODB.Recordset
        ippanKihonRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    ippanKihonRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
        
    '�Z���^�[�敪�����ɂ��āA�Z���^�[�R�[�h���擾����ׂ�RS
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(�Z���^�[�敪�����ɂ��āA�Z���^�[�R�[�h���擾����ׂ�RS����)"
    strSQL = _
    "SELECT jigyousyo_cd, center_cd " & _
    "  FROM v_jigyousyo WITH(READCOMMITTED)"
    
    Dim centerCDRS As ADODB.Recordset
        centerCDRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    centerCDRS.Fields("jigyousyo_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U

    'WF�X�V�p�t���O(WF�X�V���̌������̈�String��"0"����"1"������)
    Dim strBaikaAdded As String     '�����ύXRS���X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    Dim strSiireAdded As String     '�d���ύXRS���X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    Dim strCenterAdded As String    '�Z���^�[�R�[�h�ύXRS���X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    
    ' ************** 1���R�[�h����RS���X�V *******************
    
    Dim strBaikaFlg As String       '�������X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    Dim strSiireFlg As String       '�d�����X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    Dim strCenterFlg As String    '�Z���^�[�R�[�h���X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(1���R�[�h����RS���X�V �J�n)"
    Dim iRecord As Long
    
    For iRecord = LBound(arryData) To UBound(arryData)
    
    
        'WF�X�V�p�t���O�̏�����
        strBaikaAdded = "0"
        strSiireAdded = "0"
        strCenterAdded = "0"
        
        '�����A�d�����A�Z���^�[�R�[�h���X�V(�V�K���͍X�V)���ꂽ���̃t���O
        strBaikaFlg = "0"
        strSiireFlg = "0"
        strCenterFlg = "0"
        
         '2009.06.25 Mod by Supattra T.(Tostem Thai)Start.

        '�����A�d���A�Z���^�[�R�[�hRS��JanCD�ƕύX���Ńt�B���^�����O
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
        
        '�f�[�^�X�V�ΏۓX�R�[�h�̎擾
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(�f�[�^�X�V�ΏۓX�R�[�h�̎擾)"
        
        '���̃��R�[�h�̓X�R�[�h
        Dim TenCd As String
        TenCd = arryData(iRecord)(enmTenColNo.TenCd)
        
        '�f�[�^�X�V�ΏۓX�R�[�h�̎擾�ɊY�����ƃR�[�h�`�F�b�N�pRS���ė��p����
        With pTChkJigyouRS
            '�t�B���^�[���������Ă���
            .Filter = adFilterNone

            If TenCd = ALL_TEN_CD Or TenCd = ALL_JIG_CD Or TenCd = HON_TEN_CD Then
                '�S�X�A�S���ƁA�{���̏ꍇ�Ȃɂ����Ȃ�
            ElseIf Len(TenCd) = 2 Then
                '�e���Ƃ̏ꍇ
                '�Y�����ƃR�[�h�`�F�b�N�pRS�����ƃR�[�h�Ńt�B���^�����O����

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

                '�e�X�̏ꍇ�A�Y�����ƃR�[�h�`�F�b�N�pRS��X�R�[�h�Ńt�B���^�����O����
                .Filter = "jigyousyo_cd = '" & TenCd & "'"
            End If
        End With
        
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(�f�[�^�X�V�ΏۓX�R�[�h�̃��[�v �J�n) "
' ************** �f�[�^�X�V�ΏۓX�R�[�h�̃��[�v *******************

        Dim strTenCd As String
        Dim i As Integer
        Dim strCaseHenkouDate As String
        
        Do Until pTChkJigyouRS.EOF
           
'�{��("0000")�̏ꍇ�͓X�R�[�h��"0000"�𒼐ړ����(���R�[�h�Z�b�g����͎擾���Ȃ�)
            If (TenCd = HON_TEN_CD) Then
                strTenCd = HON_TEN_CD
            Else
                strTenCd = pTChkJigyouRS.Fields("jigyousyo_cd")
            
            End If
                        
           '�����A�����ύX���R�̂����ꂩ������ꍇ�A�����ύX���RS�ɒǉ����X�V
            If arryData(iRecord)(enmTenColNo.Baika) <> "" Or arryData(iRecord)(enmTenColNo.BaikaRiyuu) <> "" Then
                '�����ύX���RS�ɒǉ����X�V
                With baikaRS
                    If Not .BOF Then
                        .MoveFirst
                        .Find ("ten_cd = '" & Trim(strTenCd) & "'")
                    End If
                    
                    If .EOF Then
                        '�V�K�ǉ�
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
                        
                        
                        '����
                        '********** 2008/05/27 ������i��A�j �ǉ� **********��
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�������O�敪��t�B�[���h�͢���O��ɂȂ��
                        'CSV����擾�������������o�^�E�X�V����Ȃ�(NULL�ɃZ�b�g)�B
                                                                                            
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�������O�敪��t�B�[���h�͢���O��ł͂Ȃ����
                        '���i�I��DB�i�����ύX���j�e�[�u����CSV����擾�������������o�^�E�X�V�����B
                       
                        If (strTenCd <> HON_TEN_CD) And _
                            (pTChkJigyouRS.Fields("baika_jyogai_kbn").Value = "1") Then
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�������O�敪��t�B�[���h�͢���O��ɂȂ��
                        'NULL�ɃZ�b�g
                            .Fields("baika").Value = dbNull.value
                        Else
                        
                            If arryData(iRecord)(enmTenColNo.Baika) <> "" Then
                                .Fields("baika").Value = CLng(arryData(iRecord)(enmTenColNo.Baika))
                                strBaikaFlg = "1"
                            Else
                                .Fields("baika").Value = dbNull.value
                            End If
                        End If
                        
                        
                        '�����ύX���R�R�[�h
                        If arryData(iRecord)(enmTenColNo.BaikaRiyuu) <> "" Then
                            .Fields("baika_henkou_riyuu_cd").Value = arryData(iRecord)(enmTenColNo.BaikaRiyuu)
                        Else
                            .Fields("baika_henkou_riyuu_cd").Value = dbNull.value
                        End If
'������t�ύX��񂪊DB�Ɋ����̏ꍇ�͓o�^�敪�2����Z�b�g����(���F�����Ɛ�t�ύX�����㏑����f�[�^�ƂȂ�)
                        If InStr(1, arrTourokuKbn2(iRecord, conColBaika), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
                        .Fields("touroku_user_id").Value = lstrTantousyaId
                        .Fields("touroku_ymd_hms").Value = lstrNow
                        
                    Else
'                        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables(�X�V����)"
                        ' �X�V����
                        '����
                        '********** 2008/05/27 ������i��A�j �ǉ� **********��
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�������O�敪��t�B�[���h�͢���O��ɂȂ��
                        'CSV����擾�������������o�^�E�X�V����Ȃ�(NULL�ɃZ�b�g)�B
                                                                                            
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�������O�敪��t�B�[���h�͢���O��ł͂Ȃ����
                        '���i�I��DB�i�����ύX���j�e�[�u����CSV����擾�������������o�^�E�X�V�����B
                       
                       
                        If (strTenCd <> HON_TEN_CD) And _
                            (pTChkJigyouRS.Fields("baika_jyogai_kbn").Value = "1") Then
                       
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�������O�敪��t�B�[���h�͢���O��ɂȂ��
                        'NULL�ɃZ�b�g
                            .Fields("baika").Value = dbNull.value
                        Else
                            If arryData(iRecord)(enmTenColNo.Baika) <> "" Then
                                .Fields("baika").Value = CLng(arryData(iRecord)(enmTenColNo.Baika))
                                strBaikaFlg = "1"
                            End If
                        End If
                       
                        '�����ύX���R�R�[�h
                        If arryData(iRecord)(enmTenColNo.BaikaRiyuu) <> "" Then
                            .Fields("baika_henkou_riyuu_cd").Value = arryData(iRecord)(enmTenColNo.BaikaRiyuu)
                        End If
'������t�ύX��񂪊DB�Ɋ����̏ꍇ�͓o�^�敪�2����Z�b�g����(���F�����Ɛ�t�ύX�����㏑����f�[�^�ƂȂ�)
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
            
            
            '�����A�����A�����P�ʂ̂ǂꂩ�����邩�`�F�b�N����B
            If arryData(iRecord)(enmTenColNo.Genka) <> "" _
                Or arryData(iRecord)(enmTenColNo.Torihikisaki) <> "" _
                Or arryData(iRecord)(enmTenColNo.Hacchuutani) <> "" Then
                
                '�d���ύX���RS�X�V
                With siireRS
'                   If Not .EOF Then
                    If Not .BOF Then
                        .MoveFirst
                        .Find ("ten_cd = '" & strTenCd & "'")
                    End If
                    
                    If .EOF Then
                         '�V�K�ǉ�
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
                        
                        '�����R�[�h
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢����揜�O�敪��͢���O��ł͂Ȃ��ꍇ
                        '�����R�[�h�͓o�^�����B
                       
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
                       
                        '�����P��
                         '���i�I��WF(�ύX���O�X��)�e�[�u���̢�����P�ʏ��O�敪��͢���O��ł͂Ȃ��ꍇ
                         '�����P�ʂ͓o�^�����B
                 
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
                        
                        '����
                         '���i�I��WF(�ύX���O�X��)�e�[�u���̢�������O�敪��͢���O��ł͂Ȃ��ꍇ
                         '�����͓o�^�����B
                         
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
                       
'�d����t�ύX��񂪊DB�Ɋ����̏ꍇ�͓o�^�敪�2����Z�b�g����(���F�����Ɛ�t�ύX�����㏑����f�[�^�ƂȂ�)
                        If InStr(1, arrTourokuKbn2(iRecord, conColSiire), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If
                        .Fields("touroku_user_id").Value = lstrTantousyaId
                        .Fields("touroku_ymd_hms").Value = lstrNow
                    Else
                        '�X�V����
                        '�����R�[�h
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢����揜�O�敪��͢���O��ł͂Ȃ��ꍇ
                        '�����R�[�h�͍X�V�����B
                        If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("torihikisaki_jyogai_kbn").Value = "1" Then
                    
                            .Fields("torihikisaki_cd").Value = dbNull.value
                        Else
                        
                            If arryData(iRecord)(enmTenColNo.Torihikisaki) <> "" Then
                                 .Fields("torihikisaki_cd").Value = arryData(iRecord)(enmTenColNo.Torihikisaki)
                                 strSiireFlg = "1"
                            End If
                        End If
                        
                        '�����P��
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�����P�ʏ��O�敪��͢���O��ł͂Ȃ��ꍇ
                        '�����P�ʂ͍X�V�����B
                        If (strTenCd <> HON_TEN_CD) And _
                                pTChkJigyouRS.Fields("hattyuu_tani_jyogai_kbn").Value = "1" Then
                            .Fields("hattyuu_tani").Value = dbNull.value
                        Else
                            If arryData(iRecord)(enmTenColNo.Hacchuutani) <> "" Then
                                .Fields("hattyuu_tani").Value = arryData(iRecord)(enmTenColNo.Hacchuutani)
                                strSiireFlg = "1"
                            End If
                        End If
                        
                        '����
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�������O�敪��͢���O��ł͂Ȃ��ꍇ
                        '�����͍X�V�����B
                         If (strTenCd <> HON_TEN_CD) And _
                            pTChkJigyouRS.Fields("genka_jyogai_kbn").Value = "1" Then
                            .Fields("genka").Value = dbNull.value
                                
                        Else
                            
                            If arryData(iRecord)(enmTenColNo.Genka) <> "" Then
                                .Fields("genka").Value = CDbl(arryData(iRecord)(enmTenColNo.Genka))
                                strSiireFlg = "1"
                            End If
                        End If
                        
'�d����t�ύX��񂪊DB�Ɋ����̏ꍇ�͓o�^�敪�2����Z�b�g����(���F�����Ɛ�t�ύX�����㏑����f�[�^�ƂȂ�)
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
        
            '�Z���^�[�敪�����邩�`�F�b�N����B
            If arryData(iRecord)(enmTenColNo.CenterKbn) <> "" Then
                '�Z���^�[�R�[�h�ύX���RS�X�V
                '�Z���^�[�R�[�h�����߂�F
                '�Z���^�[�敪��0�̏ꍇ��NULL�A1�̏ꍇ��'0151'�A2�̏ꍇ�͎��Ə��}�X�^����擾����
'   �Z���^�[�敪���O�̏ꍇ��""�ł͂Ȃ�NULL������悤�C��
                Dim centerCD
                Select Case arryData(iRecord)(enmTenColNo.CenterKbn)
                    Case "0":       '�Z���^�[�敪��0�̏ꍇ�ANull���Z�b�g
'                        strCenterCD = ""
                        centerCD = dbNull.value
                    
                    '2015/09/15 wadak7 upd
                    'Case "1":       '�Z���^�[�敪��1�̏ꍇ�A0151(���ʃZ���^�[)���Z�b�g
                    Case "1", "3":      '�Z���^�[�敪��1,3�̏ꍇ
                        'centerCD = CENTERCD_KODAMA
                        '�o�^�[�̃Z���^�[�R�[�h���Z�b�g
                        centerCD = arryData(iRecord)(enmModColNo.centerCD)
                    
                    Case "2":       '�Z���^�[�敪��2�̏ꍇ�A�X�R�[�h�����Ɏ��Ə��}�X�^����擾�����Z���^�[�R�[�h���Z�b�g
                        With centerCDRS
                            If Not .BOF Then
                                .MoveFirst
                                .Find ("jigyousyo_cd = '" & strTenCd & "'")
                            End If
                        
                            '�f�[�^������������Null���Z�b�g����B
                            If .EOF Then
                                centerCD = dbNull.value
                            Else
                                centerCD = centerCDRS.Fields("center_cd")
                            End If
                        End With
                        
                    Case Else:       '�Z���^�[�敪��0,1,2�ȊO�̏ꍇ�ANull���Z�b�g
                        centerCD = dbNull.value
                        
                End Select
                
                With centerRS
                    If Not .BOF Then
                        .MoveFirst
                        .Find ("ten_cd = '" & strTenCd & "'")
                    End If
                    
                    If .EOF Then
                         '�V�K�ǉ�
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
                        '���i�I��WF(�ύX���O�X��)�e�[�u���̢�Z���^�[���O�敪��͢���O��ł���΁ANULL�ɃZ�b�g
                            .Fields("center_kbn").Value = dbNull.value
                        Else
                            .Fields("center_kbn").Value = arryData(iRecord)(enmTenColNo.CenterKbn)
                            '�Z���^�[�R�[�h�F�Z���^�[�敪��0�̏ꍇ��NULL�A1�̏ꍇ��'0151'�A2�̏ꍇ�͎��Ə��}�X�^����擾����
                            .Fields("center_cd").Value = centerCD
                            strCenterFlg = "1"
                        End If
                       
'�Z���^�[��t�ύX��񂪊DB�Ɋ����̏ꍇ�͓o�^�敪�2����Z�b�g����(���F�����Ɛ�t�ύX�����㏑����f�[�^�ƂȂ�)
                        If InStr(1, arrTourokuKbn2(iRecord, conColCenter), strTenCd, vbTextCompare) > 0 Then
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                        Else
                            .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SINKI
                        End If

                        .Fields("touroku_user_id").Value = lstrTantousyaId
                        .Fields("touroku_ymd_hms").Value = lstrNow
                    Else
                        ' �X�V����
                        '�Z���^�[�敪
                        If (strTenCd <> HON_TEN_CD) And _
                                pTChkJigyouRS.Fields("center_jyogai_kbn").Value = "1" Then
                            '���i�I��WF(�ύX���O�X��)�e�[�u���̢�Z���^�[���O�敪��͢���O��ł���΁ANULL�ɃZ�b�g
                                .Fields("center_kbn").Value = dbNull.value
                        Else
                            If arryData(iRecord)(enmTenColNo.CenterKbn) <> "" Then
                                .Fields("center_kbn").Value = arryData(iRecord)(enmTenColNo.CenterKbn)
                                '�Z���^�[�R�[�h�F�Z���^�[�敪��0�̏ꍇ��NULL�A1�̏ꍇ��'0151'�A2�̏ꍇ�͎��Ə��}�X�^����擾����
                                .Fields("center_cd").Value = centerCD
                                strCenterFlg = "1"
                            End If
                        End If
                        
'�Z���^�[��t�ύX��񂪊DB�Ɋ����̏ꍇ�͓o�^�敪�2����Z�b�g����(���F�����Ɛ�t�ύX�����㏑����f�[�^�ƂȂ�)
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
           
'�{��("0000")�̏ꍇ�͈���Do���[�v�𔲂���
            If (strTenCd = HON_TEN_CD) Then Exit Do
            pTChkJigyouRS.MoveNext
            
        Loop
            
        '�����A�d���A�Z���^�[�R�[�hRS�̂����ꂩ�Œǉ��������������ꍇ
        If (strBaikaAdded = "1" Or strCenterAdded = "1" Or strSiireAdded = "1") Then


            'JAN�R�[�h�ƕR�t����L�[���擾����
            Dim strKey_cd

'            With vk1RS
            With pChkvk1JANCodeRS
                If Not .BOF Then
                    .MoveFirst
                    .Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
                    strKey_cd = .Fields("key_cd")
                End If
            End With

            
            'WF�Ǘ����ڂ��ŐV���ɍX�V����
            With wfRS
                If Not .BOF Then
                    .MoveFirst
                    .Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
                End If
                'WF�Ǘ����ڂɖ����ꍇ�́A�V�K�ǉ�
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
        
                'WF�Ǘ����ڂɂ���ꍇ�́A�X�V
                Else
                    .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                    .Fields("touroku_tantousya").Value = lstrTantousyaId
                    .Fields("kikaku_no").Value = lstrKikakuSyurui & lstrKikakuNenrenban & "-" & lstrKikakuEdaban
                    .Fields("naibu_cd").Value = strKey_cd
                    .Fields("hyouji_siyou_touroku_kbn").Value = 1

                    If strBaikaFlg = "1" Then
                        If strBaikaAdded = "1" Then
                            '�����ύX���o�^�敪
                            .Fields("baika_henkou_touroku_kbn").Value = strBaikaAdded
                        End If
                        
                    End If
                    If strSiireFlg = "1" Then
                        If strSiireAdded = "1" Then
                            '�d���ύX���o�^�敪
                            .Fields("siire_henkou_touroku_kbn").Value = strSiireAdded
                        End If
                    End If
                    If strCenterFlg = "1" Then
                        If strCenterAdded = "1" Then
                            '�������ޕύX���o�^�敪
                            .Fields("center_cd_henkou_touroku_kbn").Value = strCenterAdded
                        End If
                    End If

                    .Fields("kousin_user_id").Value = lstrTantousyaId
                    .Fields("kousin_ymd_hms").Value = lstrNow
                End If
                .Update
            End With
        End If
        
        '�eRS�̃t�B���^�[����
        baikaRS.Filter = adFilterNone
        siireRS.Filter = adFilterNone
        centerRS.Filter = adFilterNone
        pTChkkikakuRS.Filter = adFilterNone
        
'VMST��ʊ�{�e�[�u����菤�i�����擾���AVSWF�\���d�l�e�[�u���Ƀ��R�[�h��ǉ�����
        With hyoujiSiyouRS
            If Not .BOF Then
                .MoveFirst
                .Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
            End If
            
            'VSWF�\���d�l�e�[�u���Ƀ��R�[�h�����݂��Ȃ��ꍇ��VMST��ʊ�{�e�[�u����菤�i�����擾��VSWF�\���d�l�e�[�u���Ƀ��R�[�h��ǉ�����
            If .EOF Then
                With ippanKihonRS
                    If Not .BOF Then
                        .MoveFirst
                        .Find ("jan_cd = '" & arryData(iRecord)(enmTenColNo.janCD) & "'")
                    End If
                    If .EOF Then
                        GoTo DefectErrorHandler
                    Else
                        'VSWF�\���d�l�e�[�u��RS���X�V����
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

    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(RS�̃o�b�`�X�V0)"
    'RS�̃o�b�`�X�V
    Call objCDBAccess.UpdateDCRecordset("VSWF", baikaRS)
    Call objCDBAccess.UpdateDCRecordset("VSWF", siireRS)
    Call objCDBAccess.UpdateDCRecordset("VSWF", centerRS)
    Call objCDBAccess.UpdateDCRecordset("VSWF", wfRS)
    Call objCDBAccess.UpdateDCRecordset("VSWF", hyoujiSiyouRS)
    
    'GetObjectContext.SetComplete
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(�I�u�W�F�N�g�J��)"
    
    
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
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTablesMosi(�o�^�����p�����̂ł����ŊJ������)"
'�o�^�����p�����̂ł����ŊJ������(�쐬��FncTenTouroku�ōs���Ă���)
    If Not pTChkJigyouRS Is Nothing Then
        pTChkJigyouRS.Close
            pTChkJigyouRS = Nothing
    End If
    If Not pTChkkikakuRS Is Nothing Then
        pTChkkikakuRS.Close
            pTChkkikakuRS = Nothing
    End If
    
'�o�^�����p�����̂ł����ŊJ������(�쐬��FncTenTouroku�ōs���Ă���)
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

'�o�^�����p�����̂ł����ŊJ������(�쐬���̂�FncTenTouroku�ōs���Ă���)
    If Not pTChkJigyouRS Is Nothing Then
       pTChkJigyouRS.Close
            pTChkJigyouRS = Nothing
    End If
    
    If Not pTChkkikakuRS Is Nothing Then
        pTChkkikakuRS.Close
            pTChkkikakuRS = Nothing
    End If

'�o�^�����p�����̂ł����ŊJ������(�쐬��FncTenTouroku�ōs���Ă���)
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

    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
            'Err.Number, Err.Source, Err.Description
End Sub


'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckParameter_CorrectForm
'*�y�����T�v�z
'*?    CSV�t�@�C���̃p�����[�^�ɑ΂��āADB�Ɉˑ����Ȃ��`���`�F�b�N���s���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckParameter_CorrectForm; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
'
Private Function FncCheckParameter_CorrectForm(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheckParameter_CorrectForm = False
    
    Dim blnIsOKRequired
    Dim blnIsOKNumeric
    Dim blnIsOKLimit
    Dim blnIsOKHenkoubi
    Dim blnIsOKCheckDisit
'********** 2005/05/07 �k�V **********��
    Dim blnIsOKNotZero
'********** 2005/05/07 �k�V **********��
    
    'a.�K�{���ځi�Y���X�R�[�h�A�ύX���AJAN�R�[�h�j�������Ă��邩�m�F����B
    blnIsOKRequired = FncCheck_RequiredParameter(arryCSVRecord(), intNowLine)

    'b.�S�Ă̍��ڂɂ����āA���p�����ł��邩�m�F����B
    blnIsOKNumeric = FncCheck_IsParameterNumeric(arryCSVRecord(), intNowLine)
    
    'c.�����������𒴂��Ă��Ȃ����m�F����B
    blnIsOKLimit = FncCheck_Limit(arryCSVRecord(), intNowLine)

    'd.�ύX���̓��t���L���ł��邩�m�F����B
    blnIsOKHenkoubi = FncCheck_Henkoubi(arryCSVRecord(enmTenColNo.HenkouDate), intNowLine)
    
    'e.JAN�R�[�h�ɐ������`�F�b�N�f�W�b�g�̒l���Z�b�g����Ă��邩�A
'********** 2005/04/26 ���� �ύX **********��
    blnIsOKCheckDisit = FncCheck_CheckDigit(arryCSVRecord(enmTenColNo.janCD), intNowLine, 1)
'    blnIsOKCheckDisit = FncCheck_CheckDigit(arryCSVRecord(enmTenColNo.JanCd), intNowLine)
'********** 2005/04/26 ���� �ύX **********��
    
'********** 2005/05/07 �k�V **********��
'���z�ނ����̒l�����`�F�b�N����
'********** 2005/06/16 �k�V **********��
'�����ύX���R�́A�X=0000������<>""�̏ꍇ�͓��͕K�{�A����ȊO�̏ꍇ�͓��͕s�� �̃`�F�b�N��ǉ�
'********** 2005/06/16 �k�V **********��
    blnIsOKNotZero = FncCheck_NotZero(arryCSVRecord(), intNowLine, "TEN")
'********** 2005/05/07 �k�V **********��
    
'********** 2005/06/16 �k�V **********��
'    If blnIsOKRequired And blnIsOKNumeric And blnIsOKLimit And blnIsOKHenkoubi And blnIsOKCheckDisit Then
    If blnIsOKRequired And blnIsOKNumeric And blnIsOKLimit And _
       blnIsOKHenkoubi And blnIsOKCheckDisit And blnIsOKNotZero Then
'********** 2005/06/16 �k�V **********��
        FncCheckParameter_CorrectForm = True
    End If
    
End Function
'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckParameter_CorrectForm_Modosi
'*�y�����T�v�z
'*?    CSV�t�@�C���̃p�����[�^�ɑ΂��āADB�Ɉˑ����Ȃ��`���`�F�b�N���s���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckParameter_CorrectForm_Modosi; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2009/06/06; Supattra T.(Tostem Thai); �V�K�쐬;
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
    
    'a.�K�{���ځi�Y���X�R�[�h�A�ύX���AJAN�R�[�h�j�������Ă��邩�m�F����B
    blnIsOKRequired = FncCheck_RequiredParameter_Modosi(arryCSVRecord(), intNowLine)

    'b.�S�Ă̍��ڂɂ����āA���p�����ł��邩�m�F����B
    blnIsOKNumeric = FncCheck_IsParameterNumeric_Modosi(arryCSVRecord(), intNowLine)
    
    'c.�����������𒴂��Ă��Ȃ����m�F����B
    blnIsOKLimit = FncCheck_Limit_Modosi(arryCSVRecord(), intNowLine)

    'd.�ύX���̓��t���L���ł��邩�m�F����B
    blnIsOKHenkoubi = FncCheck_Henkoubi(arryCSVRecord(enmTenColNo.HenkouDate), intNowLine)
    
    'e.JAN�R�[�h�ɐ������`�F�b�N�f�W�b�g�̒l���Z�b�g����Ă��邩�A
    blnIsOKCheckDisit = FncCheck_CheckDigit(arryCSVRecord(enmTenColNo.janCD), intNowLine, 1)
    
    '�����ύX���R�́A�X=0000������<>""�̏ꍇ�͓��͕K�{�A����ȊO�̏ꍇ�͓��͕s�� �̃`�F�b�N��ǉ�
    blnIsOKNotZero = FncCheck_NotZero(arryCSVRecord(), intNowLine, "TEN")
    
    If blnIsOKRequired And blnIsOKNumeric And blnIsOKLimit And _
       blnIsOKHenkoubi And blnIsOKCheckDisit And blnIsOKNotZero Then
        FncCheckParameter_CorrectForm_Modosi = True
    End If
    
End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckParameter_PDC_Modosi
'*�y�����T�v�z
'*?    CSV�t�@�C���̃p�����[�^�ɑ΂��āA�u�ύX���v�A�uJAN�R�[�h�v������̍s�P�ʂŃ`�F�b�N����B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckParameter_PDC_Modosi; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2012/08/10; �k�o(��A); �V�K�쐬;
'*< Comment End >**************************************************'
'
Private Function FncCheckParameter_PDC_Modosi(ByRef arryCSVRecord() As String, intNowLine As Integer, _
                                        ByRef arryCheckRecord As Object, ByRef flgFirst As Boolean) As Boolean

    FncCheckParameter_PDC_Modosi = True
    
    '�ȉ��́u�ύX���v�A�uJAN�R�[�h�v������̍s�P�ʂŃ`�F�b�N����B
    Dim aryLineInfo As Object
    Dim intLoopCnt As Integer
    
    If arryCSVRecord(enmModColNo.CenterKbn) = "2" Or arryCSVRecord(enmModColNo.TenCd) = "0000" Then
        '�z��̗L������
        If Not flgFirst Then
            '�z��̈��`
            ReDim arryCheckRecord(0)
            '���[�N�ϐ�
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
                       '�G���[�̏ꍇ
                        FncSetErrorPDC intNowLine, enmModColNo.Genka, PARAM_GENKA, arryCSVRecord(enmModColNo.HenkouDate) & "," & arryCSVRecord(enmModColNo.janCD)
                        FncCheckParameter_PDC_Modosi = False
                    End If
                    If arryCheckRecord(intLoopCnt)(3) <> arryCSVRecord(enmModColNo.Hacchuutani) Then
                        If arryCheckRecord(intLoopCnt)(3) = "" Then
                            arryCheckRecord(intLoopCnt)(3) = arryCSVRecord(enmModColNo.Hacchuutani)
                        End If
                       '�G���[�̏ꍇ
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
            '���[�N�ϐ�
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
'*�y�v���V�[�W�����z
'*!    FncCheck_RequiredParameter
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�K�{���ڂ��󔒂łȂ����ǂ����`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_RequiredParameter; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'

Private Function FncCheck_RequiredParameter(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_RequiredParameter = True
    
    If arryCSVRecord(enmTenColNo.TenCd) = "" Then       '�Y���X�R�[�h
        FncSetError (intNowLine, enmTenColNo.TenCd, PARAM_EMPTY
        FncCheck_RequiredParameter = False
    End If
    If arryCSVRecord(enmTenColNo.HenkouDate) = "" Then  '�ύX��
        FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_EMPTY
        FncCheck_RequiredParameter = False
    End If
    If arryCSVRecord(enmTenColNo.janCD) = "" Then       'JAN�R�[�h
        FncSetError (intNowLine, enmTenColNo.janCD, PARAM_EMPTY
        FncCheck_RequiredParameter = False
    End If
    

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_RequiredParameter_Modosi
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�K�{���ڂ��󔒂łȂ����ǂ����`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_RequiredParameter_Modosi; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*M    2011/02/10; �c�ƃV�X�e��/�a�c;         �Z���^�[�R�[�h�ǉ�;
'*M    2015/09/15; ���ʃV�X�e��G/�a�c;        TC�V�X�e���\�z;
'*< Comment End >**************************************************'

Private Function FncCheck_RequiredParameter_Modosi(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_RequiredParameter_Modosi = True
    
    If arryCSVRecord(enmModColNo.TenCd) = "" Then       '�Y���X�R�[�h
        FncSetError (intNowLine, enmModColNo.TenCd, PARAM_EMPTY
        FncCheck_RequiredParameter_Modosi = False
    End If
    If arryCSVRecord(enmModColNo.HenkouDate) = "" Then  '�ύX��
        FncSetError (intNowLine, enmModColNo.HenkouDate, PARAM_EMPTY
        FncCheck_RequiredParameter_Modosi = False
    End If
    If arryCSVRecord(enmModColNo.janCD) = "" Then       'JAN�R�[�h
        FncSetError (intNowLine, enmModColNo.janCD, PARAM_EMPTY
        FncCheck_RequiredParameter_Modosi = False
    End If
    '2015/09/15 wadak7 upd
    'If arryCSVRecord(enmModColNo.CenterKbn) = "1" Then       '�Z���^�[�敪��1�F���~�Z���^�[�̏ꍇ�A�K�{����
    '�Z���^�[�敪��1�F���~�Z���^�[�E3�FTC�Z���^�[�̏ꍇ�A�K�{����
    If arryCSVRecord(enmModColNo.CenterKbn) = "1" Or arryCSVRecord(enmModColNo.CenterKbn) = "3" Then
        If arryCSVRecord(enmModColNo.centerCD) = "" Then       '�Z���^�[�R�[�h
            FncSetError (intNowLine, enmModColNo.centerCD, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
    Else    '�Z���^�[�敪��1�F���~�Z���^�[�E3�FTC�Z���^�[�ȊO�̏ꍇ�A���͂ł��Ȃ�
        If arryCSVRecord(enmModColNo.centerCD) <> "" Then       '�Z���^�[�R�[�h
            FncSetError (intNowLine, enmModColNo.centerCD, PARAM_KINSI
            FncCheck_RequiredParameter_Modosi = False
        End If
    End If
    
    '********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
    If arryCSVRecord(enmModColNo.CenterKbn) = "2" Then       '�u�Z���^�[�敪�v��"2"�FPDC�̏ꍇ
        If arryCSVRecord(enmModColNo.Hacchuutani) = "" Then          '�u�����P�ʁv�������͂̏ꍇ
            FncSetError (intNowLine, enmModColNo.Hacchuutani, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
        If arryCSVRecord(enmModColNo.Genka) = "" Then       '�u�����v�������͂̏ꍇ
            FncSetError (intNowLine, enmModColNo.Genka, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
    End If
    '********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
    
    'letKikakuSyurui
    If pstrKikakuSyurui = "012" Then
        If arryCSVRecord(enmModColNo.MoGenka) = "" Then  '�߂�����
            FncSetError (intNowLine, enmModColNo.MoGenka, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
        If arryCSVRecord(enmModColNo.MoBaika) = "" Then  '�߂�����
            FncSetError (intNowLine, enmModColNo.MoBaika, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
        If arryCSVRecord(enmModColNo.HattyuuTani) = "" Then  '�߂������P��
            FncSetError (intNowLine, enmModColNo.HattyuuTani, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
        If arryCSVRecord(enmModColNo.PopSize) = "" Then  'POP�T�C�Y
            FncSetError (intNowLine, enmModColNo.PopSize, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
    ElseIf pstrKikakuSyurui = "015" Then
         If arryCSVRecord(enmModColNo.MoGenka) = "" Then  '�߂�����
            FncSetError (intNowLine, enmModColNo.MoGenka, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
        If arryCSVRecord(enmModColNo.HattyuuTani) = "" Then  '�߂������P��
            FncSetError (intNowLine, enmModColNo.HattyuuTani, PARAM_EMPTY
            FncCheck_RequiredParameter_Modosi = False
        End If
    End If

End Function
'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_IsParameterNumeric
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�S�t�B�[���h�����p���l���ǂ����`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_IsParameterNumeric; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterNumeric(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_IsParameterNumeric = True
    
    If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.TenCd)) = False Then     '�Y���X�R�[�h
        FncSetError (intNowLine, enmTenColNo.TenCd, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.HenkouDate)) = False Then '�ύX��
        FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.janCD)) = False Then     'JAN�R�[�h
        FncSetError (intNowLine, enmTenColNo.janCD, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric = False
    End If
    
    If arryCSVRecord(enmTenColNo.Genka) <> "" Then
        If IsNumeric(arryCSVRecord(enmTenColNo.Genka)) = False Then   '����
'        If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.Genka)) = False Then '����
            FncSetError (intNowLine, enmTenColNo.Genka, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric = False
        End If
    End If
    
    If arryCSVRecord(enmTenColNo.Baika) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.Baika)) = False Then   '����
            FncSetError (intNowLine, enmTenColNo.Baika, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric = False
        End If
    End If
    
    If arryCSVRecord(enmTenColNo.BaikaRiyuu) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.BaikaRiyuu)) = False Then  '�����ύX���R
            FncSetError (intNowLine, enmTenColNo.BaikaRiyuu, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric = False
        End If
    End If
    
    If arryCSVRecord(enmTenColNo.Torihikisaki) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.Torihikisaki)) = False Then '�����R�[�h
            FncSetError (intNowLine, enmTenColNo.Torihikisaki, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric = False
        End If
    End If
    
    If arryCSVRecord(enmTenColNo.Hacchuutani) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.Hacchuutani)) = False Then '�����P��
            FncSetError (intNowLine, enmTenColNo.Hacchuutani, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric = False
        End If
    End If

    If arryCSVRecord(enmTenColNo.CenterKbn) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmTenColNo.CenterKbn)) = False Then '�Z���^�[�敪
            FncSetError (intNowLine, enmTenColNo.CenterKbn, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric = False
        End If
    End If

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_IsParameterNumeric_Modosi
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�S�t�B�[���h�����p���l���ǂ����`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_IsParameterNumeric_Modosi; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2009/06/06; Supattra T.(Tostem Thai); �V�K�쐬;
'*M    2011/02/10; �c�ƃV�X�e��/�a�c;         �Z���^�[�R�[�h�ǉ�;
'*M    2015/09/15; ���ʃV�X�e��G/�a�c;        TC�V�X�e���\�z;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterNumeric_Modosi(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_IsParameterNumeric_Modosi = True
    
    If FncChkHankakuNumber(arryCSVRecord(enmModColNo.TenCd)) = False Then     '�Y���X�R�[�h
        FncSetError (intNowLine, enmModColNo.TenCd, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Modosi = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmModColNo.HenkouDate)) = False Then '�ύX��
        FncSetError (intNowLine, enmModColNo.HenkouDate, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Modosi = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmModColNo.janCD)) = False Then     'JAN�R�[�h
        FncSetError (intNowLine, enmModColNo.janCD, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Modosi = False
    End If
    
    If arryCSVRecord(enmModColNo.Genka) <> "" Then
        If IsNumeric(arryCSVRecord(enmModColNo.Genka)) = False Then   '����
            FncSetError (intNowLine, enmModColNo.Genka, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Modosi = False
        End If
    End If
    
    If arryCSVRecord(enmModColNo.Baika) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmModColNo.Baika)) = False Then   '����
            FncSetError (intNowLine, enmModColNo.Baika, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Modosi = False
        End If
    End If
    
    If arryCSVRecord(enmModColNo.BaikaRiyuu) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmModColNo.BaikaRiyuu)) = False Then  '�����ύX���R
            FncSetError (intNowLine, enmModColNo.BaikaRiyuu, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Modosi = False
        End If
    End If
    
    If arryCSVRecord(enmModColNo.Torihikisaki) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmModColNo.Torihikisaki)) = False Then '�����R�[�h
            FncSetError (intNowLine, enmModColNo.Torihikisaki, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Modosi = False
        End If
    End If
    
    If arryCSVRecord(enmModColNo.Hacchuutani) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmModColNo.Hacchuutani)) = False Then '�����P��
            FncSetError (intNowLine, enmModColNo.Hacchuutani, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Modosi = False
        End If
    End If

    If arryCSVRecord(enmModColNo.CenterKbn) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmModColNo.CenterKbn)) = False Then '�Z���^�[�敪
            FncSetError (intNowLine, enmModColNo.CenterKbn, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Modosi = False
        End If
    End If
    
    If arryCSVRecord(enmModColNo.centerCD) <> "" Then
        '2015/09/15 wadak7 upd
        'If arryCSVRecord(enmModColNo.CenterKbn) = "1" Then      '�Z���^�[�敪��1�F���~�Z���^�[�̏ꍇ�A�`�F�b�N����B
        '�Z���^�[�敪��1�F���~�Z���^�[�E3�FTC�Z���^�[�̏ꍇ�A�`�F�b�N����B
        If arryCSVRecord(enmModColNo.CenterKbn) = "1" Or arryCSVRecord(enmModColNo.CenterKbn) = "3" Then
            If FncChkHankakuNumber(arryCSVRecord(enmModColNo.centerCD)) = False Then '�Z���^�[�R�[�h
                FncSetError (intNowLine, enmModColNo.centerCD, PARAM_NOTDISIT)
                FncCheck_IsParameterNumeric_Modosi = False
            End If
        End If
    End If
    
    If pstrKikakuSyurui = "012" Or pstrKikakuSyurui = "015" Then
    
        If arryCSVRecord(enmModColNo.MoGenka) <> "" Then
            If IsNumeric(arryCSVRecord(enmModColNo.MoGenka)) = False Then   '�߂�����
                FncSetError (intNowLine, enmModColNo.MoGenka, PARAM_NOTDISIT)
                FncCheck_IsParameterNumeric_Modosi = False
            End If
        End If
        
        If arryCSVRecord(enmModColNo.MoBaika) <> "" Then
            If FncChkHankakuNumber(arryCSVRecord(enmModColNo.MoBaika)) = False Then   '�߂�����
                FncSetError (intNowLine, enmModColNo.MoBaika, PARAM_NOTDISIT)
                FncCheck_IsParameterNumeric_Modosi = False
            End If
        End If
        
        If arryCSVRecord(enmModColNo.HattyuuTani) <> "" Then
            If FncChkHankakuNumber(arryCSVRecord(enmModColNo.HattyuuTani)) = False Then '�߂������P��
                FncSetError (intNowLine, enmModColNo.HattyuuTani, PARAM_NOTDISIT)
                FncCheck_IsParameterNumeric_Modosi = False
            End If
        End If
        
        
    '2009.06.19 Del by Supattra T.(Tostem Thai)Start.
'    ElseIf pstrKikakuSyurui = "015" Then
'
'        If arryCSVRecord(enmModColNo.MoGenka) <> "" Then
'            If IsNumeric(arryCSVRecord(enmModColNo.MoGenka)) = False Then   '�߂�����
'                FncSetError (intNowLine, enmModColNo.MoGenka, PARAM_NOTDISIT)
'                FncCheck_IsParameterNumeric_Modosi = False
'            End If
'        End If
'
'        If arryCSVRecord(enmModColNo.HattyuuTani) <> "" Then
'            If FncChkHankakuNumber(arryCSVRecord(enmModColNo.HattyuuTani)) = False Then '�߂������P��
'                FncSetError (intNowLine, enmModColNo.HattyuuTani, PARAM_NOTDISIT)
'                FncCheck_IsParameterNumeric_Modosi = False
'            End If
'        End If
     '2009.06.19 Del by Supattra T.(Tostem Thai)End.
    End If

End Function


'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncChkHankakuNumber
'*�y�����T�v�z
'*?    �w�肳�ꂽ�����񂪔��p�����񂩂ǂ������`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strValue ; String �`�F�b�N������������
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncChkHankakuNumber; Boolean;���p������:True,���p���l�ȊO���������Ă���:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
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
'*�y�v���V�[�W�����z
'*!    FncChkHankakuNarrow
'*�y�����T�v�z
'*?    �w�肳�ꂽ�����񂪔��p�����񂩂ǂ������`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strValue ; String �`�F�b�N������������
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncChkHankakuNarrow; Boolean;���p������:True,���p������ȊO���������Ă���:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/08/31; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
'
Private Function FncChkHankakuNarrow(ByVal strValue As String) As Boolean
        
    Dim lngCount As Long
    Dim LetterCount As Long          '������̕�����
    Dim ByteCount As Long            '������̃o�C�g��
    
    FncChkHankakuNarrow = False

    ''���p�`�F�b�N
    LetterCount = Len(strValue)
    ByteCount = LenB(StrConv(strValue, vbFromUnicode))

    If LetterCount = ByteCount Then
    Else
        Exit Function
    End If
    
    FncChkHankakuNarrow = True

End Function
'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncChkHankakuWide
'*�y�����T�v�z
'*?    �w�肳�ꂽ�����񂪑S�p�����񂩂ǂ������`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strValue ; String �`�F�b�N������������
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncChkHankakuWide; Boolean;���p������:True,���p������ȊO���������Ă���:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/08/31; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
'
Private Function FncChkHankakuWide(ByVal strValue As String) As Boolean
        
    Dim lngCount As Long
    Dim LetterCount As Long          '������̕�����
    Dim ByteCount As Long            '������̃o�C�g��
    
    FncChkHankakuWide = False

    ''�S�p�`�F�b�N
    LetterCount = Len(strValue)
    ByteCount = LenB(StrConv(strValue, vbFromUnicode))

    If LetterCount <> ByteCount / 2 Then
        Exit Function
    End If
    
    FncChkHankakuWide = True

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_Limit
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�t�B�[���h�������������z���Ă��Ȃ����`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_Limit; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_Limit(ByRef arryCSVRecord() As String, intNowLine As Integer)

    FncCheck_Limit = True
    
    '(a)�Y���X�R�[�h   �����ł̓`�F�b�N���Ȃ��i�}�X�^�o�^�`�F�b�N�ő�p�ł���ׁj�B
'    If FncCheck_TenCode(arryCSVRecord(enmTenColNo.TenCd)) = False Then
'        FncSetError (intNowLine, enmTenColNo.TenCd, PARAM_SHORTDISIT)
'        FncCheck_Limit = False
'    End If
    '(b)JAN�R�[�h
    If FncCheck_JANCode(arryCSVRecord(enmTenColNo.janCD)) = False Then
        FncSetError (intNowLine, enmTenColNo.janCD, PARAM_SHORTDISIT)
        FncCheck_Limit = False
    End If
    '(c)����
    If arryCSVRecord(enmTenColNo.Genka) <> "" Then
        If FncCheck_Genka(arryCSVRecord(enmTenColNo.Genka)) = False Then
            FncSetError (intNowLine, enmTenColNo.Genka, PARAM_SHORTDISIT)
            FncCheck_Limit = False
        End If
    End If
    '(d)����
    If arryCSVRecord(enmTenColNo.Baika) <> "" Then
        If FncCheck_Baika(arryCSVRecord(enmTenColNo.Baika)) = False Then
            FncSetError (intNowLine, enmTenColNo.Baika, PARAM_SHORTDISIT)
            FncCheck_Limit = False
        End If
    End If
    '(e)�����P��
    If arryCSVRecord(enmTenColNo.Hacchuutani) <> "" Then
        If FncCheck_Hattyu_Tani(arryCSVRecord(enmTenColNo.Hacchuutani)) = False Then
            FncSetError (intNowLine, enmTenColNo.Hacchuutani, PARAM_SHORTDISIT)
            FncCheck_Limit = False
        End If
    End If
    
End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_Limit_Modosi
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�t�B�[���h�������������z���Ă��Ȃ����`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_Limit_Modosi; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2009/06/06; Supattra T.(Tostem Thai); �V�K�쐬;
'*M    2011/02/10; �c�ƃV�X�e��/�a�c;         �Z���^�[�R�[�h�ǉ�;
'*M    2015/09/15; ���ʃV�X�e��G/�a�c;        TC�V�X�e���\�z;
'*< Comment End >**************************************************'
Private Function FncCheck_Limit_Modosi(ByRef arryCSVRecord() As String, intNowLine As Integer)

    FncCheck_Limit_Modosi = True
    
    '(b)JAN�R�[�h
    If FncCheck_JANCode(arryCSVRecord(enmModColNo.janCD)) = False Then
        FncSetError (intNowLine, enmModColNo.janCD, PARAM_SHORTDISIT)
        FncCheck_Limit_Modosi = False
    End If
    '(c)����
    If arryCSVRecord(enmModColNo.Genka) <> "" Then
        If FncCheck_Genka(arryCSVRecord(enmModColNo.Genka)) = False Then
            FncSetError (intNowLine, enmModColNo.Genka, PARAM_SHORTDISIT)
            FncCheck_Limit_Modosi = False
        End If
    End If
    '(d)����
    If arryCSVRecord(enmModColNo.Baika) <> "" Then
        If FncCheck_Baika(arryCSVRecord(enmModColNo.Baika)) = False Then
            FncSetError (intNowLine, enmModColNo.Baika, PARAM_SHORTDISIT)
            FncCheck_Limit_Modosi = False
        End If
    End If
    '(e)�����P��
    If arryCSVRecord(enmModColNo.Hacchuutani) <> "" Then
        If FncCheck_Hattyu_Tani(arryCSVRecord(enmModColNo.Hacchuutani)) = False Then
            FncSetError (intNowLine, enmModColNo.Hacchuutani, PARAM_SHORTDISIT)
            FncCheck_Limit_Modosi = False
        End If
    End If
    '�Z���^�[�R�[�h
    If arryCSVRecord(enmModColNo.centerCD) <> "" Then
        '2015/09/15 wadak7 upd
        'If arryCSVRecord(enmModColNo.CenterKbn) = "1" Then      '�Z���^�[�敪��1�F���~�Z���^�[�̏ꍇ�A�`�F�b�N����B
        '�Z���^�[�敪��1�F���~�Z���^�[�E3�FTC�Z���^�[�̏ꍇ�A�`�F�b�N����B
        If arryCSVRecord(enmModColNo.CenterKbn) = "1" Or arryCSVRecord(enmModColNo.CenterKbn) = "3" Then
            If FncCheck_CenterCd(arryCSVRecord(enmModColNo.centerCD)) = False Then
                FncSetError (intNowLine, enmModColNo.centerCD, PARAM_SHORTDISIT)
                FncCheck_Limit_Modosi = False
            End If
        End If
    End If
    
    If Trim(pstrKikakuSyurui) = "012" Or Trim(pstrKikakuSyurui) = "015" Then
    
         '�߂�����
        If arryCSVRecord(enmModColNo.MoGenka) <> "" Then
            If FncCheck_Genka(arryCSVRecord(enmModColNo.MoGenka)) = False Then
                FncSetError (intNowLine, enmModColNo.MoGenka, PARAM_SHORTDISIT)
                FncCheck_Limit_Modosi = False
            End If
        End If
        
        '�߂�����
        If arryCSVRecord(enmModColNo.MoBaika) <> "" Then
            If FncCheck_Baika(arryCSVRecord(enmModColNo.MoBaika)) = False Then
                FncSetError (intNowLine, enmModColNo.MoBaika, PARAM_SHORTDISIT)
                FncCheck_Limit_Modosi = False
            End If
        End If
        
        '�߂������P��
        If arryCSVRecord(enmModColNo.HattyuuTani) <> "" Then
            If FncCheck_Hattyu_Tani(arryCSVRecord(enmModColNo.HattyuuTani)) = False Then
                FncSetError (intNowLine, enmModColNo.HattyuuTani, PARAM_SHORTDISIT)
                FncCheck_Limit_Modosi = False
            End If
        End If
        'POP�T�C�Y
        If arryCSVRecord(enmModColNo.PopSize) <> "" Then
            If LenB(StrConv(arryCSVRecord(enmModColNo.PopSize), vbFromUnicode)) > 20 Then
                FncSetError (intNowLine, enmModColNo.PopSize, PARAM_SHORTDISIT)
                FncCheck_Limit_Modosi = False
            End If
        End If
        
    '2009.06.19 Del by Supattra T.(Tostem Thai)Start.
'    ElseIf Trim(pstrKikakuSyurui) = "015" Then
'
'          '�߂�����
'        If arryCSVRecord(enmModColNo.MoGenka) <> "" Then
'            If FncCheck_Genka(arryCSVRecord(enmModColNo.MoGenka)) = False Then
'                FncSetError (intNowLine, enmModColNo.MoGenka, PARAM_SHORTDISIT)
'                FncCheck_Limit_Modosi = False
'            End If
'        End If
'
'         '�߂������P��
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
'*�y�v���V�[�W�����z
'*!    FncCheck_TenCode
'*�y�����T�v�z
'*?    �X�R�[�h�̕����񒷃`�F�b�N���s���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strTenCode ; String �X�R�[�h
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_TenCode; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
'���̃`�F�b�N�͍s��Ȃ��i�}�X�^�o�^�`�F�b�N�ő�p�ł���ׁj�B
'Private Function FncCheck_TenCode(strTenCode As String) As Boolean
'    '�Y���X�R�[�h�F����4���܂�
'    FncCheck_TenCode = FncCheck_Length_Less(strTenCode, LENGTH_TENCODE)
'End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_Baika
'*�y�����T�v�z
'*?    �����̕����񒷃`�F�b�N���s���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strBaika ; String ����
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_Baika; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_Baika(strBaika As String) As Boolean
    '�����F����7���܂�
    FncCheck_Baika = FncCheck_Length_Less(strBaika, LENGTH_BAIKA)
End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_Length_Kotei
'*�y�����T�v�z
'*?    ������̕����񒷃`�F�b�N���s���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strNumberString ; String ������
'*P    lngLength       ; Long   ����
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_Length_Kotei; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/08/31; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_Length_Kotei(strNumberString As String, lngLength As Long) As Boolean
    FncCheck_Length_Kotei = False
    If Len(strNumberString) = lngLength Then
        FncCheck_Length_Kotei = True
    End If
End Function


'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_Hattyu_Tani
'*�y�����T�v�z
'*?    �����P�ʂ̕����񒷃`�F�b�N���s���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strHattyu_Tani ; String �����P��
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_Hattyu_Tani; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_Hattyu_Tani(strHattyu_Tani As String) As Boolean
    '�����P�ʁF����5���܂�
    FncCheck_Hattyu_Tani = FncCheck_Length_Less(strHattyu_Tani, LENGTH_HATTYU_TANI)
End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_Length_Less
'*�y�����T�v�z
'*?    �w�肳�ꂽ�����񂪁A�w�肳�ꂽ�����ȉ��ł��邩�ǂ����𒲂ׂ܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strNumberString ; String ������
'*P    lngLength ; Long ����������
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_Length_Less; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_Length_Less(strNumberString As String, lngLength As Long)
    FncCheck_Length_Less = False
    If Len(strNumberString) <= lngLength Then
        FncCheck_Length_Less = True
    End If
End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_JANCode
'*�y�����T�v�z
'*?    JAN�R�[�h�̕����񒷃`�F�b�N���s���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strJanCode ; String JAN�R�[�h
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_JANCode; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_JANCode(strJanCode As String) As Boolean
    FncCheck_JANCode = False
    'JAN�R�[�h�F8������13��
    If Len(strJanCode) = LENGTH_JANCODE1 Or Len(strJanCode) = LENGTH_JANCODE2 Then
        FncCheck_JANCode = True
    End If
End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_CenterCd
'*�y�����T�v�z
'*?    JAN�R�[�h�̕����񒷃`�F�b�N���s���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strCenterCd ; String �Z���^�[�R�[�h
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_CenterCd; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2011/02/10; �c�ƃV�X�e��/�a�c; �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_CenterCd(strCenterCd As String) As Boolean
    FncCheck_CenterCd = False
    '�Z���^�[�R�[�h�F4��
    If Len(strCenterCd) = LENGTH_CENTERCODE Then
        FncCheck_CenterCd = True
    End If
End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_Genka
'*�y�����T�v�z
'*?    �����̕����񒷃`�F�b�N���s���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strGenka ; String ����
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_Genka; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_Genka(strGenka As String) As Boolean
    '�����F�������V���A�������Q��
    FncCheck_Genka = FncCheck_Syousuu(strGenka, LENGTH_GENKA_SEISU, LENGTH_GENKA_SYOUSUU)
End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_Syousuu
'*�y�����T�v�z
'*?    �����`���̎����ɑ΂��A�w�肳�ꂽ�����ȉ��ɂȂ��Ă��邩�ǂ����𒲂ׂ܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strValue ; String ����������
'*P    intSeisuuKeta ; Integer �������ő包��
'*P    intSyousuuKeta ; Integer �������ő包��
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_Syousuu; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_Syousuu(ByVal strValue As String, _
                    ByVal intSeisuuKeta As Integer, _
                    ByVal intSyousuuKeta As Integer) As Boolean

    Dim lngCount As Long
    Dim arrySyousuu() As String

    FncCheck_Syousuu = False
    
    If IsNumeric(strValue) = False Then
        ' ���l�ł͂Ȃ�
        Exit Function
    End If
    
    For lngCount = 1 To Len(strValue)
        If Not Mid(strValue, lngCount, 1) Like "[0-9.]" Then
            ' ���l�ȊO�̕����������Ă���
            Exit Function
        End If
            
        '���������͂��ꂽ�ꍇ
        If InStr(strValue, ".") = 0 Then
            '�����ӂ�`�F�b�N
            If Len(strValue) > intSeisuuKeta Then
              
                Exit Function
            End If
        '�����̏ꍇ
        Else
            arrySyousuu = Split(strValue, ".")
            '�����ӂ�`�F�b�N
            If Len(arrySyousuu(0)) > intSeisuuKeta Then
                Exit Function
            End If
            
            '�����ӂ�`�F�b�N
            If Len(arrySyousuu(1)) > intSyousuuKeta Then
                '�����������ӂ�
                Exit Function
            End If
                    
        End If
    
    Next
    
    FncCheck_Syousuu = True

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_Henkoubi
'*�y�����T�v�z
'*?    CSV�f�[�^�̕ύX���ɑ΂��`�F�b�N���s���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strHenkoubi ; String �ύX��������������
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_Henkoubi; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_Henkoubi(strHenkoubi As String, intNowLine As Integer) As Boolean

    FncCheck_Henkoubi = True
    ' �������W���ł��鎖�ƁA�ύX���̓��t���L���ł��邩���m�F����B
    If Len(strHenkoubi) <> 8 Or FncCheckDate(strHenkoubi) = False Then
        FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_WRONGDATE)
        FncCheck_Henkoubi = False
'********** 2005/05/18 �y�c �ǉ� **********��
    Else
        If pstrMinDate > strHenkoubi Then
            pstrMinDate = strHenkoubi
        ElseIf pstrMaxDate < strHenkoubi Then
            pstrMaxDate = strHenkoubi
        End If
'********** 2005/05/18 �y�c �ǉ� **********��
    End If

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckDate
'*�y�����T�v�z
'*?    ���t������W�������������t�ł��邱�Ƃ��`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strDate ; YYYYMMDD�Ŏ������W���̓��t������
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckDate; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*M    2005/04/27; ����(USC); ���t�`�F�b�N��CSV�捞�ɍ��킹��悤�ɕύX�B;
'*M                          �i19000101 �` 90001231�j;
'*< Comment End >**************************************************'
Private Function FncCheckDate(strDate As String) As Boolean
    
    Dim dtChk    As Date
    FncCheckDate = False
    
    '�͈̓`�F�b�N 19000101 �` 90001231 �Ƃ���
    If StrComp(strDate, 90001231) > 0 Then
        FncCheckDate = False
        Exit Function
    End If
    
    If StrComp(strDate, 19000101) < 0 Then
        FncCheckDate = False
        Exit Function
    End If
    
    '�󔒂������Ԃɂ���ꍇ�́A�G���[
    If strDate <> Replace(strDate, " ", "") Then
        FncCheckDate = False
        Exit Function
    End If
    
    Dim strCheckDate
    strCheckDate = Format$(strDate, "@@@@/@@/@@")
    FncCheckDate = IsDate(strCheckDate)
End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_CheckDigit
'*�y�����T�v�z
'*?    CSV�f�[�^��JAN�R�[�h�ɑ΂��`�F�b�N�f�B�W�b�g�𒲂ׂ܂��B
'*?    �`�F�b�N�f�B�W�b�g�̓E�F�C�g 3-1 �̃��W�����X10�ɂ����
'*?    �v�Z����A���ׂ�JAN�R�[�h�̖����ɒǉ�����Ă�����̂Ƃ��܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strJanCode ; String JAN�R�[�h������������
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*P    intFieldType; Integer; 1:�X�ʏ���JAN�R�[�h
'*P    �@�@�@�@�@�@�@�@�@�@�@�@2:�������~������JAN�R�[�h
'*P    �@�@�@�@�@�@�@�@�@�@�@�@3:�������~�����̑�֏��i�R�[�h
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_CheckDigit; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*M    2005/04/26; ����(USC); �������~�p�ł��g�p�ł���悤�Ɉ�����ǉ�;
'*< Comment End >**************************************************'
Private Function FncCheck_CheckDigit(strJanCode As String, intNowLine As Integer, intFieldType As Integer) As Boolean

    Dim varTmp As Object
    Dim strChkDigit As String
    Dim strCode As String
    Dim lngFactor As Long
    Dim lngSum As Long
    Dim lngCount As Long
    
    FncCheck_CheckDigit = False

'********** 2005/09/05 ���� �ǉ� **********��
'********** 2005/04/26 ���� �ǉ� **********��
'���ږ��̍쐬�i�G���[���b�Z�[�W�p�j
    Dim intField As Integer
    Select Case intFieldType
        Case 1:     intField = enmTenColNo.janCD
        Case 2:     intField = enmHacColNo.janCD
        Case 3:     intField = enmHacColNo.DaiJanCd
        Case 4:     intField = enmSyoColNo.janCD
        Case 5:     intField = enmBunColNo.janCD
        Case Else:  intField = enmTenColNo.janCD
    End Select
'********** 2005/04/26 ���� �ǉ� **********��
'********** 2005/09/05 ���� �ǉ� **********��

    '���̃}�X�^�`�F�b�N
    lngFactor = 3
    lngSum = 0

    varTmp = strJanCode

    '�����`�F�b�N
    If Len(varTmp) <> LENGTH_JANCODE1 And Len(varTmp) <> LENGTH_JANCODE2 Then
'********** 2005/04/26 ���� �ύX **********��
        FncSetError (intNowLine, intField, PARAM_WRONGCHECKDIGIT
'        FncSetError (intNowLine, enmTenColNo.JanCd, PARAM_WRONGCHECKDIGIT
'********** 2005/04/26 ���� �ύX **********��
        GoTo Proc_End
    End If
            
    'JAN�R�[�h���ő包�ɕ�����0����
    varTmp = Format(varTmp, String(LENGTH_JANCODE2, "0"))
            
    '�`�F�b�N�f�W�b�g�擾
    strChkDigit = Right(varTmp, 1)
            
    '�R�[�h���擾
    strCode = Left(varTmp, Len(varTmp) - 1)
    
    If IsNumeric(strCode) Then
        For lngCount = Len(strCode) To 1 Step -1
            lngSum = lngSum + Mid(strCode, lngCount, 1) * lngFactor
            lngFactor = 4 - lngFactor
        Next
            
        '�`�F�b�N�f�W�b�g�m�F
        If strChkDigit <> ((1000 - lngSum) Mod 10) Then
    '********** 2005/04/26 ���� �ύX **********��
            FncSetError (intNowLine, intField, PARAM_WRONGCHECKDIGIT
    '        FncSetError (intNowLine, enmTenColNo.JanCd, PARAM_WRONGCHECKDIGIT
    '********** 2005/04/26 ���� �ύX **********��
            GoTo Proc_End
        End If
    End If
    
    FncCheck_CheckDigit = True

Proc_End:

End Function
'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckParameter_DataBase
'*�y�����T�v�z
'*?    CSV�t�@�C���̃p�����[�^�ɑ΂��āADB�ɃA�N�Z�X���ă`�F�b�N���s���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVRecord() ; String CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckParameter_DataBase; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*M    2011/02/16; �c�ƃV�X�e��/�a�c; �Z���^�[�R�[�h�ǉ�;
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
    
    'a.�Y���X�R�[�h�����Ə��}�X�^�EWF���̃}�X�^�e�[�u���ɓo�^����Ă��邩�m�F����B
    blnIsOKTenCode = FncCheckDB_TenCode(arryCSVRecord(enmTenColNo.TenCd), intNowLine)
    
    'b.JAN�R�[�h���L�����m�F����B
'********** 2005/04/26 ���� �ύX **********��
    blnIsOKJANCode = FncCheckDB_JANCode(arryCSVRecord(enmTenColNo.janCD), strKeyCode, intNowLine, RET_FLD_TENB)
'    blnIsOKJANCode = FncCheckDB_JANCode(arryCSVRecord(enmTenColNo.JanCd), strKeyCode, intNowLine)
'********** 2005/04/26 ���� �ύX **********��
    
'********** 2005/09/13 ���� �ύX **********��
    'b2.KEY�R�[�h���L�����m�F����B
    blnIsOKNaibuCode = FncCheckDB_NaibuCode(strKeyCode, intNowLine, RET_FLD_TENB)
'********** 2005/09/13 ���� �ύX **********��
    
    'c.�����ύX���R���L�����m�F����B
    blnIsOKBaikaHenkouRiyuu = True
    If arryCSVRecord(enmTenColNo.BaikaRiyuu) <> "" Then
        blnIsOKBaikaHenkouRiyuu = FncCheckDB_BaikaHenkouRiyuu(arryCSVRecord(enmTenColNo.BaikaRiyuu), intNowLine)
    End If
    
    'd.����悪�L�����m�F����B
    blnIsOKTorihikisaki = True
    If arryCSVRecord(enmTenColNo.Torihikisaki) <> "" Then
        blnIsOKTorihikisaki = FncCheckDB_Torihikisaki(arryCSVRecord(enmTenColNo.Torihikisaki), intNowLine)
    End If
    
    'e.�Z���^�[�敪�A�Z���^�[�R�[�h���L�����m�F����B
    strIsOKCenterClass = "0"    '0:���� 1:�Z���^�[�敪�ُ� 2:�Z���^�[�R�[�h�ُ�
    If arryCSVRecord(enmTenColNo.CenterKbn) <> "" Then
        strIsOKCenterClass = FncCheckDB_CenterClass(arryCSVRecord(enmTenColNo.CenterKbn), arryCSVRecord(enmModColNo.centerCD), intNowLine)
    End If
    
    'f.�����}�X�^�̐�t�ύX���ɓ����̕ύX����񂪖������m�F����B
    blnIsOKSakituke = True
    If blnIsOKTenCode = True Then   'a.�̃`�F�b�N��OK�̏ꍇ�Ƀ`�F�b�N����B
        blnIsOKSakituke = FncCheckDB_Sakituke(strKeyCode, arryCSVRecord, intNowLine)
    End If
    
    If blnIsOKTenCode And blnIsOKJANCode And blnIsOKBaikaHenkouRiyuu And _
       blnIsOKTorihikisaki And strIsOKCenterClass = "0" And blnIsOKSakituke And blnIsOKNaibuCode Then
        FncCheckParameter_DataBase = True
    End If
 
End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckDB_TenCode
'*�y�����T�v�z
'*?    �Y���X�R�[�h�����Ə��}�X�^�EWF���̃}�X�^�ɓo�^����Ă��邩�ǂ����m�F���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strTenCode ; String �X�R�[�h������킷������
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckDB_TenCode; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheckDB_TenCode(ByVal strTenCode As String, intNowLine As Integer) As Boolean
    
'********** 2005/05/11 �y�c �g�p���Ă��Ȃ��̂ō폜 **********��
'    Dim blnIsFindTenCode As Boolean
'    Dim blnIsFindJigyouCode As Boolean
'    Dim intLoop As Integer
'    Dim strJigyouCode As String
'********** 2005/05/11 �y�c �폜 **********��

    FncCheckDB_TenCode = True
    
    If FncCheckDB_FindTenCode_TenCode(strTenCode) = False Then
        ' �w�肳�ꂽ�Y���X�R�[�h�͓W�J���Ă��Ȃ�
        FncSetError (intNowLine, enmTenColNo.TenCd, PARAM_NOTEXISTDB
        FncCheckDB_TenCode = False
    End If
    
End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckDB_FindTenCode_TenCode
'*�y�����T�v�z
'*?    �Y���X�R�[�h�����e�[��DB���Ə��}�X�^�e�[�u���ɑ��݂��邩���m�F���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strTenCode ; String �X�R�[�h������킷������
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckDB_FindTenCode_TenCode; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*M    2005/05/11; �y�c�@�S��(USC); �p�t�H�[�}���X����̈פ���̊֐����ł̐ؒf���R�[�h�擾����ߑ��ϐ��̐ؒf���R�[�h�𗘗p����l�C��;
'*< Comment End >**************************************************'
Private Function FncCheckDB_FindTenCode_TenCode(strTenCode As String) As Boolean

'********** 2005/05/11 �y�c �폜 **********��
'DB�ڑ��Ɋւ��鏈�����폜
'    Dim strSql            As String
'    Dim objCDBAccess      As Object
'    Dim objRS      As Object
'********** 2005/05/11 �y�c �ǉ� **********��
            
    FncCheckDB_FindTenCode_TenCode = False
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncCheckDB_FindTenCode_TenCode(" & strTenCode & ")"
 
    ' �{��(0000)���w�肳�ꂽ�ꍇ�̓`�F�b�N�����ɁA(0000)�����̂܂ܓX�R�[�h�Ƃ��ĕԂ�
    If StrComp(strTenCode, HON_TEN_CD) = 0 Then
        FncCheckDB_FindTenCode_TenCode = True
'********** 2005/05/11 �y�c �ύX **********��
'DB�ڑ��Ɋւ��鏈�����폜
'        GoTo Proc_End
        Exit Function
'********** 2005/05/11 �y�c �ύX **********��
    End If
    
    
'********** 2005/05/11 �y�c �폜 **********��
'���̊֐����ł̐ؒf���R�[�h�Z�b�g�擾�����폜
'    'SQL���쐬�i�����F���ʋ敪=5�A�X����NULLor�ꊇ�o�^���ȍ~�AKEY�敪=W02�A�ݒ���e��4���ڂ�1�A�X�R�[�h�j
'    strSql = "SELECT A.jigyousyo_cd "
'    strSql = strSql & "FROM v_jigyousyo AS A WITH(readcommitted) "
'    strSql = strSql & "INNER JOIN m_wf_meisyou AS B WITH(readcommitted) ON A.jigyou_cd = B.key_cd " & vbCrLf
'    strSql = strSql & "WHERE A.hanbetu_kbn = '" & HANBETU_KBN_TENPO & "' AND "
'    strSql = strSql & "(RTrim(A.heiten_date) = '' OR A.heiten_date is null OR A.heiten_date >= " & FncGetNowDate() & ") "
'    strSql = strSql & "AND Substring(B.settei_naiyou,4,1) = '1' " & vbCrLf
'    strSql = strSql & "AND B.key_kbn = '" & KEYCD_KIKAKU_DOUNYU_TAISYOU & "' " & vbCrLf
'    '�S�X�A�S���Ƃ̏ꍇ�́A�ȉ��̏����͕K�v�Ȃ��B
'    If strTenCode <> ALL_TEN_CD And strTenCode <> ALL_JIG_CD Then
'        If Len(strTenCode) = 2 Then
'            strSql = strSql & "AND A.jigyou_cd = '" & strTenCode & "'"      '�e����
'        Else
'            strSql = strSql & "AND A.jigyousyo_cd = '" & strTenCode & "'"   '�e�X
'        End If
'    End If
    
'    'CDBAccess�C���X�^���X����
'    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
'    'SQL���s
'        objRS = objCDBAccess.GetDCRecordset("VSWF", strSql)
'********** 2005/05/11 �y�c �폜 **********��
    
'********** 2005/05/11 �y�c �ǉ� **********��
'���ϐ��̐ؒf���R�[�h�Z�b�g���g�p���G���[�`�F�b�N

    

    If strTenCode <> ALL_TEN_CD And strTenCode <> ALL_JIG_CD Then
        
        Dim strTarget As String '�����Ώہi���Ƃ��X���j
        If Len(strTenCode) = 2 Then
            strTarget = "jigyou_cd"      '�e����
        Else
            strTarget = "jigyousyo_cd"  '�e�X
        End If
        
        With pTChkJigyouRS

'********** 2005/05/17 �y�c �ύX **********��
'********** 2005/05/18 �y�c �ύX **********��
'            If Not .EOF Then
           .Filter = adFilterNone
            If Not .BOF Then
'********** 2005/05/18 �y�c �ύX **********��
                .MoveFirst
                .Find strTarget & " = '" & strTenCode & "'"
            End If
'********** 2005/05/17 �y�c �ύX **********��
        End With
    
    End If
    
    
    '�X�i���Ɓj�R�[�h�������ꍇ�̓G���[
    FncCheckDB_FindTenCode_TenCode = Not pTChkJigyouRS.EOF
'********** 2005/05/11 �y�c �ǉ� **********��

'********** 2005/05/11 �y�c �폜 **********��
'DB�ڑ��Ɋւ��鏈�����폜
'    If objRS.EOF = True Then
'        FncCheckDB_FindTenCode_TenCode = False
'    Else
'        FncCheckDB_FindTenCode_TenCode = True
'    End If
    
    'DB�N���[�Y
'    objRS.Close
'        objRS = Nothing
'    objCDBAccess = Nothing
    
'    'GetObjectContext.SetComplete
'********** 2005/05/11 �y�c �폜 **********��
Exit Function

'********** 2005/05/11 �y�c �ύX **********��
'DB�ڑ��Ɋւ��鏈�����폜
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
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description
'********** 2005/05/11 �y�c �ύX **********��

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncGetNowDate
'*�y�����T�v�z
'*?    ���݂̓��t���W�����̔��p���l�ŕԂ��܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncGetNowDate; String;YYYYMMDD�ŕ\�������W���̓��t
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncGetNowDate() As String

    Dim dtNowDate, strNowDate
    dtNowDate = Date
    strNowDate = Format(dtNowDate, "YYYYMMDD")
    
    FncGetNowDate = strNowDate

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckDB_JANCode
'*�y�����T�v�z
'*?    JAN�R�[�h���L�������`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strJanCode ; String JAN�R�[�h������킷������
'*P    strKeyCode ; String JAN�R�[�h�ɑΉ����铝���}�X�^�̃L�[�R�[�h���i�[����̈�
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*P    intFieldType; Integer; RET_FLD_HACH:�������~�����
'*P    �@�@�@�@�@�@�@�@�@�@�@�@ RET_FLD_TENB:�X�ʏ��
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    strKeyCode ; String JAN�R�[�h�ɑΉ����铝���}�X�^�̃L�[�R�[�h���i�[����
'*R    FncCheckDB_JANCode; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*M    2005/04/26; ����(USC); �������~�ł��g�p�ł���悤�Ɉ����̒ǉ�;
'*M    2006/09/12; �a�c(IC); ���i�I��v�e�⍇���o�ŉ��P;
'*M    2009/05/12; �a�c(IC); �O���[�v�i�̓G���[�ɂ���;
'*< Comment End >**************************************************'
'********** 2005/04/26 ���� �ύX **********��
Private Function FncCheckDB_JANCode(ByVal strJanCode As String, ByRef strKeyCode As String, intNowLine As Integer, intFieldType As Integer) As Boolean
'Private Function FncCheckDB_JANCode(ByVal strJanCode As String, ByRef strKeyCode As String, intNowLine As Integer) As Boolean
'********** 2005/04/26 ���� �ύX **********��
    Dim strKikakuNo As String   'JAN�R�[�h�����ɓo�^����Ă����懂
    
    'b.JAN�R�[�h���L�����m�F����B
    FncCheckDB_JANCode = True
    
'********** 2005/04/26 ���� �ǉ� **********��
'���ږ��̍쐬�i�G���[���b�Z�[�W�p�j
    Dim intField As Integer
    Select Case intFieldType
        Case RET_FLD_HACH:      intField = enmHacColNo.janCD
        Case RET_FLD_TENB:      intField = enmTenColNo.janCD
        Case RET_FLD_SYO:       intField = enmSyoColNo.janCD
        Case RET_FLD_BUN:       intField = enmBunColNo.janCD
        Case Else:              intField = enmTenColNo.janCD
    End Select
'********** 2005/04/26 ���� �ǉ� **********��
    
    '�@(a)WF�}�X�^�ɓ����JAN�R�[�h�����������m�F����B
    If FncCheckDB_SameJANCode(strJanCode, strKikakuNo) = True Then
'********** 2005/04/26 ���� �ύX **********��
        FncSetError (intNowLine, intField, PARAM_SAMECODEDB, strKikakuNo
'        FncSetError (intNowLine, enmTenColNo.JanCd, PARAM_SAMECODEDB, strJanCode
'********** 2005/04/26 ���� �ύX **********��
        FncCheckDB_JANCode = False
    End If
            
    '�@(b)�����}�X�^�ɓo�^�ς݂��m�F����B
    If FncCheckDB_EntryJANCode(strJanCode, strKeyCode) = False Then
'********** 2005/04/26 ���� �ύX **********��
        FncSetError (intNowLine, intField, PARAM_NOTEXISTDB, strJanCode
'        FncSetError (intNowLine, enmTenColNo.JanCd, PARAM_NOTEXISTDB, strJanCode
'********** 2005/04/26 ���� �ύX **********��
        FncCheckDB_JANCode = False
    End If

    '�O���[�v�i�̏ꍇ�A�G���[
    If intFieldType = RET_FLD_TENB Then
        If FncCheckDB_Group(strJanCode) = False Then
            FncSetError (intNowLine, intField, PARAM_GROUP, strJanCode
            FncCheckDB_JANCode = False
        End If
    End If

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckDB_SameJANCode
'*�y�����T�v�z
'*?    ���i�I��DB�iWF�Ǘ����ځj�e�[�u���ɏd������JAN�R�[�h�����݂��邩�𒲂ׂ܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strJanCode ; String JAN�R�[�h������킷������
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckDB_SameJANCode; Boolean;�d�����đ���:True,���݂��Ȃ�:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*M    2005/05/11; �y�c�@�S��(USC); �p�t�H�[�}���X����̈פ���̊֐����ł̐ؒf���R�[�h�擾����ߑ��ϐ��̐ؒf���R�[�h�𗘗p����l�C��;
'*M    2006/09/12; �a�c(IC); ���i�I��v�e�⍇���o�ŉ��P;
'*< Comment End >**************************************************'
Private Function FncCheckDB_SameJANCode(ByVal strJanCode As String, _
                                  ByRef strKikakuNo As String) As Boolean

    '   �@DB�F      ���i�I��DB�iWF�Ǘ����ځj�e�[�u���it_syousen_wf_kanri�j
    '   �@�L�[�F        JAN�R�[�h�ijan_cd�j                         CSV�t�@�C������擾
    '   �J�ڎ��Ɏ擾��������ރR�[�h�E�N�A�ԁE�}�Ԃ�����JAN�R�[�h�ȊO��
    '   �o�^����Ă��Ȃ����m�F����B
    
'********** 2005/05/11 �y�c �폜 **********��
'���̊֐����ł̐ؒf���R�[�h�Z�b�g�擾�����폜
'    Dim strSql            As String
 
'    strSql = "SELECT jan_cd FROM t_syousen_wf_kanri WITH(readcommitted) WHERE jan_cd = '" & strJanCode & "'"
'    strSql = strSql & " AND NOT("
'    strSql = strSql & "kikaku_syurui_cd = '" & pstrKikakuSyurui & "' "
'    strSql = strSql & " AND nen_ren_no = '" & pstrKikakuNenrenban & "' "
'    strSql = strSql & " AND eda_no = '" & pstrKikakuEdaban & "'"
'    strSql = strSql & ")"
     
'    FncCheckDB_SameJANCode = FncCheckDB_IsRecordExists("VSWF", strSql)
'********** 2005/05/11 �y�c �폜 **********��

'********** 2005/05/11 �y�c �ǉ� **********��
'���ϐ��̐ؒf���R�[�h�Z�b�g���g�p���G���[�`�F�b�N
    With pChkSameJANCodeRS
'********** 2005/05/17 �y�c �ύX **********��
'********** 2005/05/18 �y�c �ύX **********��
'        If Not .EOF Then
        If Not .BOF Then
'********** 2005/05/18 �y�c �ύX **********��
            .MoveFirst
            .Find "jan_cd = '" & strJanCode & "'"
            If Not .EOF Then
                strKikakuNo = .Fields("kikaku_no").Value
            End If
        End If
'********** 2005/05/17 �y�c �ύX **********��
    End With
 
    FncCheckDB_SameJANCode = Not pChkSameJANCodeRS.EOF
'********** 2005/05/11 �y�c �ǉ� **********��

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckDB_EntryJANCode
'*�y�����T�v�z
'*?    JAN�R�[�h�������}�X�^��KEY���1�e�[�u���ɓo�^����Ă��邩�𒲂ׁA
'*?    �o�^����Ă���΂��̃L�[�R�[�h��Ԃ��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strJanCode ; String JAN�R�[�h������킷������
'*P    strKeyCode ; String �L�[�R�[�h���i�[���镶����̈�
'*�y�߂�l�z
'*    < Name Type Comment >
'*     strKeyCode   ; String�@�o�^����Ă���΃L�[�R�[�h��Ԃ�
'*R    FncCheckDB_EntryJANCode; Boolean;�o�^����Ă���:True,���݂��Ȃ�:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*M    2005/05/11; �y�c�@�S��(USC); �p�t�H�[�}���X����̈פ���̊֐����ł̐ؒf���R�[�h�擾����ߑ��ϐ��̐ؒf���R�[�h�𗘗p����l�C��;
'*< Comment End >**************************************************'
Private Function FncCheckDB_EntryJANCode(ByVal strJanCode, ByRef strKeyCode As String) As Boolean
    
    '   �}�X�^�o�^����Ă��邩�́A�����}�X�^��KEY���1�e�[�u���Ŋm�F����B
    '   �@DB�F      �����}�X�^ KEY���1�iv_key_1�j
    '   �@�L�[�F        JAN�R�[�h�ijan_cd�j                         CSV�t�@�C������擾
    '   JAN�R�[�h�Ǝ�L�[���R�t�����Ă��鎖���m�F����B

'********** 2005/05/11 �y�c �폜 **********��
'DB�ڑ��Ɋւ��鏈�����폜
'    Dim strSql            As String
'    Dim objCDBAccess      As Object
'    Dim objRS             As Object
'********** 2005/05/11 �y�c �폜 **********��
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncCheckDB_EntryJANCode(" & strJanCode & "," & strKeyCode & ")"
    
    FncCheckDB_EntryJANCode = False

'********** 2005/05/11 �y�c �폜 **********��
'���̊֐����ł̐ؒf���R�[�h�Z�b�g�擾�����폜
'    strSql = "SELECT key_cd , jan_cd FROM v_key_1 WITH(readcommitted) WHERE jan_cd = '" & strJanCode & "'"
    
'    'CDBAccess�C���X�^���X����
'    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")

'    'SQL���s
'        objRS = objCDBAccess.GetDCRecordset("VSWF", strSql)
    
'    ' �f�[�^���擾�ł��Ȃ���΁A�e�[�u�����ɂ͊Y�����R�[�h�͑��݂��Ȃ�
'    If objRS.EOF = True Then
'        FncCheckDB_EntryJANCode = False
'        GoTo Proc_End
'    Else
'        ' �f�[�^���擾�ł�����A���̃L�[�R�[�h���擾����
'        FncCheckDB_EntryJANCode = True
'        strKeyCode = objRS.Fields("key_cd")
'    End If
    
'    objRS.Close
'        objRS = Nothing
'    objCDBAccess = Nothing
    
'    'GetObjectContext.SetComplete
'********** 2005/05/11 �y�c �폜 **********��

'********** 2005/05/11 �y�c �ǉ� **********��
'���ϐ��̐ؒf���R�[�h�Z�b�g���g�p���G���[�`�F�b�N
    ' �f�[�^���擾�ł�����A���̃L�[�R�[�h���擾����
    With pChkvk1JANCodeRS
'********** 2005/05/17 �y�c �ύX **********��
'********** 2005/05/18 �y�c �ύX **********��
'         If Not .EOF Then
        If Not .BOF Then
'********** 2005/05/18 �y�c �ύX **********��
            .MoveFirst
            .Filter = "jan_cd = '" & strJanCode & "' "
'            .Find "jan_cd = '" & strJanCode & "'"
            If .EOF = False Then
                strKeyCode = .Fields("key_cd")
            End If
        End If
'********** 2005/05/17 �y�c �ύX **********��
    End With
 
    FncCheckDB_EntryJANCode = Not pChkvk1JANCodeRS.EOF
    pChkvk1JANCodeRS.Filter = adFilterNone
'********** 2005/05/11 �y�c �ǉ� **********��
Exit Function
    
'********** 2005/05/11 �y�c �ύX **********��
'DB�ڑ��Ɋւ��鏈�����폜
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
'********** 2005/05/11 �y�c �ύX **********��
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckDB_Torihikisaki
'*�y�����T�v�z
'*?    ����悪�L������Ԃ��܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strTorihikisaki ; String ����������킷������
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckDB_Torihikisaki; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*M    2005/05/11; �y�c�@�S��(USC); �p�t�H�[�}���X����̈פ���̊֐����ł̐ؒf���R�[�h�擾����ߑ��ϐ��̐ؒf���R�[�h�𗘗p����l�C��;
'*< Comment End >**************************************************'
Private Function FncCheckDB_Torihikisaki(strTorihikisaki As String, intNowLine As Integer)

    'd.����悪�L�����m�F����B
    '�@����悪���e�[��DB�����}�X�^�ɓo�^����Ă��邩�m�F����B
    '   �@DB�F      ���e�[��DB�����}�X�^�im_torihikisaki�j
    '   �@�L�[�F        �����R�[�h(torihikisaki_cd)                           CSV�t�@�C������擾
    '           ������itorikesi_date�j                         ���t�������Ă��邩
    '�@������ɓ��t�������Ă���ꍇ�́A�����~�ƌ��Ȃ��A�G���[�Ƃ���B�@2005/04/19�ǉ�

'********** 2005/05/11 �y�c �폜 **********��
'DB�ڑ��Ɋւ��鏈�����폜
'    Dim strSql            As String
'    Dim objCDBAccess      As Object
'    Dim objRS             As Object
'********** 2005/05/11 �y�c �폜 **********��
    
    Dim strMethodCallInfo As String

    On Error GoTo DefectErrorHandler
    
    FncCheckDB_Torihikisaki = False
        
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncCheckDB_Torihikisaki(" & strTorihikisaki & "," & intNowLine & ")"

'********** 2005/05/11 �y�c �폜 **********��
'���̊֐����ł̐ؒf���R�[�h�Z�b�g�擾�����폜
'    strSql = "SELECT torikesi_date FROM v_torihikisaki WITH(readcommitted) WHERE torihikisaki_cd = '" & strTorihikisaki & "'"
    
'    'CDBAccess�C���X�^���X����
'    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
'    'SQL���s
'        objRS = objCDBAccess.GetDCRecordset("VSWF", strSql)
'********** 2005/05/11 �y�c �폜 **********��

'********** 2005/05/11 �y�c �ǉ� **********��
'���ϐ��̐ؒf���R�[�h�Z�b�g���g�p���G���[�`�F�b�N
    With pTChkTorihikisakiRS
'********** 2005/05/17 �y�c �ύX **********��
'********** 2005/05/18 �y�c �ύX **********��
'        If Not .EOF Then
        .Filter = adFilterNone
       If Not .BOF Then
'********** 2005/05/18 �y�c �ύX **********��
            .MoveFirst
            .Find "torihikisaki_cd = '" & strTorihikisaki & "'"
        End If
'********** 2005/05/17 �y�c �ύX **********��
    End With
'********** 2005/05/11 �y�c �ǉ� **********��

'********** 2005/05/12 �y�c �ύX **********��
    ' ����悪���݂��Ȃ��ꍇ�̓G���[
'    If objRS.EOF = True Then
    If pTChkTorihikisakiRS.EOF = True Then
        FncSetError (intNowLine, enmTenColNo.Torihikisaki, PARAM_NOTEXISTDB
'        GoTo Proc_End
        Exit Function
    End If

    '���t�`�F�b�N�͐ؒf���R�[�h�擾���ɍς܂��Ă���ׂ����ł̓`�F�b�N���Ȃ�
'    Dim strTorikesiDate
'    strTorikesiDate = objRS.Fields("torikesi_date")
'    ' ���t���󗓂łȂ��ꍇ�̓G���[
'    If Trim(strTorikesiDate) <> "" Then
'        FncSetError (intNowLine, enmTenColNo.Torihikisaki, PARAM_TORIHIKISTOP
'        GoTo Proc_End
'    End If
'********** 2005/05/12 �y�c �ύX **********��
    
    FncCheckDB_Torihikisaki = True
        
'********** 2005/05/11 �y�c �ύX **********��
'DB�ڑ��Ɋւ��鏈�����폜
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
'********** 2005/05/11 �y�c �ύX **********��
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckDB_BaikaHenkouRiyuu
'*�y�����T�v�z
'*?    �����ύX���R���L������Ԃ��܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strBaika_HenkouRiyuu ; String �����ύX���R������킷������
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckDB_BaikaHenkouRiyuu; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*M    2005/05/11; �y�c�@�S��(USC); �p�t�H�[�}���X����̈פ���̊֐����ł̐ؒf���R�[�h�擾����ߑ��ϐ��̐ؒf���R�[�h�𗘗p����l�C��;
'*< Comment End >**************************************************'
Private Function FncCheckDB_BaikaHenkouRiyuu(strBaika_HenkouRiyuu As String, intNowLine As Integer)
        
    'c.�����ύX���R���L�����m�F����B
    '�@�����ύX���R�����e�[��DB���̃}�X�^�ɓo�^����Ă��邩�m�F����B
    '   �@DB�F      ���e�[��DB���̃}�X�^�iv_meisyou�j
    '   �@�L�[�F        KEY�敪�ikey_kbn�j                          �u25�v
    '           KEY�R�[�h�ikey_cd�j                         CSV�t�@�C������擾
    
'********** 2005/05/11 �y�c �폜 **********��
'���̊֐����ł̐ؒf���R�[�h�Z�b�g�擾�����폜
'    Dim strSql As String
    
'    FncCheckDB_BaikaHenkouRiyuu = True

    

'    'SQL���쐬�i�����FKEY�敪=025�A�����ύX���R�R�[�h�j
'    strSql = "SELECT key_cd FROM v_meisyou WITH(readcommitted) "
'    strSql = strSql & "WHERE key_kbn = '" & KEYCD_BAIKARIYUU & "' And "
'    strSql = strSql & "key_cd = '" & strBaika_HenkouRiyuu & "'"
    
'    If FncCheckDB_IsRecordExists("VSWF", strSql) = False Then
'        FncSetError (intNowLine, enmTenColNo.BaikaRiyuu, PARAM_NOTEXISTDB
'        FncCheckDB_BaikaHenkouRiyuu = False
'    End If
'********** 2005/05/11 �y�c �폜 **********��
 
'********** 2005/05/11 �y�c �ǉ� **********��
'���ϐ��̐ؒf���R�[�h�Z�b�g���g�p���G���[�`�F�b�N
    With pTChkBaikaHenkouRS
'********** 2005/05/17 �y�c �ύX **********��
'********** 2005/05/18 �y�c �ύX **********��
'        If Not .EOF Then
       .Filter = adFilterNone
        If Not .BOF Then
'********** 2005/05/18 �y�c �ύX **********��
            .MoveFirst
            .Find "key_cd = '" & strBaika_HenkouRiyuu & "'"
        End If
'********** 2005/05/17 �y�c �ύX **********��
    End With
 
    If pTChkBaikaHenkouRS.EOF Then
        FncSetError (intNowLine, enmTenColNo.BaikaRiyuu, PARAM_NOTEXISTDB
    End If
    
    FncCheckDB_BaikaHenkouRiyuu = Not pTChkBaikaHenkouRS.EOF
'********** 2005/05/11 �y�c �폜 **********��

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckDB_CenterClass
'*�y�����T�v�z
'*?    �Z���^�[�敪���L������Ԃ��܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strCenterClass ; String �Z���^�[�敪������킷������
'*P    strCenterCode ; String �Z���^�[�R�[�h������킷������
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckDB_CenterClass; String;0:���� 1:�Z���^�[�敪�ُ� 2:�Z���^�[�R�[�h�ُ�
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*M    2005/05/12; �y�c�@�S��(USC); �p�t�H�[�}���X����̈פ���̊֐����ł̐ؒf���R�[�h�擾����ߑ��ϐ��̐ؒf���R�[�h�𗘗p����l�C��;
'*M    2011/02/16; �c�ƃV�X�e��/�a�c; �Z���^�[�R�[�h�ǉ�;
'*M    2015/09/15; ���ʃV�X�e��G/�a�c; TC�V�X�e���\�z;
'*< Comment End >**************************************************'
Private Function FncCheckDB_CenterClass(strCenterClass As String, _
                                  strCenterCode As String, _
                                  intNowLine As Integer) As String

    'e.�Z���^�[�敪���L�����m�F����B
    '�@�Z���^�[�敪��WF���̃}�X�^�ɓo�^����Ă��邩�m�F����B
    '   �@DB�F      WF���̃}�X�^�im_wf_meisyou�j
    '   �@�L�[�F    KEY�敪�ikey_kbn�j                      �uM09�v
    '           KEY�R�[�h�ikey_cd�j                         CSV�t�@�C������擾
    
    '�Z���^�[�敪��"1"�̏ꍇ�A�Z���^�[�敪�ƃZ���^�[�R�[�h
    '��A�����������̂��AKEY�R�[�h�ɂȂ�܂��B
    '�܂�A
    '�@�E�Z���^�[�敪="1"�̏ꍇ
    '�@�@KEY�敪�ikey_kbn�j�F�uM09�v
    '�@�@KEY�R�[�h�ikey_cd�j�F�Z���^�[�敪�{�Z���^�[�R�[�h
    '�@�E�Z���^�[�敪="1"�ȊO�̏ꍇ
    '�@�@KEY�敪�ikey_kbn�j�F�uM09�v
    '�@�@KEY�R�[�h�ikey_cd�j�F�Z���^�[�敪

'********** 2005/05/12�y�c �폜 **********��
'���̊֐����ł̐ؒf���R�[�h�Z�b�g�擾�����폜
'    Dim strSql            As String
    
'    strSql = "SELECT key_cd FROM m_wf_meisyou WITH(readcommitted) "
'    strSql = strSql & "WHERE key_kbn = '" & KEYCD_CENTER_KBN & "' "

'    If strCenterClass = "1" Then
'    '�@�E�Z���^�[�敪="1"�̏ꍇ�AKEY�R�[�h��"10151"�Ń`�F�b�N
'        strSql = strSql & " AND key_cd = '" & strCenterClass & CENTERCD_KODAMA & "'"
'    Else
'    '�@�E�Z���^�[�敪="1"�ȊO�̏ꍇ
'        strSql = strSql & " AND key_cd = '" & strCenterClass & "'"
'    End If
'
'    If FncCheckDB_IsRecordExists("VSWF", strSql) = False Then
'        FncCheckDB_CenterClass = False
'        FncSetError (intNowLine, enmTenColNo.CenterKbn, PARAM_NOTEXISTDB
'    Else
'        FncCheckDB_CenterClass = True
'    End If
'********** 2005/05/12�y�c �폜 **********��

'********** 2005/05/12�y�c �ǉ� **********��
'���ϐ��̐ؒf���R�[�h�Z�b�g���g�p���G���[�`�F�b�N

    Dim strKeyCode As String '�`�F�b�N����KEY�R�[�h
    Dim strHanteiKbn As String '�`�F�b�N����敪���Z�b�g
    strHanteiKbn = "0"  '����(�����l���Z�b�g)
    
    '�@�E�Z���^�[�敪="1"�̏ꍇ�AKEY�R�[�h��"10151"(�Z���^�[�敪�{�Z���^�[�R�[�h)�Ń`�F�b�N
    If strCenterClass = "1" Then
        'strKeyCode = "1" & CENTERCD_KODAMA
        strKeyCode = "1" & strCenterCode
    '2015/09/15 wadak7 ins
    '�Z���^�[�敪="3"�̏ꍇ�AKEY�R�[�h��"30653"(�Z���^�[�敪�{�Z���^�[�R�[�h)�Ń`�F�b�N
    ElseIf strCenterClass = "3" Then
        strKeyCode = "3" & strCenterCode
    Else
    '�@�E�Z���^�[�敪="1"�ȊO�̏ꍇ�AKEY�R�[�h�̓Z���^�[�敪�Ń`�F�b�N
        strKeyCode = strCenterClass
    End If
    
    With pTChkCenterRS
'********** 2005/05/17 �y�c �ύX **********��
'********** 2005/05/18 �y�c �ύX **********��
'        If Not .EOF Then
       .Filter = adFilterNone
        If Not .BOF Then
'********** 2005/05/18 �y�c �ύX **********��
            .MoveFirst
            .Find "key_cd = '" & strKeyCode & " '"
        End If
'********** 2005/05/17 �y�c �ύX **********��
    End With
 
    If pTChkCenterRS.EOF Then
        '2015/09/19 wadak7 upd
        'If strCenterClass = "1" Then
        If strCenterClass = "1" Or strCenterClass = "3" Then
            FncSetError (intNowLine, enmModColNo.centerCD, PARAM_NOTEXISTDB
            strHanteiKbn = "2"  '�Z���^�[�R�[�h�ُ�
        Else
            FncSetError (intNowLine, enmTenColNo.CenterKbn, PARAM_NOTEXISTDB
            strHanteiKbn = "1"  '�Z���^�[�敪�ُ�
        End If
    End If
    
    'FncCheckDB_CenterClass = Not pTChkCenterRS.EOF
    FncCheckDB_CenterClass = strHanteiKbn

'********** 2005/05/12�y�c �ǉ� **********��

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckDB_Sakituke
'*�y�����T�v�z
'*?    �����}�X�^�̐�t�ύX���ɓ����̕ύX����񂪖��������m�F����B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strKeyCode ; String ��L�[������킷������
'*P    arrCSVRecord ; String ���R�[�h������킷������
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckDB_Sakituke; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*M    2005/05/12; �y�c�@�S��(USC); �p�t�H�[�}���X����̈פ���ϐ��̐ؒf���R�[�h�𗘗p���X�R�[�h��W�J����l�C��;
'*M    2005/05/13; �c���@���l(USC); �p�t�H�[�}���X����̈פ�ύX���ɂ��Ă����ϐ��̐ؒf���R�[�h�𗘗p����l�C��;
'*< Comment End >**************************************************'
Private Function FncCheckDB_Sakituke(strKeyCode As String, arrCSVRecord() As String, intNowLine As Integer)

'f.�����}�X�^�̐�t�ύX���ɓ����̕ύX����񂪖������m�F����B
'�@�����}�X�^�̐�t�ύX���i�����ύX���A�d���ύX���A�X�܃Z���^�[�R�[�h�ύX���j
'�@�ɓ����̕ύX���������Ă��Ȃ������m�F����B                                           2005/04/19 �ǉ�
'   �@DB�F      �����}�X�^ �����ύX���it_siire_henkou�j
'           �����}�X�^ �d���ύX���it_baika_henkou�j
'           �����}�X�^ �X�܃Z���^�[�R�[�h�ύX���it_tenpo_hattyuu_cycle_henkou�j
'   �@�L�[�F        ��L�[�ikey_cd�j                            1.(2)b.(b) KEY���1����擾
'           �X�R�[�h�iten_cd�j                          1.(2)a. �W�J���Ă���X�R�[�h
'           �ύX���ihenkou_date�j                           CSV�t�@�C������擾

    Dim strSQL            As String
    Dim strTenCode        As String
    Dim strTenCodeLine    As String
    
    FncCheckDB_Sakituke = True
    
    strTenCode = arrCSVRecord(enmTenColNo.TenCd)
    
'********** 2005/05/12�y�c �폜 **********��
'    '�W�J���Ă���X�R�[�h���擾����SQL���쐬
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
'********** 2005/05/12�y�c �폜 **********��

'********** 2005/05/12�y�c �ǉ� **********��
    If strTenCode = ALL_TEN_CD Or strTenCode = ALL_JIG_CD Then
        '�S�X�A�S���Ƃ̏ꍇ
        strTenCodeLine = pTTenCodeList
        
'********** 2005/05/20 �y�c �ǉ� **********��
    ElseIf strTenCode = HON_TEN_CD Then
        '�{���̏ꍇ
        strTenCodeLine = "'" & HON_TEN_CD & "'"
'********** 2005/05/20 �y�c �ǉ� **********��

    ElseIf Len(strTenCode) = 4 Then
        '�e�X
        strTenCodeLine = strTenCode
    Else
        '�e����
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

'********** 2005/05/12�y�c �ǉ� **********��
    
    ' �����ύX���ɑ΂��₢���킹
    If arrCSVRecord(enmTenColNo.Baika) <> "" _
        Or arrCSVRecord(enmTenColNo.BaikaRiyuu) <> "" Then
        
'********** 2005/05/13�c�� �ύX **********��
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
'********** 2005/05/17 �y�c �폜 **********��
'            .MoveFirst
'********** 2005/05/17 �y�c �폜 **********��
            .Filter = "key_cd = '" & strKeyCode & "' AND henkou_date = '" & arrCSVRecord(enmTenColNo.HenkouDate) & "'"
            If Not .EOF Then
                Do Until .EOF
                    If InStr(strTenCodeLine, .Fields("ten_cd")) > 0 Then
'********** 2005/06/15 �k�V **********��
'                        FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_REPDATEBAIKA
'                        FncCheckDB_Sakituke = False
'                        Exit Do
'������t�ύX��񂪊DB�Ɋ������Ă��G���[�ɂ����A�o�^�敪�2��Ńf�[�^�쐬(���͍X�V)�ΏۂƂ���
                        arrTourokuKbn2(intNowLine - 1, conColBaika) _
                        = arrTourokuKbn2(intNowLine - 1, conColBaika) _
                        & "'" & .Fields("ten_cd") & "',"
'********** 2005/06/15 �k�V **********��
                    End If
                    .MoveNext
                Loop
'********** 2005/06/15 �k�V **********��
'�o�^�敪=2�̓X�R�[�h�z��̉E�[�̃J���}����������
                If Len(arrTourokuKbn2(intNowLine - 1, conColBaika)) > 0 Then
                    arrTourokuKbn2(intNowLine - 1, conColBaika) _
                    = Left(arrTourokuKbn2(intNowLine - 1, conColBaika), Len(arrTourokuKbn2(intNowLine - 1, conColBaika)) - 1)
                End If
'********** 2005/06/15 �k�V **********��
            End If
            .Filter = adFilterNone
        End With
'********** 2005/05/13�c�� �ύX **********��

    End If
    
    ' �d���ύX���ɑ΂��₢���킹

'********** 2005/05/13�c�� �ύX **********��
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
'********** 2005/05/17 �y�c �폜 **********��
'            .MoveFirst
'********** 2005/05/17 �y�c �폜 **********��
            .Filter = "key_cd = '" & strKeyCode & "' AND henkou_date = '" & arrCSVRecord(enmTenColNo.HenkouDate) & "'"
            If Not .EOF Then
                Do Until .EOF
                    If InStr(strTenCodeLine, .Fields("ten_cd")) > 0 Then
'********** 2005/06/15 �k�V **********��
'                        FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_REPDATESIIRE
'                        FncCheckDB_Sakituke = False
'                        Exit Do
'�d����t�ύX��񂪊DB�Ɋ������Ă��G���[�ɂ����A�o�^�敪�2��Ńf�[�^�쐬(���͍X�V)�ΏۂƂ���
                        arrTourokuKbn2(intNowLine - 1, conColSiire) _
                        = arrTourokuKbn2(intNowLine - 1, conColSiire) _
                        & "'" & .Fields("ten_cd") & "',"
'********** 2005/06/15 �k�V **********��
                    End If
                    .MoveNext
                Loop
'********** 2005/06/15 �k�V **********��
'�o�^�敪=2�̓X�R�[�h�z��̉E�[�̃J���}����������
                If Len(arrTourokuKbn2(intNowLine - 1, conColSiire)) > 0 Then
                    arrTourokuKbn2(intNowLine - 1, conColSiire) _
                    = Left(arrTourokuKbn2(intNowLine - 1, conColSiire), Len(arrTourokuKbn2(intNowLine - 1, conColSiire)) - 1)
                End If
'********** 2005/06/15 �k�V **********��
            End If
            .Filter = adFilterNone
        End With
    End If
'********** 2005/05/13�c�� �ύX **********��
    
    ' �X�܃Z���^�[�R�[�h�ύX���ɑ΂��₢���킹

'********** 2005/05/13�c�� �ύX **********��
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
'********** 2005/05/17 �y�c �폜 **********��
'            .MoveFirst
'********** 2005/05/17 �y�c �폜 **********��
            .Filter = "key_cd = '" & strKeyCode & "' AND henkou_date = '" & arrCSVRecord(enmTenColNo.HenkouDate) & "'"
            If Not .EOF Then
                Do Until .EOF
                    If InStr(strTenCodeLine, .Fields("ten_cd")) > 0 Then
'********** 2005/06/15 �k�V **********��
'                        FncSetError (intNowLine, enmTenColNo.HenkouDate, PARAM_REPDATETENPO
'                        FncCheckDB_Sakituke = False
'                        Exit Do
'�Z���^�[��t�ύX��񂪊DB�Ɋ������Ă��G���[�ɂ����A�o�^�敪�2��Ńf�[�^�쐬(���͍X�V)�ΏۂƂ���
                        arrTourokuKbn2(intNowLine - 1, conColCenter) _
                        = arrTourokuKbn2(intNowLine - 1, conColCenter) _
                        & "'" & .Fields("ten_cd") & "',"
'********** 2005/06/15 �k�V **********��
                    End If
                    .MoveNext
                Loop
'********** 2005/06/15 �k�V **********��
'�o�^�敪=2�̓X�R�[�h�z��̉E�[�̃J���}����������
                If Len(arrTourokuKbn2(intNowLine - 1, conColCenter)) > 0 Then
                    arrTourokuKbn2(intNowLine - 1, conColCenter) _
                    = Left(arrTourokuKbn2(intNowLine - 1, conColCenter), Len(arrTourokuKbn2(intNowLine - 1, conColCenter)) - 1)
                End If
'********** 2005/06/15 �k�V **********��
            End If
            .Filter = adFilterNone
        End With
    End If
'********** 2005/05/13�c�� �ύX **********��
    
End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckDB_IsRecordExists
'*�y�����T�v�z
'*?    �w�肳�ꂽSQL�������s���A�e�[�u�����ɊY�����郌�R�[�h�����݂��邩�ǂ�����Ԃ��܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strUDL_Name ; String �ڑ�������
'*P    strSQLInstruction ; String SQL��
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckDB_IsRecordExists; Boolean;���R�[�h�����݂���:True,���݂��Ȃ�:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheckDB_IsRecordExists(strUDL_Name As String, strSQLInstruction As String) As Boolean

    Dim objRs             As Object
    Dim objCDBAccess      As Object
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncCheckDB_IsRecordExists(" & strUDL_Name & "," & strSQLInstruction & ")"
        
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")

    'SQL���s
        objRs = objCDBAccess.GetDCRecordset(strUDL_Name, strSQLInstruction)
    
    ' �f�[�^���擾�ł��Ȃ���΁A�e�[�u�����ɂ͊Y�����R�[�h�͑��݂��Ȃ�
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
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'

'*�y�v���V�[�W�����z
'*!    FncTenTourokuJig

'*�y�����T�v�z
'*?    ���ƃR�[�h�ɕR�t�����X�R�[�h�̃f�[�^���擾���āA�X�R�[�h����
'*?    DB�Ɏ�荞��

'*�y�p�����[�^�z
'*P    strErrMsg;        string; �G���[���b�Z�[�W;
'*P    intN;             int; ���Ԗڂ̃��R�[�h��;
'*P    strTenCd;         string; ���ƃR�[�h;

'*�y�߂�l�z
'*R    Boolean; TRUE:�����AFALSE:���s;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/20; ����(���j�V�X�e���������); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncTenTourokuJig(ByRef strErrMsg, ByVal intN, ByVal strTenCd) As Boolean
    Dim sql, objCDB, intSelNum, blnRet
    Dim rsSel As ADODB.Recordset
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTourokuJig(" & CStr(intN) + 1 & "�s�ځF�y���ƃR�[�h�F" & strTenCd & "�z)"
    
    FncTenTourokuJig = False

    '���ƃR�[�h����X�R�[�h���擾����B
        objCDB = CreateObject("COMDBUtil.CDBAccess")    'CDBAccess���C���X�^���X��

    'SQL���쐬�i�����F���ʋ敪=5�A�X����NULLor�ꊇ�o�^���ȍ~�AKEY�敪=W02�A�ݒ���e�̂S���ڂ�1�j
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
    '�S�X�A�S���Ƃ̏ꍇ�́A�ȉ��̏����͕K�v�Ȃ��B
    If strTenCd <> ALL_TEN_CD And strTenCd <> ALL_JIG_CD Then
        sql = sql & " AND A.jigyou_cd = '" & strTenCd & "' " & vbCrLf
    End If
    
        rsSel = objCDB.GetDCRecordset("VSWF", sql)
        objCDB = Nothing
    If rsSel.EOF = True Then    '�Ώۂ̓X�R�[�h�������ꍇ�̓G���[
        strErrMsg = CStr(intN + 1) & "�s�ځF�y���ƃR�[�h�F" & strTenCd & "�z�̓X�R�[�h�͂P��������܂���ł����B"
        rsSel.Close
            rsSel = Nothing
        Exit Function
    End If
    
    '�X���ޖ��Ƀf�[�^�����i�I��WF��DB�Ɏ�荞�ށiFncTenTourokuTen�֐����Ăԁj�B
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
    
    'DB�N���[�Y
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
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'

'*�y�v���V�[�W�����z
'*!    FncTenTourokuTen

'*�y�����T�v�z
'*?    �w�肳�ꂽ�X�R�[�h�̃f�[�^��DB�Ɏ�荞��

'*�y�p�����[�^�z
'*P    strErrMsg;        string; �G���[���b�Z�[�W;
'*P    intN;             int; ���Ԗڂ̃��R�[�h��;
'*P    strTenCd;         string; �X�R�[�h;

'*�y�߂�l�z
'*R    Boolean; TRUE:�����AFALSE:���s;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/20; ����(���j�V�X�e���������); �V�K�쐬;
'*M    2018/12/14; ������(��A)              ; �ꊇ�o�^���A�������[�U�[�̂ݎ��s�ł���;
'*< Comment End >**************************************************'
Private Function FncTenTourokuTen(ByRef strErrMsg, ByVal intN, ByVal strTenCd) As Boolean

    Dim sql, objCDB, intUdtNum, intInsNum, intSelNum
    Dim rsSel
    Dim arrParamSel() As Object                             ' SQL�p�����[�^�[�l
    Dim arrParamIns() As Object                             ' SQL�p�����[�^�[�l
    Dim arrParamUpd() As Object                             ' SQL�p�����[�^�[�l
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncTenTourokuTen(" & CStr(intN) + 1 & "�s�ځF�y�X�R�[�h�F" & strTenCd & "�z)"
    
    FncTenTourokuTen = False

    '�擾�����f�[�^�����ɁADB���X�V����B
        objCDB = CreateObject("COMDBUtil.CDBAccess")    'CDBAccess���C���X�^���X��
    
    '�ꊇ�o�^���A�������[�U�[�̂ݎ��s�ł���
    Call LockTableIkkatu(objCDB)

    '�����A�����ύX���R�̂ǂꂩ�����邩�`�F�b�N����B
    If paryCsvData(intN)(enmTenColNo.Baika) <> "" Or paryCsvData(intN)(enmTenColNo.BaikaRiyuu) <> "" Then
       '���i�I��DB(�����ύX���)�̍X�V
       '�f�[�^�̑��݃`�F�b�N���s���B
        Call SubSqlSelTen(intN, strTenCd, sql, arrParamSel, TBL_BAI_HEN)
            rsSel = objCDB.GetDCRecordset("VSWF", sql, arrParamSel)
        
       '�f�[�^�������ꍇ
        If rsSel.EOF = True Then    'insert����
            Call SubSqlInsTen(intN, strTenCd, sql, arrParamIns, TBL_BAI_HEN) 'SQL���쐬
            intInsNum = objCDB.ExecuteSql("VSWF", sql, arrParamIns) 'SQL���s
            If intInsNum <= 0 Then
                strErrMsg = CStr(intN + 1) & "�s�ځF�y�X�R�[�h�F" & strTenCd & "�z��"
                strErrMsg = strErrMsg & "�����ύX����ǉ�����̂Ɏ��s���܂����B"
                    objCDB = Nothing
                    rsSel = Nothing
                Exit Function
            End If
            pflgBaikaToukbn = True  '���i�I��DB(WF�Ǘ�����)�̔����ύX�̓o�^�敪�X�V�p�t���O
        '�f�[�^������ꍇ
        Else                        'Update����
            Call SubSqlUpdTen(intN, strTenCd, sql, arrParamUpd, TBL_BAI_HEN) 'SQL���쐬
            intUdtNum = objCDB.ExecuteSql("VSWF", sql, arrParamUpd) 'SQL���s
            If intUdtNum <= 0 Then
                strErrMsg = CStr(intN + 1) & "�s�ځF�y�X�R�[�h�F" & strTenCd & "�z��"
                strErrMsg = strErrMsg & "�����ύX�����X�V����̂Ɏ��s���܂����B"
                    objCDB = Nothing
                rsSel.Close
                    rsSel = Nothing
                Exit Function
            End If
            rsSel.Close
        End If
            rsSel = Nothing
    End If
        
    '�����A�����A�����P�ʂ̂ǂꂩ�����邩�`�F�b�N����B
    If paryCsvData(intN)(enmTenColNo.Genka) <> "" _
        Or paryCsvData(intN)(enmTenColNo.Torihikisaki) <> "" _
        Or paryCsvData(intN)(enmTenColNo.Hacchuutani) <> "" Then
        '���i�I��DB(�d���ύX���)�̍X�V
        '�f�[�^�̑��݃`�F�b�N���s���B
        Call SubSqlSelTen(intN, strTenCd, sql, arrParamSel, TBL_SII_HEN)
            rsSel = objCDB.GetDCRecordset("VSWF", sql, arrParamSel)
    
        '�f�[�^�������ꍇ
        If rsSel.EOF = True Then   'insert����
            Call SubSqlInsTen(intN, strTenCd, sql, arrParamIns, TBL_SII_HEN) 'SQL���쐬
            intInsNum = objCDB.ExecuteSql("VSWF", sql, arrParamIns) 'SQL���s
            If intInsNum <= 0 Then
                strErrMsg = CStr(intN + 1) & "�s�ځF�y�X�R�[�h�F" & strTenCd & "�z��"
                strErrMsg = strErrMsg & "�d���ύX����ǉ�����̂Ɏ��s���܂����B"
                    objCDB = Nothing
                    rsSel = Nothing
                Exit Function
            End If
            
            pflgSiireToukbn = True  '���i�I��DB(WF�Ǘ�����)�̎d���ύX�̓o�^�敪�X�V�p�t���O
           
        Else                        'Update����
            Call SubSqlUpdTen(intN, strTenCd, sql, arrParamUpd, TBL_SII_HEN) 'SQL���쐬
            intUdtNum = objCDB.ExecuteSql("VSWF", sql, arrParamUpd) 'SQL���s
            If intUdtNum <= 0 Then
                strErrMsg = CStr(intN + 1) & "�s�ځF�y�X�R�[�h�F" & strTenCd & "�z��"
                strErrMsg = strErrMsg & "�d���ύX�����X�V����̂Ɏ��s���܂����B"
                    objCDB = Nothing
                rsSel.Close
                    rsSel = Nothing
                Exit Function
            End If
        End If
            rsSel = Nothing
    End If
       
    '�Z���^�[�敪�����邩�`�F�b�N����B
    If paryCsvData(intN)(enmTenColNo.CenterKbn) <> "" Then
        '���i�I��DB(�Z���^�[�R�[�h�ύX���)�̍X�V
        '�f�[�^�̑��݃`�F�b�N���s���B
        Call SubSqlSelTen(intN, strTenCd, sql, arrParamSel, TBL_CEN_HEN)
            rsSel = objCDB.GetDCRecordset("VSWF", sql, arrParamSel)
            
        '�f�[�^�������ꍇ
        If rsSel.EOF = True Then   'insert����
            Call SubSqlInsTen(intN, strTenCd, sql, arrParamIns, TBL_CEN_HEN) 'SQL���쐬
            intInsNum = objCDB.ExecuteSql("VSWF", sql, arrParamIns) 'SQL���s
            If intInsNum <= 0 Then
                strErrMsg = CStr(intN + 1) & "�s�ځF�y�X�R�[�h�F" & strTenCd & "�z��"
                strErrMsg = strErrMsg & "�Z���^�[�R�[�h�ύX����ǉ�����̂Ɏ��s���܂����B"
                    objCDB = Nothing
                    rsSel = Nothing
                Exit Function
            End If
            
            pflgCenterToukbn = True '���i�I��DB(WF�Ǘ�����)�̃Z���^�[�R�[�h�ύX�̓o�^�敪�X�V�p�t���O
            
        Else                        'Update����
            Call SubSqlUpdTen(intN, strTenCd, sql, arrParamUpd, TBL_CEN_HEN) 'SQL���쐬
            intUdtNum = objCDB.ExecuteSql("VSWF", sql, arrParamUpd) 'SQL���s
            If intUdtNum <= 0 Then
                strErrMsg = CStr(intN + 1) & "�s�ځF�y�X�R�[�h�F" & strTenCd & "�z��"
                strErrMsg = strErrMsg & "�Z���^�[�R�[�h�ύX�����X�V����̂Ɏ��s���܂����B"
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
    
    'DB�N���[�Y
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
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'

'*�y�v���V�[�W�����z
'*!    FncSetWFKanri

'*�y�����T�v�z
'*?    WF�Ǘ����ڂ��ŐV���ɍX�V����B�i���R�[�h���j

'*�y�p�����[�^�z
'*P    strErrMsg;        string; �G���[���b�Z�[�W;
'*P    intN;             int; ���Ԗڂ̃��R�[�h��;
'*P    intType;          int; �������~or�X�ʏ��;

'*�y�߂�l�z
'*R    Boolean; TRUE:�����AFALSE:���s;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/20; ����(���j�V�X�e���������); �V�K�쐬;
'*M    2005/04/26; ����(���j�V�X�e���������); �֐����ύX�iFncUpdWFKanri��FncSetWFKanri�j;
'*M    2005/04/26; ����(���j�V�X�e���������); �����ǉ��iintType�j;
'*< Comment End >**************************************************'
'********** 2005/04/26 ���� �ύX **********��
Private Function FncSetWFKanri(ByRef strErrMsg, ByVal intN, ByVal intType) As Boolean
'Private Function FncUpdWFKanri(ByRef strErrMsg, ByVal intN) As Boolean
'********** 2005/04/26 ���� �ύX **********��
    Dim sql, objCDB, intUdtNum
    
'********** 2005/04/26 ���� �ύX **********��
    Dim strSelSql, strInsSql, strUpdSql, lngP_Cnt
    Dim arrParam() As Object
'    Dim arrParamUpd() As Object
'********** 2005/04/26 ���� �ύX **********��

'********** 2005/04/26 ���� �ύX **********��
    FncSetWFKanri = False
'    FncUpdWFKanri = False
'********** 2005/04/26 ���� �ύX **********��
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
'********** 2005/04/26 ���� �ύX **********��
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncUpdWFKanri(" & CStr(intN) + 1 & "�s�ځF�yWF�Ǘ����ړo�^�G���[�z)"
'    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncUpdWFKanri(" & CStr(intN) + 1 & "�s�ځF�yWF�Ǘ����ڍX�V�G���[�z)"
'********** 2005/04/26 ���� �ύX **********��
    
        objCDB = CreateObject("COMDBUtil.CDBAccess")    'CDBAccess���C���X�^���X��
    
'********** 2005/04/26 ���� �ύX **********��
    lngP_Cnt = 0
    Erase arrParam
    Call SubSqlSelWFKanri(intN, strSelSql, arrParam, lngP_Cnt, intType)
    Call SubSqlInsWFKanri(intN, strInsSql, arrParam, lngP_Cnt, intType)
    Call SubSqlUpdWFKanri(intN, strUpdSql, arrParam, lngP_Cnt, intType)
    sql = "If (" & strSelSql & ") = 0 " & vbCrLf        '�f�[�^�����ɑ��݂��邩�`�F�b�N
    sql = sql & "BEGIN " & strInsSql & " END " & vbCrLf '���݂��Ȃ��ꍇ�́A�ǉ�
    sql = sql & "ELSE " & vbCrLf
    sql = sql & "BEGIN " & strUpdSql & " END " & vbCrLf '���݂���ꍇ�́A�X�V
    intUdtNum = objCDB.ExecuteSql("VSWF", sql, arrParam)
'    Call SubSqlUpdWFKanri(intN, sql, arrParamUpd)
'    intUdtNum = objCDB.ExecuteSql("VSWF", sql, arrParamUpd) 'SQL���s
'********** 2005/04/26 ���� �ύX **********��

    If intUdtNum <= 0 Then
'********** 2005/04/26 ���� �ύX **********��
        strErrMsg = CStr(intN + 1) & "�s�ځFWF�Ǘ����ڂ̃f�[�^��o�^����̂Ɏ��s���܂����B"
'        strErrMsg = CStr(intN + 1) & "�s�ځFWF�Ǘ����ڂ��X�V����̂Ɏ��s���܂����B"
'********** 2005/04/26 ���� �ύX **********��
            objCDB = Nothing
        Exit Function
    End If

'********** 2005/04/26 ���� �ύX **********��
    FncSetWFKanri = True
'    FncUpdWFKanri = True
'********** 2005/04/26 ���� �ύX **********��
    'DB�N���[�Y
        objCDB = Nothing
    'GetObjectContext.SetComplete
Exit Function

DefectErrorHandler:
    If Not objCDB Is Nothing Then
            objCDB = Nothing
    End If
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description

End Function

'*< UNISYSTEM >****************************************************'

'*�y�v���V�[�W�����z
'*!    SubHacSqlSelJig

'*�y�����T�v�z
'*?    SELECT���쐬�i�Y�����ƃR�[�h�̔������~���f�[�^�`�F�b�N�p�j

'*�y�p�����[�^�z
'*P    intN;            int;     ���Ԗڂ̃��R�[�h��;
'*P    strJigCd;        string;  ���ƃR�[�h;
'*P    sql;             string;  SQL���i�i�[�p�j;
'*P    arrPrm;          Object; �p�����[�^�i�i�[�p�j;
'*P    lngP_Cnt;        Long;    �p�����[�^���i�i�[�p�j;

'*�y�߂�l�z
'*R    �Ȃ�;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/25; ����(���j�V�X�e���������); �V�K�쐬;
'*< Comment End >**************************************************'
Private Sub SubHacSqlSelJig(ByVal intN, ByVal strJigCD, ByRef sql, ByRef arrPrm As Object, ByRef lngP_Cnt)

    sql = "SELECT COUNT(*) FROM t_syousen_jigyoubetu_settei WITH(UPDLOCK) " & vbCrLf
    sql = sql & "WHERE" & vbCrLf
    sql = sql & " kikaku_syurui_cd = ? AND " & vbCrLf
    sql = sql & " nen_ren_no = ? AND " & vbCrLf
    sql = sql & " eda_no = ? AND " & vbCrLf
    sql = sql & " jan_cd = ? AND " & vbCrLf
    sql = sql & " jigyou_cd = ? " & vbCrLf
    
    '�p�����[�^�Z�b�g
    ReDim Preserve arrPrm(lngP_Cnt + 4) As Object
    arrPrm(lngP_Cnt + 0) = MakeParam("@KIKAKU_SYURUI", DT_Char, 3, pstrKikakuSyurui)
    arrPrm(lngP_Cnt + 1) = MakeParam("@NEN_REN_NO", DT_Char, 10, pstrKikakuNenrenban)
    arrPrm(lngP_Cnt + 2) = MakeParam("@EDA_NO", DT_Char, 3, pstrKikakuEdaban)
    arrPrm(lngP_Cnt + 3) = MakeParam("@JAN_CD", DT_VarChar, 20, paryCsvData(intN)(enmHacColNo.janCD))
    arrPrm(lngP_Cnt + 4) = MakeParam("@JIGYOU_CD", DT_Char, 2, strJigCD)
    lngP_Cnt = lngP_Cnt + 5

End Sub

'*< UNISYSTEM >****************************************************'

'*�y�v���V�[�W�����z
'*!    SubHacSqlInsJig

'*�y�����T�v�z
'*?    INSERT���쐬�i�Y�����ƃR�[�h�̔������~���f�[�^�������ꍇ�p�j

'*�y�p�����[�^�z
'*P    intN;            int;     ���Ԗڂ̃��R�[�h��;
'*P    strJigCd;        string;  ���ƃR�[�h;
'*P    sql;             string;  SQL���i�i�[�p�j;
'*P    arrPrm;          Object; �p�����[�^�i�i�[�p�j;
'*P    lngP_Cnt;        Long;    �p�����[�^���i�i�[�p�j;

'*�y�߂�l�z
'*R    �Ȃ�;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/25; ����(���j�V�X�e���������); �V�K�쐬;
'*< Comment End >**************************************************'
Private Sub SubHacSqlInsJig(ByVal intN, ByVal strJigCD, ByRef sql, ByRef arrParam As Object, ByRef lngP_Cnt)

'********** 2005/05/09 �k�V **********��
    On Error GoTo Err_SubHacSqlInsJig
    Dim strSyobunHhouhou As String
    Dim strDaitaiSyouhinCd As String
'********** 2005/05/09 �k�V **********��
    
    'SQL���쐬
    sql = "INSERT INTO t_syousen_jigyoubetu_settei " & vbCrLf
    sql = sql & "(kikaku_syurui_cd, nen_ren_no, eda_no, jan_cd, jigyou_cd," & vbCrLf
    sql = sql & " hattyuu_kaisi_date, hattyuu_syuuryou_date," & vbCrLf
    sql = sql & " hattyuu_tyuusi_riyuu_kbn, syobun_houhou, daitai_syouhin_cd," & vbCrLf
    sql = sql & " dounyuu_baika, jyougen_baika, kagen_baika, " & vbCrLf
    sql = sql & " touroku_user_id, kousin_user_id, touroku_ymd_hms, kousin_ymd_hms) " & vbCrLf
    sql = sql & "SELECT TOP 1 ?,?,?,B.jan_cd,A.jigyou_cd, " & vbCrLf
    sql = sql & "A.hattyuu_kaisi_date, " & vbCrLf   '�����J�n��
    sql = sql & "CONVERT(char(8),'" & paryCsvData(intN)(enmHacColNo.SyuryouDate) & "') " & vbCrLf '�����I����
    sql = sql & ",?,?,?, " & vbCrLf             '�������~���R�敪�A�������@�A���JAN�R�[�h
    sql = sql & "A.dounyuu_baika, A.jyougen_baika, A.kagen_baika, " & vbCrLf
    sql = sql & "'" & pstrTantousyaId & "', null, GETDATE(), null " & vbCrLf
    sql = sql & "FROM v_jigyou_betu AS A WITH(readcommitted) " & vbCrLf
    sql = sql & "INNER JOIN v_key_1 AS B WITH(readcommitted) ON A.key_cd = B.key_cd " & vbCrLf
    sql = sql & "WHERE B.jan_cd = ? AND " & vbCrLf
    sql = sql & " A.jigyou_cd = ? " & vbCrLf
    
    '�p�����[�^�Z�b�g
    ReDim Preserve arrParam(lngP_Cnt + 7) As Object
    arrParam(lngP_Cnt + 0) = MakeParam("@KIKAKU_SYURUI", DT_Char, 3, pstrKikakuSyurui)
    arrParam(lngP_Cnt + 1) = MakeParam("@NEN_REN_NO", DT_Char, 10, pstrKikakuNenrenban)
    arrParam(lngP_Cnt + 2) = MakeParam("@EDA_NO", DT_Char, 3, pstrKikakuEdaban)

'********** 2005/05/09 �k�V **********��
    '�������~�敪(���͕K�{�Ȃ̂Ńu�����N�͂��蓾�Ȃ�)
'********** 2005/05/09 �k�V **********��
    arrParam(lngP_Cnt + 3) = MakeParam("@HACCHU_KBN", DT_Char, 2, paryCsvData(intN)(enmHacColNo.ChushiKbn))

'********** 2005/05/09 �k�V **********��
'���s�\�[�X�R�[�h�R�����g��
'    arrParam(lngP_Cnt + 4) = MakeParam("@SYOBUN_HOU", DT_Char, 1, paryCsvData(intN)(enmHacColNo.Syobun))
'    arrParam(lngP_Cnt + 5) = MakeParam("@DAITAI_CD", DT_Char, 13, paryCsvData(intN)(enmHacColNo.DaiJanCd))

'EXCEL�̏������@�E���JAN�R�[�h���u�����N�̏ꍇ�͊DB���l���擾����
    ''EXCEL�̏������@�E���JAN�R�[�h���u�����N�̏ꍇ�͊DB���l���擾����
    If paryCsvData(intN)(enmHacColNo.Syobun) = "" _
    Or paryCsvData(intN)(enmHacColNo.DaiJanCd) = "" Then
        
        '''�ϐ���`
        Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess") '���ʊ֐�(SQLServer�ڑ��p)
        Dim rsAdo As ADODB.Recordset            '���R�[�h�Z�b�g
        Dim strMethodCallInfo As String         '�G���[���
        Dim strSQL As String                    '��Ɨp������
        Const strUdl As String = "VSWF"         'SQLServer�ڑ��pUDL�t�@�C����
        
        '''�G���[�����쐬���Ă���
        'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.SubHacSqlInsJig(JAN:" & paryCsvData(intN)(enmHacColNo.janCD) & " ����:" & strJigCD & ")"
        
        '''SQL�����쐬����
        '''(TABLE: ���ƕʃr���[�C��L�[�}�X�^)
        '''(FIELDS: �������@�C���JAN�R�[�h)
        '''(KEY:JAN�R�[�h, ���ƃR�[�h)
        strSQL = _
        "SELECT v_jigyou_betu.syobun_houhou, v_jigyou_betu.daitai_syouhin_cd" & _
        "  FROM v_jigyou_betu WITH(READCOMMITTED)" & _
        " INNER JOIN v_key_1 WITH(READCOMMITTED) ON v_jigyou_betu.key_cd = v_key_1.key_cd" & _
        " WHERE v_key_1.jan_cd = '" & paryCsvData(intN)(enmHacColNo.janCD) & "'" & _
        "   AND v_jigyou_betu.jigyou_cd = '" & Trim(strJigCD) & "'"
        
        '''���ʊ֐��I�u�W�F�N�g���쐬����
        objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
        
        '''���R�[�h�Z�b�g���擾����
            rsAdo = objCDBAccess.GetDCRecordset(strUdl, strSQL)
        
        '''�������@�E���JAN�R�[�h��ϐ��ɕۊǂ���
        If rsAdo.EOF = False Then
            If IsNull(rsAdo("syobun_houhou")) = False Then
                strSyobunHhouhou = rsAdo("syobun_houhou")
            End If
            If IsNull(rsAdo("daitai_syouhin_cd")) = False Then
                strDaitaiSyouhinCd = rsAdo("daitai_syouhin_cd")
            End If
        End If
        
        '''�f�[�^�x�[�X�����
            rsAdo = Nothing
        objCDBAccess = Nothing
        'GetObjectContext.SetComplete
    End If
    
    ''�������@���Z�b�g����
        If paryCsvData(intN)(enmHacColNo.Syobun) = "" Then
            If strSyobunHhouhou <> "" Then
                arrParam(lngP_Cnt + 4) = MakeParam("@SYOBUN_HOU", DT_Char, 1, strSyobunHhouhou)
            Else
                arrParam(lngP_Cnt + 4) = MakeParam("@SYOBUN_HOU", DT_Char, 1, "0")
            End If
        Else
            arrParam(lngP_Cnt + 4) = MakeParam("@SYOBUN_HOU", DT_Char, 1, paryCsvData(intN)(enmHacColNo.Syobun))
        End If
        
        '''���JAN�R�[�h���Z�b�g����
        If paryCsvData(intN)(enmHacColNo.DaiJanCd) = "" Then
            If strDaitaiSyouhinCd <> "" Then
                arrParam(lngP_Cnt + 5) = MakeParam("@DAITAI_CD", DT_Char, 13, strDaitaiSyouhinCd)
            Else
                arrParam(lngP_Cnt + 5) = MakeParam("@DAITAI_CD", DT_Char, 13, Null)
            End If
        Else
            arrParam(lngP_Cnt + 5) = MakeParam("@DAITAI_CD", DT_Char, 13, paryCsvData(intN)(enmHacColNo.DaiJanCd))
        End If
'********** 2005/05/09 �k�V **********��
    
    arrParam(lngP_Cnt + 6) = MakeParam("@JAN_CD", DT_VarChar, 20, paryCsvData(intN)(enmHacColNo.janCD))
    arrParam(lngP_Cnt + 7) = MakeParam("@JIGYOU_CD", DT_Char, 2, strJigCD)
    lngP_Cnt = lngP_Cnt + 8

'********** 2005/05/09 �k�V **********��
    '�G���[���̏���
Ext_SubHacSqlInsJig:
    Exit Sub
Err_SubHacSqlInsJig:
    If Not rsAdo Is Nothing Then Set rsAdo = Nothing
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    Call ErrRaise(DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description)
    Resume Ext_SubHacSqlInsJig
'********** 2005/05/09 �k�V **********��

End Sub

'*< UNISYSTEM >****************************************************'

'*�y�v���V�[�W�����z
'*!    SubHacSqlUpdJig

'*�y�����T�v�z
'*?    UPDATE���쐬�i�Y�����ƃR�[�h�̔������~���f�[�^������ꍇ�p�j

'*�y�p�����[�^�z
'*P    intN;            int;     ���Ԗڂ̃��R�[�h��;
'*P    strJigCd;        string;  ���ƃR�[�h;
'*P    sql;             string;  SQL���i�i�[�p�j;
'*P    arrParamUpd;     Object; �p�����[�^�i�i�[�p�j;
'*P    lngParam_Cnt;    Long;    �p�����[�^���i�i�[�p�j;

'*�y�߂�l�z
'*R    �Ȃ�;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/25; ����(���j�V�X�e���������); �V�K�쐬;
'*< Comment End >**************************************************'
Private Sub SubHacSqlUpdJig(ByVal intN, ByVal strJigCD, ByRef sql, ByRef arrPrm As Object, ByRef lngP_Cnt)
    
    sql = "UPDATE t_syousen_jigyoubetu_settei " & vbCrLf
    sql = sql & "SET" & vbCrLf
    sql = sql & " hattyuu_syuuryou_date = " & vbCrLf                '�����I����
    sql = sql & "CONVERT(char(8),'" & paryCsvData(intN)(enmHacColNo.SyuryouDate) & "'), " & vbCrLf '�����I����
    
'********** 2005/05/09 �k�V **********��
    '�������~�敪(���͕K�{�Ȃ̂Ńu�����N�͂��蓾�Ȃ�)
'********** 2005/05/09 �k�V **********��
    If paryCsvData(intN)(enmHacColNo.ChushiKbn) <> "" Then
        sql = sql & " hattyuu_tyuusi_riyuu_kbn = ?," & vbCrLf
        ReDim Preserve arrPrm(lngP_Cnt) As Object
        arrPrm(lngP_Cnt) = MakeParam("@HAC_RIYUU_KBN", DT_Char, 2, paryCsvData(intN)(enmHacColNo.ChushiKbn))
        lngP_Cnt = lngP_Cnt + 1
    End If

'********** 2005/05/09 �k�V **********��
    '�������@(EXCEL���ڂ��u�����N�̏ꍇ�͉������Ȃ�(�DB�̒l�𐶂���))
'********** 2005/05/09 �k�V **********��
    If paryCsvData(intN)(enmHacColNo.Syobun) <> "" Then
        sql = sql & " syobun_houhou = ?," & vbCrLf
        ReDim Preserve arrPrm(lngP_Cnt) As Object
        arrPrm(lngP_Cnt) = MakeParam("@SYO_HOU", DT_Char, 1, paryCsvData(intN)(enmHacColNo.Syobun))
        lngP_Cnt = lngP_Cnt + 1
    End If

'********** 2005/05/06 �k�V **********��
    '���JAN�R�[�h(EXCEL���ڂ��u�����N�̏ꍇ�͉������Ȃ�(�DB�̒l�𐶂���))
'********** 2005/05/09 �k�V **********��
    If paryCsvData(intN)(enmHacColNo.DaiJanCd) <> "" Then
        sql = sql & " daitai_syouhin_cd = ?," & vbCrLf
        ReDim Preserve arrPrm(lngP_Cnt) As Object
        arrPrm(lngP_Cnt) = MakeParam("@DAITAI_CD", DT_Char, 13, paryCsvData(intN)(enmHacColNo.DaiJanCd))
        lngP_Cnt = lngP_Cnt + 1
    End If

    sql = sql & " kousin_user_id = '" & pstrTantousyaId & "'," & vbCrLf
    sql = sql & " kousin_ymd_hms = GETDATE() " & vbCrLf
    'WHERE��
    sql = sql & "WHERE " & vbCrLf
    sql = sql & " kikaku_syurui_cd = ? AND " & vbCrLf
    sql = sql & " nen_ren_no = ? AND " & vbCrLf
    sql = sql & " eda_no = ? AND " & vbCrLf
    sql = sql & " jan_cd = ? AND " & vbCrLf
    sql = sql & " jigyou_cd = ? " & vbCrLf
    
    '�p�����[�^�Z�b�g
    ReDim Preserve arrPrm(lngP_Cnt + 4) As Object
    arrPrm(lngP_Cnt + 0) = MakeParam("@KIKAKU_SYURUI", DT_Char, 3, pstrKikakuSyurui)
    arrPrm(lngP_Cnt + 1) = MakeParam("@NEN_REN_NO", DT_Char, 10, pstrKikakuNenrenban)
    arrPrm(lngP_Cnt + 2) = MakeParam("@EDA_NO", DT_Char, 3, pstrKikakuEdaban)
    arrPrm(lngP_Cnt + 3) = MakeParam("@JAN_CD", DT_Char, 13, paryCsvData(intN)(enmHacColNo.janCD))
    arrPrm(lngP_Cnt + 4) = MakeParam("@JIGYOU_CD", DT_Char, 2, strJigCD)
    lngP_Cnt = lngP_Cnt + 5

End Sub

'*< UNISYSTEM >****************************************************'

'*�y�v���V�[�W�����z
'*!    SubSqlSelTen

'*�y�����T�v�z
'*?    SELECT���쐬�i�X�ʏ��j

'*�y�p�����[�^�z
'*P    intN;            int;     ���Ԗڂ̃��R�[�h��;
'*P    strTenCd;        string;  �X�R�[�h;
'*P    sql;             string;  SQL���i�i�[�p�j;
'*P    arrParamSel;     Object; �p�����[�^�i�i�[�p�j;
'*P    strTbl;          string;  �X�V�e�[�u����;

'*�y�߂�l�z
'*R    �Ȃ�;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/20; ����(���j�V�X�e���������); �V�K�쐬;
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

'*�y�v���V�[�W�����z
'*!    SubSqlInsTen

'*�y�����T�v�z
'*?    INSERT���쐬�i�X�ʏ��j

'*�y�p�����[�^�z
'*P    intN;            int;     ���Ԗڂ̃��R�[�h��;
'*P    strTenCd;        string;  �X�R�[�h;
'*P    sql;             string;  SQL���i�i�[�p�j;
'*P    arrParamIns;     Object; �p�����[�^�i�i�[�p�j;
'*P    strTbl;          string;  �X�V�e�[�u����;

'*�y�߂�l�z
'*R    �Ȃ�;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/20; ����(���j�V�X�e���������); �V�K�쐬;
'*< Comment End >**************************************************'
Private Sub SubSqlInsTen(ByVal intN, ByVal strTenCd, ByRef sql, ByRef arrParamIns As Object, ByVal strTbl)
    'SQL���쐬
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
    
    If strTbl = TBL_BAI_HEN Then        '�����ύX���
        If paryCsvData(intN)(enmTenColNo.Baika) <> "" Then
            sql = sql & "'" & CLng(paryCsvData(intN)(enmTenColNo.Baika)) & "'," & vbCrLf
        Else
            sql = sql & "null," & vbCrLf
        End If

'********** 2005/05/06 �k�V **********��
'���s�\�[�X�R�[�h�R�����g��
'        sql = sql & "'" & paryCsvData(intN)(enmTenColNo.BaikaRiyuu) & "'," & vbCrLf
'EXCEL���ڂ��u�����N�̏ꍇ�̓e�[�u���ɂ�NULL���Z�b�g����
        If IsNull(paryCsvData(intN)(enmTenColNo.BaikaRiyuu)) _
        Or paryCsvData(intN)(enmTenColNo.BaikaRiyuu) = "" Then
            sql = sql & "Null," & vbCrLf
        Else
            sql = sql & "'" & paryCsvData(intN)(enmTenColNo.BaikaRiyuu) & "'," & vbCrLf
        End If
'********** 2005/05/06 �k�V **********��
    
    ElseIf strTbl = TBL_SII_HEN Then    '�d���ύX���

'********** 2005/05/06 �k�V **********��
'���s�\�[�X�R�[�h�R�����g��
'        sql = sql & "'" & paryCsvData(intN)(enmTenColNo.Torihikisaki) & "'," & vbCrLf
'EXCEL���ڂ��u�����N�̏ꍇ�̓e�[�u���ɂ�NULL���Z�b�g����
        If IsNull(paryCsvData(intN)(enmTenColNo.Torihikisaki)) _
        Or paryCsvData(intN)(enmTenColNo.Torihikisaki) = "" Then
            sql = sql & "Null," & vbCrLf
        Else
            sql = sql & "'" & paryCsvData(intN)(enmTenColNo.Torihikisaki) & "'," & vbCrLf
        End If
'********** 2005/05/06 �k�V **********��
        
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
    ElseIf strTbl = TBL_CEN_HEN Then    '�Z���^�[�R�[�h�ύX���

'********** 2005/05/06 �k�V **********��
'���s�\�[�X�R�[�h�R�����g��
'        sql = sql & "'" & paryCsvData(intN)(enmTenColNo.CenterKbn) & "'," & vbCrLf
'EXCEL���ڂ��u�����N�̏ꍇ�̓e�[�u���ɂ�NULL���Z�b�g����
        If IsNull(paryCsvData(intN)(enmTenColNo.CenterKbn)) _
        Or paryCsvData(intN)(enmTenColNo.CenterKbn) = "" Then
            sql = sql & "Null," & vbCrLf
        Else
            sql = sql & "'" & paryCsvData(intN)(enmTenColNo.CenterKbn) & "'," & vbCrLf
        End If
'********** 2005/05/06 �k�V **********��
        
        '�Z���^�[�敪��0�̏ꍇ��NULL�A1�̏ꍇ��'0151'�A2�̏ꍇ�͎��Ə��}�X�^����擾
'********** 2005/05/06 �k�V **********��
'���s�\�[�X�R�[�h�R�����g��
'        sql = sql & "'" & FncGetCenterCd(paryCsvData(intN)(enmTenColNo.CenterKbn), strTenCd) & "'," & vbCrLf  '�Z���^�[�R�[�h
'EXCEL���ڂ��u�����N�̏ꍇ�̓e�[�u���ɂ�NULL���Z�b�g����
        Dim strWorkString As String
        strWorkString = FncGetCenterCd(paryCsvData(intN)(enmTenColNo.CenterKbn), strTenCd)
        If strWorkString = "" Then
            sql = sql & "Null," & vbCrLf
        Else
            sql = sql & "'" & strWorkString & "'," & vbCrLf
        End If
'********** 2005/05/06 �k�V **********��
    
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

'*�y�v���V�[�W�����z
'*!    SubSqlUpdTen

'*�y�����T�v�z
'*?    UPDATE���쐬�i�X�ʏ��j

'*�y�p�����[�^�z
'*P    intN;            int;     ���Ԗڂ̃��R�[�h��;
'*P    strTenCd;        string;  �X�R�[�h;
'*P    sql;             string;  SQL���i�i�[�p�j;
'*P    arrParamUpd;     Object; �p�����[�^�i�i�[�p�j;
'*P    strTbl;          string;  �X�V�e�[�u����;

'*�y�߂�l�z
'*R    �Ȃ�;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/20; ����(���j�V�X�e���������); �V�K�쐬;
'*< Comment End >**************************************************'
Private Sub SubSqlUpdTen(ByVal intN, ByVal strTenCd, ByRef sql, ByRef arrParamUpd As Object, ByVal strTbl)
    'SQL���쐬
    sql = "UPDATE " & strTbl & " WITH(UPDLOCK) " & vbCrLf
    sql = sql & " SET " & vbCrLf
    
    If strTbl = TBL_BAI_HEN Then    '�����ύX���i�����A�����ύX���R�j
        If paryCsvData(intN)(enmTenColNo.Baika) <> "" Then
            sql = sql & "baika = '" & CLng(paryCsvData(intN)(enmTenColNo.Baika)) & "', " & vbCrLf
        End If
        If paryCsvData(intN)(enmTenColNo.BaikaRiyuu) <> "" Then
            sql = sql & "baika_henkou_riyuu_cd = '" & paryCsvData(intN)(enmTenColNo.BaikaRiyuu) & "', " & vbCrLf
        End If
    ElseIf strTbl = TBL_SII_HEN Then    '�d���ύX���i�����A�����P�ʁA�����j
        If paryCsvData(intN)(enmTenColNo.Torihikisaki) <> "" Then
            sql = sql & "torihikisaki_cd = '" & paryCsvData(intN)(enmTenColNo.Torihikisaki) & "', " & vbCrLf
        End If
        If paryCsvData(intN)(enmTenColNo.Hacchuutani) <> "" Then
            sql = sql & "hattyuu_tani = '" & CLng(paryCsvData(intN)(enmTenColNo.Hacchuutani)) & "', " & vbCrLf
        End If
        If paryCsvData(intN)(enmTenColNo.Genka) <> "" Then
            sql = sql & "genka = '" & CDbl(paryCsvData(intN)(enmTenColNo.Genka)) & "', " & vbCrLf
        End If
    ElseIf strTbl = TBL_CEN_HEN Then    '�Z���^�[�R�[�h�ύX���i�Z���^�[�敪�A�Z���^�[�R�[�h�j
        If paryCsvData(intN)(enmTenColNo.CenterKbn) <> "" Then
            sql = sql & "center_kbn = '" & paryCsvData(intN)(enmTenColNo.CenterKbn) & "', " & vbCrLf
            '�Z���^�[�敪��0�̏ꍇ��NULL�A1�̏ꍇ��'0151'�A2�̏ꍇ�͎��Ə��}�X�^����擾
'********** 2005/05/06 �k�V **********��
'���s�\�[�X�R�[�h�R�����g��
'            sql = sql & "center_cd = '" & FncGetCenterCd(paryCsvData(intN)(enmTenColNo.CenterKbn), strTenCd) & "', " & vbCrLf
'EXCEL���ڂ��u�����N�̏ꍇ�̓e�[�u���ɂ�NULL���Z�b�g����
            Dim strWorkString As String
            strWorkString = FncGetCenterCd(paryCsvData(intN)(enmTenColNo.CenterKbn), strTenCd)
            If strWorkString = "" Then
                sql = sql & "center_cd = dbNull.value," & vbCrLf
            Else
                sql = sql & "center_cd = '" & strWorkString & "'," & vbCrLf
            End If
'********** 2005/05/06 �k�V **********��
        End If
    End If
        
'    sql = sql & "touroku_kbn = '" & WF_KANRI_TOUROKU_KBN_SINKI & "', " & vbCrLf    '�o�^�敪�͍X�V���Ȃ�
    sql = sql & "kousin_user_id = '" & pstrTantousyaId & "', " & vbCrLf
    sql = sql & "kousin_ymd_hms = '" & Now & "' " & vbCrLf
    'WHERE��
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

'*�y�v���V�[�W�����z
'*!    SubSqlSelWFKanri

'*�y�����T�v�z
'*?    SELECT���쐬�iWF�Ǘ����ڃf�[�^�`�F�b�N�p�j

'*�y�p�����[�^�z
'*P    intN;            int;     ���Ԗڂ̃��R�[�h��;
'*P    sql;             string;  SQL���i�i�[�p�j;
'*P    arrPrm;          Object; �p�����[�^�i�i�[�p�j;
'*P    lngP_Cnt;        Long;    �p�����[�^���i�i�[�p�j;
'*P    intType;         int;     �������~or�X�ʏ��;

'*�y�߂�l�z
'*R    �Ȃ�;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/26; ����(���j�V�X�e���������); �V�K�쐬;
'*< Comment End >**************************************************'
Private Sub SubSqlSelWFKanri(ByVal intN, ByRef sql, ByRef arrPrm As Object, ByRef lngP_Cnt, ByVal intType)

    sql = "SELECT COUNT(*) FROM t_syousen_wf_kanri WITH(UPDLOCK) " & vbCrLf
    sql = sql & "WHERE" & vbCrLf
    sql = sql & " kikaku_syurui_cd = ? AND " & vbCrLf
    sql = sql & " nen_ren_no = ? AND " & vbCrLf
    sql = sql & " eda_no = ? AND " & vbCrLf
    sql = sql & " jan_cd = ? " & vbCrLf
    
    '�p�����[�^�Z�b�g
    ReDim Preserve arrPrm(lngP_Cnt + 3) As Object
    arrPrm(lngP_Cnt + 0) = MakeParam("@KIKAKU_SYURUI", DT_Char, 3, pstrKikakuSyurui)
    arrPrm(lngP_Cnt + 1) = MakeParam("@NEN_REN_NO", DT_Char, 10, pstrKikakuNenrenban)
    arrPrm(lngP_Cnt + 2) = MakeParam("@EDA_NO", DT_Char, 3, pstrKikakuEdaban)
    If intType = RET_FLD_HACH Then  '�������~
        arrPrm(lngP_Cnt + 3) = MakeParam("@JAN_CD", DT_VarChar, 20, paryCsvData(intN)(enmHacColNo.janCD))
    ElseIf intType = RET_FLD_TENB Then  '�X�ʏ��
        arrPrm(lngP_Cnt + 3) = MakeParam("@JAN_CD", DT_VarChar, 20, paryCsvData(intN)(enmTenColNo.janCD))
    ElseIf intType = RET_FLD_SYO Then
        arrPrm(lngP_Cnt + 3) = MakeParam("@JAN_CD", DT_VarChar, 20, paryCsvData(intN)(enmSyoColNo.janCD))
    ElseIf intType = RET_FLD_BUN Then
        arrPrm(lngP_Cnt + 3) = MakeParam("@JAN_CD", DT_VarChar, 20, paryCsvData(intN)(enmBunColNo.janCD))
    End If
    lngP_Cnt = lngP_Cnt + 4

End Sub

'*< UNISYSTEM >****************************************************'

'*�y�v���V�[�W�����z
'*!    SubSqlInsWFKanri

'*�y�����T�v�z
'*?    INSERT���쐬�iWF�Ǘ����ڒǉ��p�j

'*�y�p�����[�^�z
'*P    intN;            int;     ���Ԗڂ̃��R�[�h��;
'*P    sql;             string;  SQL���i�i�[�p�j;
'*P    arrPrm;          Object; �p�����[�^�i�i�[�p�j;
'*P    lngP_Cnt;        Long;    �p�����[�^���i�i�[�p�j;
'*P    intType;         int;     �������~or�X�ʏ��;

'*�y�߂�l�z
'*R    �Ȃ�;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/26; ����(���j�V�X�e���������); �V�K�쐬;
'*< Comment End >**************************************************'
Private Sub SubSqlInsWFKanri(ByVal intN, ByRef sql, ByRef arrPrm As Object, ByRef lngP_Cnt, ByVal intType)
    
    'SQL���쐬
'********** 2005/04/27 �k�V **********��
'���s�\�[�X�R�[�h�R�����g��
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
'    sql = sql & " '" & WF_KANRI_TOUROKU_KBN_SYUUSEI & "','0'," & vbCrLf '�o�^�敪=2�A���O�敪=0
'    sql = sql & " '" & pstrTantousyaId & "','0',?," & vbCrLf '�o�^�S���ҁA�v�搔=0�A���No.
'    sql = sql & " null,null,null,null,key_cd,null,null,null,null,null," & vbCrLf  '�����R�[�h
'    If pflgJigJyoToukbn = True Then '���ƕʏ��o�^�敪
'        sql = sql & "'1'," & vbCrLf
'    Else
'        sql = sql & "null," & vbCrLf
'    End If
'    sql = sql & " null,null,null,null," & vbCrLf
'    If pflgBaikaToukbn = True Then '�����ύX���o�^�敪
'        sql = sql & "'1'," & vbCrLf
'    Else
'        sql = sql & "null," & vbCrLf
'    End If
'    If pflgSiireToukbn = True Then '�d���ύX���o�^�敪
'        sql = sql & "'1'," & vbCrLf
'    Else
'        sql = sql & "null," & vbCrLf
'    End If
'    If pflgCenterToukbn = True Then '�������ޕύX���o�^�敪
'        sql = sql & "'1'," & vbCrLf
'    Else
'        sql = sql & "null," & vbCrLf
'    End If
'    sql = sql & " null,null,'" & pstrTantousyaId & "',null," & vbCrLf   '�o�^UserID
'    sql = sql & " '" & Now & "',null " & vbCrLf   '�o�^�N����
'    sql = sql & "FROM v_key_1 WITH(readcommitted) " & vbCrLf
'    sql = sql & "WHERE jan_cd = ? " & vbCrLf

'�t�B�[���h�ɃZ�b�g����l���ꕔ�C��
    sql = "INSERT INTO t_syousen_wf_kanri"
    sql = sql & "( kikaku_syurui_cd"             '����ރR�[�h
    sql = sql & ", nen_ren_no"                   '�N�A��
    sql = sql & ", eda_no"                       '�}��
    sql = sql & ", jan_cd"                       'JAN����
    sql = sql & ", jan_umu_kbn"                  'JAN�L���敪(0:�Ȃ� 1:����)
    sql = sql & ", touroku_kbn"                  '�o�^�敪(1:�V�K 2:�C�� 9:����)
    sql = sql & ", jyogai_kbn"                   '���O�敪(0:�R�t���� 1:���O)
    sql = sql & ", touroku_tantousya"            '�o�^�S����
    sql = sql & ", keikakusuu"                   '�v�搔
    sql = sql & ", kikaku_no"                    '���No.(�����+�N�A��+�}��)
    sql = sql & ", kikaku_kahi_kbn"              '���ۋ敪(0:�����F 1:���F 9:�۔F)
    sql = sql & ", kikaku_syouninsya_comment"    '��揳�F�Һ���
    sql = sql & ", hinsyu_kahi_kbn"              '�i��ۋ敪(0:�����F 1:���F 9:�۔F)
    sql = sql & ", hinsyu_syouninsya_comment"    '�i�폳�F�Һ���
    sql = sql & ", naibu_cd"                     '��������(�����}�X�^�̎�L�[)
    sql = sql & ", hyouji_siyou_touroku_kbn"     '�\��/�d�l���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", sekisai_size_touroku_kbn"     '�ύ�/���ޏ��o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", pop_touroku_kbn"              'POP���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", syanai_kihon_touroku_kbn"     '�Г���{���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", kowake_touroku_kbn"           '�������o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", jigyoubetu_touroku_kbn"       '���ƕʏ��o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", hanbai_touroku_kbn"           '�̔����o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", siire_touroku_kbn"            '�d�����o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", hinsyu_taikei_touroku_kbn"    '�i��̌n���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", uriba_touroku_kbn"            '����̌n���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", baika_henkou_touroku_kbn"     '�����ύX���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", siire_henkou_touroku_kbn"     '�d���ύX���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", center_cd_henkou_touroku_kbn" '�������ޕύX���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", kowake_henkou_touroku_kbn"    '�����ύX���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", hanki_henkou_touroku_kbn"     '�����ύX���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", touroku_user_id"              '�o�^UserID
    sql = sql & ", kousin_user_id"               '�X�VUserID
    sql = sql & ", touroku_ymd_hms"              '�o�^�N���������b
    sql = sql & ", kousin_ymd_hms"               '�X�V�N���������b
    
    sql = sql & ") SELECT"
    sql = sql & "  ?"                                                '����ރR�[�h
    sql = sql & ", ?"                                                '�N�A��
    sql = sql & ", ?"                                                '�}��
    sql = sql & ", jan_cd"                                           'JAN����
    sql = sql & ", '1'"                                              'JAN�L���敪(0:�Ȃ� 1:����)
    sql = sql & ", '" & WF_KANRI_TOUROKU_KBN_SYUUSEI & "'"           '�o�^�敪(1:�V�K 2:�C�� 9:����)
    sql = sql & ", '0'"                                              '���O�敪(0:�R�t���� 1:���O)
    sql = sql & ", '" & pstrTantousyaId & "'"                        '�o�^�S����
    sql = sql & ", '0'"                                              '�v�搔
    sql = sql & ", ?"                                                '���No.(�����+�N�A��+�}��)
    sql = sql & ", '0'"                                              '���ۋ敪(0:�����F 1:���F 9:�۔F)
    sql = sql & ", null"                                             '��揳�F�Һ���
    sql = sql & ", '0'"                                              '�i��ۋ敪(0:�����F 1:���F 9:�۔F)
    sql = sql & ", null"                                             '�i�폳�F�Һ���
    sql = sql & ", key_cd"                                           '��������(�����}�X�^�̎�L�[)
    sql = sql & ", '1'"                                              '�\��/�d�l���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", '0'"                                              '�ύ�/���ޏ��o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", '0'"                                              'POP���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", '0'"                                              '�Г���{���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", '0'"                                              '�������o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", '" & IIf(pflgJigJyoToukbn = True, "1", "0") & "'" '���ƕʏ��o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", '0'"                                              '�̔����o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", '0'"                                              '�d�����o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", '0'"                                              '�i��̌n���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", '0'"                                              '����̌n���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", '" & IIf(pflgBaikaToukbn = True, "1", "0") & "'"  '�����ύX���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", '" & IIf(pflgSiireToukbn = True, "1", "0") & "'"  '�d���ύX���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", '" & IIf(pflgCenterToukbn = True, "1", "0") & "'" '�������ޕύX���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", '0'"                                              '�����ύX���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", '0'"                                              '�����ύX���o�^�敪(0:���o�^ 1:�o�^��)
    sql = sql & ", '" & pstrTantousyaId & "'"                        '�o�^UserID
    sql = sql & ", null"                                             '�X�VUserID
    sql = sql & ", '" & Format(Now, "yyyy/mm/dd hh:nn:ss") & "'"     '�o�^�N���������b
    sql = sql & ", null"                                             '�X�V�N���������b
    
    sql = sql & "  FROM v_key_1 WITH(readcommitted)"
    sql = sql & " WHERE jan_cd = ?"
'********** 2005/04/27 �k�V **********��
    
    '�p�����[�^�Z�b�g
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

'*�y�v���V�[�W�����z
'*!    SubSqlUpdWFKanri

'*�y�����T�v�z
'*?    UPDATE���쐬�iWF�Ǘ����ڃe�[�u���X�V�p�j

'*�y�p�����[�^�z
'*P    intN;            int;     ���Ԗڂ̃��R�[�h��;
'*P    sql;             string;  SQL���i�i�[�p�j;
'*P    arrPrm;          Object; �p�����[�^�i�i�[�p�j;
'*P    lngP_Cnt;        Long;    �p�����[�^���i�i�[�p�j;
'*P    intType;         int;     �������~or�X�ʏ��;

'*�y�߂�l�z
'*R    �Ȃ�;

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/20; ����(���j�V�X�e���������); �V�K�쐬;
'*< Comment End >**************************************************'
'********** 2005/04/26 ���� �ύX **********��
Private Sub SubSqlUpdWFKanri(ByVal intN, ByRef sql, ByRef arrPrm As Object, ByRef lngP_Cnt, ByVal intType)
'Private Sub SubSqlUpdWFKanri(ByVal intN, ByRef sql, ByRef arrParamUpd As Object)
'********** 2005/04/26 ���� �ύX **********��
    'SQL���쐬
    sql = "UPDATE t_syousen_wf_kanri WITH(UPDLOCK) " & vbCrLf
    sql = sql & "SET " & vbCrLf
    sql = sql & "touroku_kbn = '" & WF_KANRI_TOUROKU_KBN_SYUUSEI & "', " & vbCrLf  '�o�^�敪�́A2:�C��
    sql = sql & "touroku_tantousya = '" & pstrTantousyaId & "', " & vbCrLf     '�o�^�S���҂́AUserID
    sql = sql & "kikaku_no = '" & pstrKikakuSyurui & pstrKikakuNenrenban & "-" & pstrKikakuEdaban & "', " & vbCrLf    '���No.
    sql = sql & "naibu_cd = A.key_cd, " & vbCrLf    '��L�[
    
'********** 2005/04/25 ���� �ǉ� **********��
    If pflgJigJyoToukbn = True Then     '���ƕʏ��o�^�敪
        sql = sql & "jigyoubetu_touroku_kbn = '1', " & vbCrLf
    End If
'********** 2005/04/25 ���� �ǉ� **********��
    If pflgBaikaToukbn = True Then      '�����ύX���o�^�敪
        sql = sql & "baika_henkou_touroku_kbn = '1', " & vbCrLf
    End If
    If pflgSiireToukbn = True Then      '�d���ύX���o�^�敪
        sql = sql & "siire_henkou_touroku_kbn = '1', " & vbCrLf
    End If
    If pflgCenterToukbn = True Then     '�Z���^�[�R�[�h�ύX���o�^�敪
        sql = sql & "center_cd_henkou_touroku_kbn = '1', " & vbCrLf
    End If
    
'********** 2005/04/27 �k�V **********��
'�\���d�l�o�^�敪�𗧂Ă�
    sql = sql & "hyouji_siyou_touroku_kbn = '1', " & vbCrLf
'********** 2005/04/27 �k�V **********��
    
    sql = sql & "kousin_user_id = '" & pstrTantousyaId & "', " & vbCrLf
    sql = sql & "kousin_ymd_hms = '" & Now & "' " & vbCrLf
    
    '��L�[�p�̃e�[�u���̐錾
    sql = sql & "FROM (SELECT key_cd FROM v_key_1 WITH(readcommitted) WHERE jan_cd = ?) AS A " & vbCrLf    '��L�[
    'WHERE��
    sql = sql & "WHERE " & vbCrLf
    sql = sql & " kikaku_syurui_cd = ? AND " & vbCrLf
    sql = sql & " nen_ren_no = ? AND " & vbCrLf
    sql = sql & " eda_no = ? AND " & vbCrLf
    sql = sql & " jan_cd = ? " & vbCrLf
    
'********** 2005/04/26 ���� �ύX **********��
    '�p�����[�^�Z�b�g
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
'********** 2005/04/26 ���� �ύX **********��

End Sub

'*< UNISYSTEM >****************************************************'

'*�y�v���V�[�W�����z
'*!    FncGetCenterCd

'*�y�����T�v�z
'*?    �Z���^�[�敪�����ɂ��āA�Z���^�[�R�[�h���擾����B

'*�y�p�����[�^�z
'*P    strCenterKbn; string; �Z���^�[�敪;
'*P    strTenCd;     string; �X�R�[�h�i�Z���^�[�敪��2�̏ꍇ�Ɏg�p�j;

'*�y�߂�l�z
'*R    String;�@�Z���^�[�R�[�h

'******************************************************************'
'*�y�ύX�����z
'*M    2005/04/21; ����(���j�V�X�e���������); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncGetCenterCd(ByVal strCenterKbn, ByVal strTenCd) As String
    Dim sql, objCDB
    Dim rsSel As ADODB.Recordset

    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncGetCenterCd(�Z���^�[�敪�F" & strCenterKbn & ",�X�R�[�h�F" & strTenCd & ")"
    
    Select Case strCenterKbn
    
        Case "0":       '�Z���^�[�敪��0�̏ꍇ�ANULL���Z�b�g
            FncGetCenterCd = ""
        Case "1":       '�Z���^�[�敪��1�̏ꍇ�A0151(���ʃZ���^�[)���Z�b�g
            FncGetCenterCd = CENTERCD_KODAMA
        Case "2":       '�Z���^�[�敪��2�̏ꍇ�A�X�R�[�h�����Ɏ��Ə��}�X�^����擾�����Z���^�[�R�[�h���Z�b�g
                objCDB = CreateObject("COMDBUtil.CDBAccess")    'CDBAccess���C���X�^���X��
            sql = "SELECT center_cd FROM v_jigyousyo WITH(readcommitted) " & vbCrLf
            sql = sql & "WHERE jigyousyo_cd = '" & strTenCd & "' " & vbCrLf
                rsSel = objCDB.GetDCRecordset("VSWF", sql)

            '�f�[�^������������NULL���Z�b�g����B
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
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description
End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncSetError
'*�y�����T�v�z
'*?    ���������G���[����ݒ肵�܂��B�G���[���\���̔z��̍Ō㕔��
'*?    �G���[����}�����܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intErrorLine ; �G���[�����������s
'*P    intErrorField ; �G���[�����������t�B�[���h
'*P    intErrorType ; ���������G���[���
'*P    strErrorParameter ; �G���[���b�Z�[�W�⏕�p������
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
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
'*�y�v���V�[�W�����z
'*!    FncSetErrorPDC
'*�y�����T�v�z
'*?    ���������G���[����ݒ肵�܂��B�G���[���\���̔z��̍Ō㕔��
'*?    �G���[����}�����܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intErrorLine ; �G���[�����������s
'*P    intErrorField ; �G���[�����������t�B�[���h
'*P    intErrorType ; ���������G���[���
'*P    strErrorParameter ; �G���[���b�Z�[�W�⏕�p������
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2012/08/10; �k�o(��A); �V�K�쐬;
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
'*�y�v���V�[�W�����z
'*!    FncSetErrorMsg_Hac
'*�y�����T�v�z
'*?    ���������G���[���ׂẴG���[���b�Z�[�W���擾���܂��B
'*?    (�������~���Łj
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncGetErrorMsg; String;���������G���[���ׂẴG���[���b�Z�[�W
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/25; ���@����(USC); �V�K�쐬;
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
                        "�y" & FncGetErrorMsgField_Hac(intErrorField) & "�z " & _
                        FncGetErrorMsgType_Hac(intErrorField, intErrorType, strErrorParameter) & _
                        "</div>" & vbCrLf & "<div align='center' class='error'>"
    Next

End Function
'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncGetErrorMsg_Bun
'*�y�����T�v�z
'*?    ���������G���[���ׂẴG���[���b�Z�[�W���擾���܂��B
'*?    (���ޏ��Łj
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncSetErrorMsg_Bun; String;���������G���[���ׂẴG���[���b�Z�[�W
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/09/01; �����@����(IC); �V�K�쐬;
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
                        "�y" & FncGetErrorMsgField_Bun(intErrorField) & "�z " & _
                        FncGetErrorMsgType_Bun(intErrorField, intErrorType, strErrorParameter) & _
                        "</div>" & vbCrLf & "<div align='center' class='error'>"
    Next

End Function
'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncGetErrorMsg_Syo
'*�y�����T�v�z
'*?    ���������G���[���ׂẴG���[���b�Z�[�W���擾���܂��B
'*?    (���ޏ��Łj
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncGetErrorMsg_Syo; String;���������G���[���ׂẴG���[���b�Z�[�W
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/09/01; �����@����(IC); �V�K�쐬;
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
                        "�y" & FncGetErrorMsgField_Syo(intErrorField) & "�z " & _
                        FncGetErrorMsgType_Syo(intErrorField, intErrorType, strErrorParameter) & _
                        "</div>" & vbCrLf & "<div align='center' class='error'>"
    Next

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncSetError
'*�y�����T�v�z
'*?    ���������G���[���ׂẴG���[���b�Z�[�W���擾���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncGetErrorMsg; String;���������G���[���ׂẴG���[���b�Z�[�W
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
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
                        "�y" & FncGetErrorMsgField(intErrorField) & "�z " & _
                        FncGetErrorMsgType(intErrorField, intErrorType, strErrorParameter) & _
                        "</div>" & vbCrLf & "<div align='center' class='error'>"
    Next

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncGetErrorMsg_Modosi
'*�y�����T�v�z
'*?    ���������G���[���ׂẴG���[���b�Z�[�W���擾���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncGetErrorMsg_Modosi; String;���������G���[���ׂẴG���[���b�Z�[�W
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2009/06/06; Supattra T.(Tostem Thai); �V�K�쐬;
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
                        "�y" & FncGetErrorMsgField_Modosi(intErrorField) & "�z " & _
                        FncGetErrorMsgType(intErrorField, intErrorType, strErrorParameter) & _
                        "</div>" & vbCrLf & "<div align='center' class='error'>"
    Next

End Function

'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncGetErrorMsg_PDC
'*�y�����T�v�z
'*?    ���������G���[���ׂẴG���[���b�Z�[�W���擾���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncGetErrorMsg_PDC; String;���������G���[���ׂẴG���[���b�Z�[�W
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2012/08/17; �k�o(��A); �V�K�쐬;
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
'*�y�v���V�[�W�����z
'*!    FncGetErrorMsgLine
'*�y�����T�v�z
'*?    �G���[�����������s�ɑΉ�����������𐶐����܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intErrorLine ; �G���[�����������s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncGetErrorMsgLine; String;�G���[�����������s�̕�����\��
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgLine(intErrorLine As Integer) As String

    FncGetErrorMsgLine = intErrorLine & "�s�ځF"

End Function

' '*< UNISYSTEM >**************************************************'
'*�y�v���V�[�W�����z
'*!    FncGetErrorMsgField_Hac
'*�y�����T�v�z
'*?    �G���[�����������t�B�[���h�����擾���܂��B
'*?    (�������~���Łj
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intErrorField ; �G���[�����������t�B�[���h
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncGetErrorMsgField; String;�G���[�����������t�B�[���h��
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/25; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgField_Hac(intErrorField As Integer) As String
    Select Case intErrorField
        
        Case enmHacColNo.jigCD:         FncGetErrorMsgField_Hac = "�Y�����ƃR�[�h"
        Case enmHacColNo.SyuryouDate:   FncGetErrorMsgField_Hac = "�����I����"
        Case enmHacColNo.janCD:         FncGetErrorMsgField_Hac = "JAN�R�[�h"
        Case enmHacColNo.Syobun:        FncGetErrorMsgField_Hac = "�������@"
        Case enmHacColNo.ChushiKbn:     FncGetErrorMsgField_Hac = "�������~���R�敪"
        Case enmHacColNo.DaiJanCd:      FncGetErrorMsgField_Hac = "��֏��i�R�[�h"
               
    End Select
End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncGetErrorMsgField_Bun
'*�y�����T�v�z
'*?    �G���[�����������t�B�[���h�����擾���܂��B
'*?    (���ޏ��Łj
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intErrorField ; �G���[�����������t�B�[���h
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncGetErrorMsgField_Bun; String;�G���[�����������t�B�[���h��
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/09/01; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgField_Bun(intErrorField As Integer) As String
    Select Case intErrorField
        
        Case enmBunColNo.HenkouDate:         FncGetErrorMsgField_Bun = "�ύX��"
        Case enmBunColNo.janCD:              FncGetErrorMsgField_Bun = "JAN�R�[�h"
        Case enmBunColNo.Bunrui1:            FncGetErrorMsgField_Bun = "��1���ރR�[�h"
        Case enmBunColNo.Bunrui2:            FncGetErrorMsgField_Bun = "��2���ރR�[�h"
        Case enmBunColNo.Bunrui3:            FncGetErrorMsgField_Bun = "��3���ރR�[�h"
    
    End Select
End Function
'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncGetErrorMsgField_Syo
'*�y�����T�v�z
'*?    �G���[�����������t�B�[���h�����擾���܂��B
'*?    (��{���i�Łj
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intErrorField ; �G���[�����������t�B�[���h
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncGetErrorMsgField_Syo; String;�G���[�����������t�B�[���h��
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/09/01; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgField_Syo(intErrorField As Integer) As String
    Select Case intErrorField
        
        Case enmSyoColNo.janCD:                         FncGetErrorMsgField_Syo = "JAN�R�[�h"
        Case enmSyoColNo.SyouhinmeiKana:                FncGetErrorMsgField_Syo = "�J�i���i��(TA�`�[�p)"
        Case enmSyoColNo.SyouhinmeiKanji:               FncGetErrorMsgField_Syo = "�������i��(SA/��ײ����ޗp)"
        Case enmSyoColNo.ReceiptmeiKana:                FncGetErrorMsgField_Syo = "���V�[�g��(��)"
        Case enmSyoColNo.ReceiptmeiKanji:               FncGetErrorMsgField_Syo = "���V�[�g��(����)"
        Case enmSyoColNo.SyouhihinmeiHansoku:           FncGetErrorMsgField_Syo = "���i��(�̑��p)"
        
        Case enmSyoColNo.MakerMei:                      FncGetErrorMsgField_Syo = "���[�J�[��"
        Case enmSyoColNo.KikakuSize:                    FncGetErrorMsgField_Syo = "�K�i/�T�C�Y"
        Case enmSyoColNo.HinbanKataban:                 FncGetErrorMsgField_Syo = "�i��/�^��"
        Case enmSyoColNo.IroGara:                       FncGetErrorMsgField_Syo = "�F/��"
        Case enmSyoColNo.Taste:                         FncGetErrorMsgField_Syo = "�e�[�X�g"
        Case enmSyoColNo.Seisantimei:                   FncGetErrorMsgField_Syo = "���Y�n��"
        Case enmSyoColNo.Tokutyou:                      FncGetErrorMsgField_Syo = "����"
        Case enmSyoColNo.SyouhinCopy:                   FncGetErrorMsgField_Syo = "���i�R�s�["
        Case enmSyoColNo.TuujyouKakaku:                 FncGetErrorMsgField_Syo = "�ʏ퉿�i"
        Case enmSyoColNo.MakerKakaku:                   FncGetErrorMsgField_Syo = "���[�J�[��]�������i"
        Case enmSyoColNo.POPLayout1:                    FncGetErrorMsgField_Syo = "POP���C�A�E�g�P"
        Case enmSyoColNo.POPLayout2:                    FncGetErrorMsgField_Syo = "POP���C�A�E�g�Q"
        Case enmSyoColNo.PopSize:                       FncGetErrorMsgField_Syo = "POP�T�C�Y"
        
        Case enmSyoColNo.AtukaiPtn:                     FncGetErrorMsgField_Syo = "�����p�^�[��"
        Case enmSyoColNo.HyoujyunBaika:                 FncGetErrorMsgField_Syo = "�W������"
        Case enmSyoColNo.SikiriGenka:                   FncGetErrorMsgField_Syo = "�d�،���"
        Case enmSyoColNo.VhJyuyouCd:                    FncGetErrorMsgField_Syo = "VH���v�R�[�h"
        Case enmSyoColNo.RdJyuyouCd:                    FncGetErrorMsgField_Syo = "RD���v�R�[�h"
        Case enmSyoColNo.VcsJyuyouCd:                   FncGetErrorMsgField_Syo = "VCS���v�R�[�h"
        Case enmSyoColNo.SvhJyuyouCd:                   FncGetErrorMsgField_Syo = "SVH���v�R�[�h"
    
    End Select
End Function

' '*< UNISYSTEM >**************************************************'
'*�y�v���V�[�W�����z
'*!    FncGetErrorMsgField
'*�y�����T�v�z
'*?    �G���[�����������t�B�[���h�����擾���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intErrorField ; �G���[�����������t�B�[���h
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncGetErrorMsgField; String;�G���[�����������t�B�[���h��
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgField(intErrorField As Integer) As String
    Select Case intErrorField
        Case enmTenColNo.TenCd:         FncGetErrorMsgField = "�Y���X�R�[�h"
        Case enmTenColNo.HenkouDate:    FncGetErrorMsgField = "�ύX��"
        Case enmTenColNo.janCD:         FncGetErrorMsgField = "JAN�R�[�h"
        Case enmTenColNo.Genka:         FncGetErrorMsgField = "����"
        Case enmTenColNo.Baika:         FncGetErrorMsgField = "����"
        Case enmTenColNo.BaikaRiyuu:    FncGetErrorMsgField = "�����ύX���R"
        Case enmTenColNo.Torihikisaki:  FncGetErrorMsgField = "�����"
        Case enmTenColNo.Hacchuutani:   FncGetErrorMsgField = "�����P��"
        Case enmTenColNo.CenterKbn:     FncGetErrorMsgField = "�Z���^�[�敪"
    End Select
End Function

' '*< UNISYSTEM >**************************************************'
'*�y�v���V�[�W�����z
'*!    FncGetErrorMsgField
'*�y�����T�v�z
'*?    �G���[�����������t�B�[���h�����擾���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intErrorField ; �G���[�����������t�B�[���h
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncGetErrorMsgField; String;�G���[�����������t�B�[���h��
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2009/06/06; Supattra T.(Tostem Thai); �V�K�쐬;
'*M    2011/02/10; �c�ƃV�X�e��/�a�c;         �Z���^�[�R�[�h�ǉ�;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgField_Modosi(intErrorField As Integer) As String
    Select Case intErrorField
        Case enmModColNo.TenCd:         FncGetErrorMsgField_Modosi = "�Y���X�R�[�h"
        Case enmModColNo.HenkouDate:    FncGetErrorMsgField_Modosi = "�ύX��"
        Case enmModColNo.janCD:         FncGetErrorMsgField_Modosi = "JAN�R�[�h"
        Case enmModColNo.Genka:         FncGetErrorMsgField_Modosi = "����"
        Case enmModColNo.Baika:         FncGetErrorMsgField_Modosi = "����"
        Case enmModColNo.BaikaRiyuu:    FncGetErrorMsgField_Modosi = "�����ύX���R"
        Case enmModColNo.Torihikisaki:  FncGetErrorMsgField_Modosi = "�����"
        Case enmModColNo.Hacchuutani:   FncGetErrorMsgField_Modosi = "�����P��"
        Case enmModColNo.CenterKbn:     FncGetErrorMsgField_Modosi = "�Z���^�[�敪"
        Case enmModColNo.centerCD:      FncGetErrorMsgField_Modosi = "�Z���^�[�R�[�h"
        Case enmModColNo.MoGenka:       FncGetErrorMsgField_Modosi = "�߂�����"
        Case enmModColNo.MoBaika:       FncGetErrorMsgField_Modosi = "�߂�����"
        Case enmModColNo.HattyuuTani:   FncGetErrorMsgField_Modosi = "�߂������P��"
        Case enmModColNo.PopSize:       FncGetErrorMsgField_Modosi = "POP�T�C�Y"
    End Select
End Function


' '*< UNISYSTEM >**************************************************'
'*�y�v���V�[�W�����z
'*!    FncGetErrorMsgType_Hac
'*�y�����T�v�z
'*?    ���������G���[�̓��e�𕶎���Ŏ擾���܂��B
'*?    (�������~���Łj
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intErrorField ; �G���[�����������t�B�[���h
'*P    intErrorType ; ���������G���[�̎��
'*P    strErrorParameter ; �G���[���b�Z�[�W�⏕�p������
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncGetErrorMsgField; String;���������G���[�̓��e(�t�B�[���h���ƂɈقȂ邱�Ƃ�����)
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/25; ���@����(USC); �V�K�쐬;
'*M    2006/09/13; �a�c(IC);        ���i�I��v�e�⍇���o�ŉ��P;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgType_Hac(intErrorField As Integer, _
                              intErrorType As Integer, _
                              Optional strErrorParameter As String = "")
     
    Select Case intErrorType
        Case PARAM_EMPTY:           FncGetErrorMsgType_Hac = "�K�{���ڂł��B"
        Case PARAM_NOTDISIT:        FncGetErrorMsgType_Hac = "���p�����œ��͂��ĉ������B"
        Case PARAM_NONARROW:        FncGetErrorMsgType_Hac = "���p�����œ��͂��ĉ������B"
        Case PARAM_NOTFINDMASTER:
            Select Case intErrorField
                Case enmHacColNo.jigCD
                    FncGetErrorMsgType_Hac = "�}�X�^�ɓo�^����Ă��Ȃ����ƃR�[�h���ݒ肳��Ă��܂��B"
                Case enmHacColNo.ChushiKbn
                    FncGetErrorMsgType_Hac = "�}�X�^�ɓo�^����Ă��Ȃ��������~���R�敪���ݒ肳��Ă��܂��B"
            End Select
        Case PARAM_WRONGCHECKDIGIT: FncGetErrorMsgType_Hac = "�`�F�b�N�f�W�b�g���m�F���ĉ������B"
        Case PARAM_WRONGDATE:       FncGetErrorMsgType_Hac = "���t�𐳂������͂��ĉ������BYYYYMMDD"
        Case PARAM_SHORTDISIT:
            Select Case intErrorField
                Case enmHacColNo.janCD:
                    FncGetErrorMsgType_Hac = "�����ɉߕs��������܂��B8����������13���̐�������͂��ĉ������B"
                Case enmHacColNo.DaiJanCd:
                    FncGetErrorMsgType_Hac = "�����ɉߕs��������܂��B8����������13���̐�������͂��ĉ������B"
            End Select
        Case PARAM_NOTEXISTDB:
            Select Case intErrorField
                Case enmHacColNo.jigCD
                    FncGetErrorMsgType_Hac = "�}�X�^�ɓo�^����Ă��Ȃ����ƃR�[�h���ݒ肳��Ă��܂��B"
                Case enmHacColNo.janCD
                    FncGetErrorMsgType_Hac = "JAN�R�[�h�͓���DB�ɓo�^����Ă��܂���B"
                Case enmHacColNo.ChushiKbn
                    FncGetErrorMsgType_Hac = "�}�X�^�ɓo�^����Ă��Ȃ��������~���R�敪���ݒ肳��Ă��܂��B"
                Case enmHacColNo.DaiJanCd
                    FncGetErrorMsgType_Hac = "��֏��i�R�[�h�͓���DB�ɓo�^����Ă��܂���B"
                Case enmHacColNo.SyuryouDate
                    FncGetErrorMsgType_Hac = "�����J�n��������DB�ɓo�^����Ă��܂���B"
            End Select
        Case PARAM_SAMECODEDB:      FncGetErrorMsgType_Hac = "JAN�R�[�h�͊�懂" & strErrorParameter & "�ɓo�^����Ă��܂��B"
        Case PARAM_SAMEKEYDB:       FncGetErrorMsgType_Hac = "�����R�[�h�͊�懂" & strErrorParameter & "�ɓo�^����Ă��܂��B"
        Case PARAM_NOTDEFINED:
            Select Case intErrorField
                Case enmHacColNo.Syobun
                    FncGetErrorMsgType_Hac = "0(���ݒ�)�A1(�ԕi)�A2(�ꎞ��~)�A3(�戵�I��(�X�U��Ώ�))�ȊO���ݒ肳��Ă��܂��B"
            End Select
        Case PARAM_EARLYCANCELDAY:   FncGetErrorMsgType_Hac = "�����J�n���������I�����ƂȂ�悤�ɓ��͂��ĉ������B"
'********** 2005/05/07 �k�V **********��
        Case PARAM_MISENTAKU: FncGetErrorMsgType_Hac = "�u00�v(���I��)�͎w��o���܂���B"
'********** 2005/05/07 �k�V **********��
    End Select

End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncGetErrorMsgType_Bun
'*�y�����T�v�z
'*?    ���������G���[�̓��e�𕶎���Ŏ擾���܂��B
'*?    (���ޏ��Łj
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intErrorField ; �G���[�����������t�B�[���h
'*P    intErrorType ; ���������G���[�̎��
'*P    strErrorParameter ; �G���[���b�Z�[�W�⏕�p������
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncGetErrorMsgType_Bun; String;���������G���[�̓��e(�t�B�[���h���ƂɈقȂ邱�Ƃ�����)
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/09/01; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgType_Bun(intErrorField As Integer, _
                              intErrorType As Integer, _
                              Optional strErrorParameter As String = "")
     
    Select Case intErrorType
        Case PARAM_EMPTY:           FncGetErrorMsgType_Bun = "�K�{���ڂł��B"
        Case PARAM_NOTDISIT:        FncGetErrorMsgType_Bun = "���p�����œ��͂��ĉ������B"
        Case PARAM_NONARROW:        FncGetErrorMsgType_Bun = "���p�����œ��͂��ĉ������B"
        Case PARAM_HANEISUU:        FncGetErrorMsgType_Bun = "���p�p�����œ��͂��ĉ������B"
        Case PARAM_WRONGCHECKDIGIT: FncGetErrorMsgType_Bun = "�`�F�b�N�f�W�b�g���m�F���ĉ������B"
        Case PARAM_WRONGDATE:       FncGetErrorMsgType_Bun = "���t�𐳂������͂��ĉ������BYYYYMMDD"
        Case PARAM_SHORTDISIT:
            Select Case intErrorField
                Case enmBunColNo.janCD
                    FncGetErrorMsgType_Bun = "�����ɉߕs��������܂��B8����������13���̐�������͂��ĉ������B"
                Case enmBunColNo.Bunrui1
                    FncGetErrorMsgType_Bun = "�����ɉߕs��������܂��B6���œ��͂��ĉ������B"
                Case enmBunColNo.Bunrui2
                    FncGetErrorMsgType_Bun = "�����ɉߕs��������܂��B6���œ��͂��ĉ������B"
                Case enmBunColNo.Bunrui3
                    FncGetErrorMsgType_Bun = "�����ɉߕs��������܂��B6���œ��͂��ĉ������B"
            End Select
        
        Case PARAM_NOTEXISTDB:      FncGetErrorMsgType_Bun = "JAN�R�[�h�͓���DB�ɓo�^����Ă��܂���B"
        Case PARAM_SAMECODEDB:      FncGetErrorMsgType_Bun = "JAN�R�[�h�͊�懂" & strErrorParameter & "�ɓo�^����Ă��܂��B"
        Case PARAM_SAMEKEYDB:       FncGetErrorMsgType_Bun = "�����R�[�h�͊�懂" & strErrorParameter & "�ɓo�^����Ă��܂��B"
        Case PARAM_BUNRUIDB:        FncGetErrorMsgType_Bun = "���ޏ�񂪑��݂��܂���B"
    End Select

End Function

' '*< UNISYSTEM >**************************************************'
'*�y�v���V�[�W�����z
'*!    FncGetErrorMsgType_Syo
'*�y�����T�v�z
'*?    ���������G���[�̓��e�𕶎���Ŏ擾���܂��B
'*?    (���i��{���Łj
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intErrorField ; �G���[�����������t�B�[���h
'*P    intErrorType ; ���������G���[�̎��
'*P    strErrorParameter ; �G���[���b�Z�[�W�⏕�p������
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncGetErrorMsgType_Syo; String;���������G���[�̓��e(�t�B�[���h���ƂɈقȂ邱�Ƃ�����)
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/25; ���@����(USC); �V�K�쐬;
'*M    2006/09/12; �a�c(IC);        ���i�I��v�e�⍇���o�ŉ��P;
'*M    2014/05/02; �a�c(LIS);       �֎~�������b�Z�[�W�ǉ�;
'*M    2015/09/15; ���ʃV�X�e��G/�a�c;        TC�V�X�e���\�z;
'*< Comment End >**************************************************'
'
Private Function FncGetErrorMsgType_Syo(intErrorField As Integer, _
                              intErrorType As Integer, _
                              Optional strErrorParameter As String = "")
     
    Select Case intErrorType
        Case PARAM_EMPTY:           FncGetErrorMsgType_Syo = "�K�{���ڂł��B"
        Case PARAM_NOTDISIT:        FncGetErrorMsgType_Syo = "���p�����œ��͂��ĉ������B"
        Case PARAM_NONARROW:        FncGetErrorMsgType_Syo = "���p�����œ��͂��ĉ������B"
        Case PARAM_HANEISUU:        FncGetErrorMsgType_Syo = "���p�p�����œ��͂��ĉ������B"
        Case PARAM_ZENKAKU:         FncGetErrorMsgType_Syo = "�S�p�����œ��͂��ĉ������B"
        Case PARAM_KINSIMOJI:       FncGetErrorMsgType_Syo = "�֎~�����w,'""\<>�x�����͂���Ă��܂��B"
        '2014/05/02 add wadak7
        Case PARAM_KINSIMOJI_RECEIPT:       FncGetErrorMsgType_Syo = "�֎~�����w�@�A�B�c�T�U�V�c�o�p�q�c�������c�x�������͂���Ă��܂��B"
        Case PARAM_NOTFINDMASTER:
            Select Case intErrorField
                Case enmSyoColNo.VhJyuyouCd
                    FncGetErrorMsgType_Syo = "VH���v�R�[�h���}�X�^�ɑ��݂��܂���B"
                Case enmSyoColNo.RdJyuyouCd
                    FncGetErrorMsgType_Syo = "R&D���v�R�[�h���}�X�^�ɑ��݂��܂���B"
                Case enmSyoColNo.VcsJyuyouCd
                    FncGetErrorMsgType_Syo = "VCS���v�R�[�h���}�X�^�ɑ��݂��܂���B"
                Case enmSyoColNo.SvhJyuyouCd
                    FncGetErrorMsgType_Syo = "SVH���v�R�[�h���}�X�^�ɑ��݂��܂���B"
                Case enmSyoColNo.AtukaiPtn
                    FncGetErrorMsgType_Syo = "�����p�^�[���R�[�h���}�X�^�ɑ��݂��܂���B"
            End Select
        Case PARAM_WRONGCHECKDIGIT: FncGetErrorMsgType_Syo = "�`�F�b�N�f�W�b�g���m�F���ĉ������B"
        Case PARAM_WRONGDATE:       FncGetErrorMsgType_Syo = "���t�𐳂������͂��ĉ������BYYYYMMDD"
        Case PARAM_SHORTDISIT:
            Select Case intErrorField
                Case enmSyoColNo.janCD:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B8����������13���̐�������͂��ĉ������B"
                Case enmSyoColNo.SyouhinmeiKana:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B40���ȓ��̔��p�J�i�����œ��͂��ĉ������B"
                Case enmSyoColNo.SyouhinmeiKanji:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B30���ȓ��̑S�p��������͂��ĉ������B"
                Case enmSyoColNo.ReceiptmeiKana:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B14���ȓ��̔��p�J�i��������͂��ĉ������B"
                Case enmSyoColNo.ReceiptmeiKanji:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B14���ȓ��̑S�p��������͂��ĉ������B"
                Case enmSyoColNo.SyouhihinmeiHansoku:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B20���ȓ��̑S�p��������͂��ĉ������B"
                
                Case enmSyoColNo.MakerMei:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B10���ȓ��̑S�p��������͂��ĉ������B"
                Case enmSyoColNo.KikakuSize:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B20���ȓ��̑S�p��������͂��ĉ������B"
                Case enmSyoColNo.HinbanKataban:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B20���ȓ��̑S�p��������͂��ĉ������B"
                Case enmSyoColNo.IroGara:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B20���ȓ��̑S�p��������͂��ĉ������B"
                Case enmSyoColNo.Taste:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B20���ȓ��̑S�p��������͂��ĉ������B"
                Case enmSyoColNo.Seisantimei:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B20���ȓ��̑S�p��������͂��ĉ������B"
                Case enmSyoColNo.Tokutyou:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B150���ȓ��̑S�p��������͂��ĉ������B"
                Case enmSyoColNo.SyouhinCopy:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B50���ȓ��̑S�p��������͂��ĉ������B"
                Case enmSyoColNo.TuujyouKakaku:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B������7���̐�������͂��ĉ������B"
                Case enmSyoColNo.MakerKakaku:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B������7���̐�������͂��ĉ������B"
                Case enmSyoColNo.POPLayout1:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B30���ȓ��̑S�p���p��������͂��ĉ������B"
                Case enmSyoColNo.POPLayout2:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B30���ȓ��̑S�p���p��������͂��ĉ������B"
                Case enmSyoColNo.PopSize:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B30���ȓ��̑S�p���p��������͂��ĉ������B"
                
                Case enmSyoColNo.AtukaiPtn:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B3���ȓ��̔��p��������͂��ĉ������B"
                Case enmSyoColNo.HyoujyunBaika:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B������7���̐�������͂��ĉ������B"
                Case enmSyoColNo.SikiriGenka:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B��������7���ȓ��A��������2���ȓ��Ŏw�肵�ĉ������B"
                Case enmSyoColNo.VhJyuyouCd:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B6���ȓ��̔��p��������͂��ĉ������B"
                Case enmSyoColNo.RdJyuyouCd:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B6���ȓ��̔��p��������͂��ĉ������B"
                Case enmSyoColNo.VcsJyuyouCd:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B6���ȓ��̔��p��������͂��ĉ������B"
                Case enmSyoColNo.SvhJyuyouCd:
                    FncGetErrorMsgType_Syo = "�����ɉߕs��������܂��B6���ȓ��̔��p��������͂��ĉ������B"
            End Select
        Case PARAM_NOTEXISTDB:
            Select Case intErrorField
                Case enmSyoColNo.janCD
                    FncGetErrorMsgType_Syo = "JAN�R�[�h�͓���DB�ɓo�^����Ă��܂���B"
            End Select
        Case PARAM_SAMECODEDB:      FncGetErrorMsgType_Syo = "JAN�R�[�h�͊�懂" & strErrorParameter & "�ɓo�^����Ă��܂��B"
        Case PARAM_SAMEKEYDB:       FncGetErrorMsgType_Syo = "�����R�[�h�͊�懂" & strErrorParameter & "�ɓo�^����Ă��܂��B"
    End Select

End Function

' '*< UNISYSTEM >**************************************************'
'*�y�v���V�[�W�����z
'*!    FncGetErrorMsgType
'*�y�����T�v�z
'*?    ���������G���[�̓��e�𕶎���Ŏ擾���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intErrorField ; �G���[�����������t�B�[���h
'*P    intErrorType ; ���������G���[�̎��
'*P    strErrorParameter ; �G���[���b�Z�[�W�⏕�p������
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncGetErrorMsgField; String;���������G���[�̓��e(�t�B�[���h���ƂɈقȂ邱�Ƃ�����)
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/04/19; ���@����(USC); �V�K�쐬;
'*M    2006/09/13; �a�c(IC);        ���i�I��v�e�⍇���o�ŉ��P;
'*M    2009/05/12; �a�c(IC);        �O���[�v�i�̓G���[�ɂ���;
'*M    2009/06/06; Supattra T.(Tostem Thai); Add error message;
'*M    2011/02/10; �a�c(IC);        �Z���^�[�R�[�h�ǉ�;
'*M    2015/09/15; �a�c(IC);        TC�V�X�e���\�z;
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
                    'FncGetErrorMsgType = "�Z���^�[�敪��1�F���~�Z���^�[�̏ꍇ�A�K�{���͍��ڂł��B"
                    FncGetErrorMsgType = "�Z���^�[�敪��1�F���~�Z���^�[�E3�FTC�Z���^�[�̏ꍇ�A�K�{���͍��ڂł��B"
                '********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
                Case enmModColNo.Genka, enmModColNo.Hacchuutani
                    FncGetErrorMsgType = "�Z���^�[�敪��2�FPDC�̏ꍇ�A�K�{���͍��ڂł��B"
                '********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
                Case Else
                    FncGetErrorMsgType = "�K�{���ڂł��B"
            End Select
        Case PARAM_NOTDISIT:        FncGetErrorMsgType = "���p�����œ��͂��ĉ������B"
        Case PARAM_NONARROW:        FncGetErrorMsgType = "���p�����œ��͂��ĉ������B"
        Case PARAM_NOTFINDMASTER:
            Select Case intErrorField
                Case enmModColNo.TenCd
                    FncGetErrorMsgType = "�}�X�^�ɓo�^����Ă��Ȃ��X�R�[�h�E���ƃR�[�h���ݒ肳��Ă��܂��B"
                Case enmModColNo.BaikaRiyuu
                    FncGetErrorMsgType = "�}�X�^�ɓo�^����Ă��Ȃ������ύX���R���ݒ肳��Ă��܂��B"
                Case enmModColNo.Torihikisaki
                    FncGetErrorMsgType = "�}�X�^�ɓo�^����Ă��Ȃ�����悪�ݒ肳��Ă��܂��B"
            End Select

        Case PARAM_WRONGCHECKDIGIT: FncGetErrorMsgType = "�`�F�b�N�f�W�b�g���m�F���ĉ������B"
        Case PARAM_WRONGDATE:       FncGetErrorMsgType = "���t�𐳂������͂��ĉ������BYYYYMMDD"
        
        Case PARAM_SHORTDISIT:
            Select Case intErrorField
                Case enmModColNo.janCD
                    FncGetErrorMsgType = "�����ɉߕs��������܂��B8����������13���̐�������͂��ĉ������B"
                Case enmModColNo.Genka
                    FncGetErrorMsgType = "�����ɉߕs��������܂��B��������7���ȓ��A��������2���ȓ��Ŏw�肵�ĉ������B"
                Case enmModColNo.Baika
                    FncGetErrorMsgType = "�����ɉߕs��������܂��B����7���ȓ��Ŏw�肵�ĉ������B"
                Case enmModColNo.Hacchuutani
                    FncGetErrorMsgType = "�����ɉߕs��������܂��B����5���ȓ��Ŏw�肵�ĉ������B"
                Case enmModColNo.centerCD
                    FncGetErrorMsgType = "�����ɉߕs��������܂��B4���̐�������͂��ĉ������B"
                '2009.06.06 Add by Supattra T.(Tostem Thai)Start.
                Case enmModColNo.MoGenka
                    FncGetErrorMsgType = "�����ɉߕs��������܂��B��������7���ȓ��A��������2���ȓ��Ŏw�肵�ĉ������B"
                Case enmModColNo.MoBaika
                    FncGetErrorMsgType = "�����ɉߕs��������܂��B����7���ȓ��Ŏw�肵�ĉ������B"
                Case enmModColNo.HattyuuTani
                    FncGetErrorMsgType = "�����ɉߕs��������܂��B����5���ȓ��Ŏw�肵�ĉ������B"
                 Case enmModColNo.PopSize
                    FncGetErrorMsgType = "���p20�����͑S�p10���ȓ��œ��͂��ĉ������B"
                '2009.06.06 Add by Supattra T.(Tostem Thai)End.
            End Select
        
        Case PARAM_NOTEXISTDB:
            Select Case intErrorField
                Case enmTenColNo.TenCd
                    FncGetErrorMsgType = "�}�X�^�ɓo�^����Ă��Ȃ��X�R�[�h�E���ƃR�[�h���ݒ肳��Ă��܂��B"
                Case enmTenColNo.janCD
                    FncGetErrorMsgType = "JAN�R�[�h�͓���DB�ɓo�^����Ă��܂���B"
                Case enmTenColNo.BaikaRiyuu
                    FncGetErrorMsgType = "�}�X�^�ɓo�^����Ă��Ȃ������ύX���R���ݒ肳��Ă��܂��B"
                Case enmTenColNo.Torihikisaki
'                    FncGetErrorMsgType = "�}�X�^�ɓo�^����Ă��Ȃ�����悪�ݒ肳��Ă��܂��"
                    FncGetErrorMsgType = "�}�X�^�ɓo�^����Ă��Ȃ����A�܂��͎������~����Ă������悪�ݒ肳��Ă��܂��B"
                Case enmTenColNo.CenterKbn
                    FncGetErrorMsgType = "�}�X�^�ɓo�^����Ă��Ȃ��Z���^�[�敪���ݒ肳��Ă��܂��B"
                Case enmModColNo.centerCD
                    FncGetErrorMsgType = "�}�X�^�ɓo�^����Ă��Ȃ��Z���^�[�R�[�h���ݒ肳��Ă��܂��B"
            End Select
        Case PARAM_SAMECODEDB:      FncGetErrorMsgType = "JAN�R�[�h�͊�懂" & strErrorParameter & "�ɓo�^����Ă��܂��B"
        Case PARAM_SAMEKEYDB:       FncGetErrorMsgType = "�����R�[�h�͊�懂" & strErrorParameter & "�ɓo�^����Ă��܂��B"
        Case PARAM_TORIHIKISTOP:    FncGetErrorMsgType = "�����Ƃ̎���͒�~����Ă��܂��B"
        Case PARAM_REPDATEBAIKA:    FncGetErrorMsgType = "���ɓ���DB(�����ύX���)�̐�t�ύX���ɓo�^����Ă�����t�ł��B"
        Case PARAM_REPDATESIIRE:    FncGetErrorMsgType = "���ɓ���DB(�d���ύX���)�̐�t�ύX���ɓo�^����Ă�����t�ł��B"
        Case PARAM_REPDATETENPO:    FncGetErrorMsgType = "���ɓ���DB(�X�܃Z���^�[�R�[�h�ύX���)�̐�t�ύX���ɓo�^����Ă�����t�ł��B"

'********** 2005/05/07 �k�V **********��
        Case PARAM_ZERO: FncGetErrorMsgType = "�O���傫���l���w�肵�ĉ������B"
'********** 2005/05/07 �k�V **********��
    
'********** 2005/06/16 �k�V **********��
        Case PARAM_BAIKARIYUU_HISSU: FncGetErrorMsgType = "�X�R�[�h�u0000�v�̏ꍇ�A���������͂���Ă�����A�����ύX���R�͓��͕K�{�ł��B"
        Case PARAM_BAIKARIYUU_FUKA:  FncGetErrorMsgType = "�X�R�[�h���u0000�v�ȊO�̏ꍇ�A���͔����������͂̏ꍇ�́A�����ύX���R�͓��͕s�ł��B"
'********** 2005/06/16 �k�V **********��
        Case PARAM_GROUP:    FncGetErrorMsgType = "�O���[�v�i�͓o�^�ł��܂���B"
        '2015/09/15 wadak7 upd
        'Case PARAM_KINSI:    FncGetErrorMsgType = "�Z���^�[�敪��1�F���~�Z���^�[�ȊO�̏ꍇ�A���͂ł��܂���B"
        Case PARAM_KINSI:    FncGetErrorMsgType = "�Z���^�[�敪��1�F���~�Z���^�[�E3�FTC�Z���^�[�ȊO�̏ꍇ�A���͂ł��܂���B"
'********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��
        Case PARAM_HONBU_NAI:       FncGetErrorMsgType = "�Z���^�[�敪��2�FPDC�̏ꍇ�A0000�F�{���̃f�[�^�͕K�{���͂ł��B�ύX�� " & Split(strErrorParameter, ",")(0) & " JAN�R�[�h " & Split(strErrorParameter, ",")(1)
        Case PARAM_HATTYUU_TANI:    FncGetErrorMsgType = "�Z���^�[�敪��2�FPDC�̔����P�ʂ͑S�ē����l����͂��ĉ������B�ύX�� " & Split(strErrorParameter, ",")(0) & " JAN�R�[�h " & Split(strErrorParameter, ",")(1)
        Case PARAM_GENKA:           FncGetErrorMsgType = "�Z���^�[�敪��2�FPDC�̌����͑S�ē����l����͂��ĉ������B�ύX�� " & Split(strErrorParameter, ",")(0) & " JAN�R�[�h " & Split(strErrorParameter, ",")(1)
        Case PARAM_GENKA_HATTYUUTANI: FncGetErrorMsgType = "�X�V�X�܈ȊO��PDC�敪�ɂȂ��Ă���X�܂�����܂��B�m�F�̏�A�čX�V���ĉ������B�ύX�� " & Split(strErrorParameter, ",")(0) & " JAN�R�[�h " & Split(strErrorParameter, ",")(1)
'********** 2012/08/03 �k�o�i��A�j���̓`�F�b�N�̒ǉ� **********��

    End Select

End Function

'********** 2005/04/27 �k�V **********��
Rem ========================================================================
Rem �@�\�����F VMST��ʊ�{�e�[�u����菤�i�����擾���AVSWF�\���d�l�e�[�u���Ƀ��R�[�h��ǉ�����
Rem �������F �G���[���b�Z�[�W, CSV�s�J�E���g, ��懂, JAN�R�[�h, ���O�C�����[�U�[ID
Rem �߂�l�@�F TRUE=����, FALSE=���s
Rem ========================================================================
Public Function Insert_t_syousen_hyouji_siyou(ByRef strErrMsg _
                                      , ByVal intN As Integer _
                                      , ByVal strKikakuSyuruiCD As String _
                                      , ByVal strNenRenNO As String _
                                      , ByVal strEdaNO As String _
                                      , ByVal strJanCd As String _
                                      , ByVal strUserId As String) As Boolean
On Error GoTo Err_Insert_t_syousen_hyouji_siyou

    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess") '���ʊ֐�(SQLServer�ڑ��p)
    Dim rsAdo1 As ADODB.Recordset           '���R�[�h�Z�b�g
    Dim rsAdo2 As ADODB.Recordset           '���R�[�h�Z�b�g
    Dim strMethodCallInfo As String         '�G���[���
    Dim strSQL As String                    '��Ɨp������
    Dim arrParam As Object                 '�L�[���i�[�p�z��
    Const strUdl As String = "VSWF"         'SQLServer�ڑ��pUDL�t�@�C����

    '�߂�l�̏����l���Z�b�g����
    Insert_t_syousen_hyouji_siyou = False
    
    '�G���[�����쐬���Ă���
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.Insert_t_syousen_hyouji_siyou(��懂:" & strKikakuSyuruiCD & strNenRenNO & "-" & strEdaNO & "JAN�R�[�h:" & strJanCd & ")"
    
    '���ʊ֐��I�u�W�F�N�g���쐬����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'VSWF�\���d�l�e�[�u���̃��R�[�h�Z�b�g���擾����
    '(TABLE: �\���d�l�e�[�u��)
    '(FIELDS: *)
    '(KEY:��懂, JAN�R�[�h)
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
    
    'VSWF�\���d�l�e�[�u���Ƀ��R�[�h�����ɑ��݂����ꍇ�͏������I������(�G���[�ł͖���)
    If rsAdo1.EOF = False Then
        
        ''�f�[�^�x�[�X�����
            rsAdo1 = Nothing
        objCDBAccess = Nothing
        'GetObjectContext.SetComplete
        
        ''�߂�l���Z�b�g����(�G���[�ł͖����̂�TRUE��Ԃ�)
        Insert_t_syousen_hyouji_siyou = True
        GoTo Ext_Insert_t_syousen_hyouji_siyou
    
    'VSWF�\���d�l�e�[�u���Ƀ��R�[�h�����݂��Ȃ��ꍇ��VMST��ʊ�{�e�[�u����菤�i�����擾��VSWF�\���d�l�e�[�u���Ƀ��R�[�h��ǉ�����
    Else
   
        ''VMST��ʊ�{�e�[�u���̃��R�[�h�Z�b�g���擾����
        ''(TABLE: ��ʊ�{�r���[�A��L�[�}�X�^)
        ''(FIELDS: *)
        ''(KEY:JAN�R�[�h)
        strSQL = _
        "SELECT v_ippan_kihon.*" & _
        "  FROM v_ippan_kihon WITH(READCOMMITTED)" & _
        " INNER JOIN v_key_1 WITH(READCOMMITTED) ON v_ippan_kihon.key_cd = v_key_1.key_cd" & _
        " WHERE v_key_1.jan_cd = ?"

        arrParam = Array(MakeParam("@v_key_1.jan_cd", DT_VarChar, 20, strJanCd))
    
            rsAdo2 = objCDBAccess.GetDCRecordset(strUdl, strSQL, arrParam)
    
        ''VMST��ʊ�{�e�[�u���Ƀ��R�[�h�����݂��Ȃ������珈�����I������(�\�����Ȃ��G���[)
        If rsAdo2.EOF Then
            
            strErrMsg = CStr(intN + 1) & "�s�ځFVMST��ʊ�{�e�[�u���Ƀf�[�^�����݂��Ȃ���VSWF�\���d�l�e�[�u���Ƀ��R�[�h��ǉ��o���܂���ł����B"
            GoTo Err_Insert_t_syousen_hyouji_siyou
    
        ''VMST��ʊ�{�e�[�u������擾�����f�[�^��VSWF�\���d�l�e�[�u���Ƀ��R�[�h��ǉ�����
        ''(TABLE: �\���d�l�e�[�u��)
        ''(FIELDS: ����ރR�[�h, �N�A��, �}��, JAN�R�[�h
        ''       , �\���p���i��(�J�i), �\���p���i��(����), ���V�[�g��(�J�i), ���V�[�g��(����)
        ''       , ���i��(����), ���[�J��(����), �K�i/�T�C�Y(����), �i�ԁ^�^��(����)
        ''       , �F�^��(����), �e�[�X�g(����), ���Y�n��(����), ����, �ܖ�����/�g�p�����L��
        ''       , �G�R�}�[�N���i, �O���[���w���@�K�����i, GPN�f�[�^�x�[�X�f�ڏ��i
        ''       , �o�^UserID, �o�^�N���������b
        ''(KEY:��懂, JAN�R�[�h)
        Else
            
            '''�ؒf���R�[�h�Z�b�g���X�V����
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
            
            '''�f�[�^�x�[�X���X�V����
            Call objCDBAccess.UpdateDCRecordset(strUdl, rsAdo1)
            
            '''�߂�l���Z�b�g����
            Insert_t_syousen_hyouji_siyou = True
        
        End If
    End If
    
    '�f�[�^�x�[�X�����
        rsAdo1 = Nothing
        rsAdo2 = Nothing
    objCDBAccess = Nothing
    'GetObjectContext.SetComplete

    '�G���[���̏���
Ext_Insert_t_syousen_hyouji_siyou:
    Exit Function
Err_Insert_t_syousen_hyouji_siyou:
    If Not rsAdo1 Is Nothing Then Set rsAdo1 = Nothing
    If Not rsAdo2 Is Nothing Then Set rsAdo2 = Nothing
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
'********** 2005/05/20 �y�c �ύX**********��
'    Call ErrRaise(DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description)
    Call ErrRaise(DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B" & strErrMsg, 'Err.Number, Err.Source, Err.Description)
'********** 2005/05/20 �y�c �ύX**********��
    Resume Ext_Insert_t_syousen_hyouji_siyou

End Function
'********** 2005/04/27 �k�V **********��

'********** 2005/05/07 �k�V **********��
Rem ========================================================================
Rem �@�\�����F ����P���Ƃ̃��R�[�h���쐬����ہA���̑S�Ă̎��Ƃ̃��R�[�h���쐬
Rem �������F ��懂, JAN�R�[�h, �쐬�������ƃR�[�h, ���O�C�����[�U�[ID
Rem �߂�l�@�F TRUE=����, FALSE=���s
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
    
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess") '���ʊ֐�(SQLServer�ڑ��p)
    Dim rsAdo1 As ADODB.Recordset           '���R�[�h�Z�b�g
    Dim rsAdo2 As ADODB.Recordset           '���R�[�h�Z�b�g
    Dim rsAdo3 As ADODB.Recordset           '���R�[�h�Z�b�g
    Dim strMethodCallInfo As String         '�G���[���
    Dim strSQL As String                    '��Ɨp������
    Const strUdl As String = "VSWF"         'SQLServer�ڑ��pUDL�t�@�C����
    
    '�߂�l�̏����l���Z�b�g����
    Insert_t_syousen_jigyoubetu_settei = False
    
    '�G���[�����쐬���Ă���
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.Insert_t_syousen_jigyoubetu_settei(��懂:" & strKikakuSyuruiCD & strNenRenNO & "-" & strEdaNO & "JAN�R�[�h:" & strJanCd & ")"
    
    '���ʊ֐��I�u�W�F�N�g���쐬����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")

    'VRTL���Ə��}�X�^����S���ƃR�[�h���擾����
    '(TABLE: ���Ə��}�X�^)
    '(FIELDS: ���ƃR�[�h)
    '(����: �X����NULL���͈ꊇ�o�^���ȍ~�CKEY�敪=W02�C�ݒ���e��4���ڂ�1�j
    strSQL = _
    "SELECT DISTINCT v_jigyousyo.jigyou_cd" & _
    "  FROM v_jigyousyo WITH(READCOMMITTED)" & _
    " INNER JOIN m_wf_meisyou WITH(READCOMMITTED)" & _
    "    ON v_jigyousyo.jigyou_cd = m_wf_meisyou.key_cd" & _
    " WHERE SUBSTRING(m_wf_meisyou.settei_naiyou,4,1) = '1'" & _
    "   AND m_wf_meisyou.key_kbn = '" & KEYCD_KIKAKU_DOUNYU_TAISYOU & "'"
    
        rsAdo1 = objCDBAccess.GetDCRecordset(strUdl, strSQL)

    '���ƃR�[�h�������ꍇ�͏������I������(���蓾�Ȃ��A�G���[�ɂ͂��Ȃ�)
    If rsAdo1.EOF = True Then
        ''�f�[�^�x�[�X�����
            rsAdo1 = Nothing
        objCDBAccess = Nothing
        'GetObjectContext.SetComplete
        ''�߂�l���Z�b�g����(�G���[�ł͖����̂�TRUE��Ԃ�)
        Insert_t_syousen_jigyoubetu_settei = True
        GoTo Ext_Insert_t_syousen_jigyoubetu_settei
    End If

    'CSV�ɓ��͂���Ă���P���Ƃ������AVSWF���ƕʐݒ�e�[�u���Ƀ��R�[�h��ǉ�����
    Do While Not rsAdo1.EOF
        If Trim(rsAdo1("jigyou_cd")) <> strJigyouCd Then
            
            ''VSWF���ƕʐݒ�e�[�u���̃��R�[�h�Z�b�g���擾����
            ''(TABLE: ���ƕʐݒ�e�[�u��)
            ''(FIELDS: *)
            ''(KEY:��懂, JAN�R�[�h, ���ƃR�[�h)
            strSQL = _
            "SELECT *" & _
            "  FROM t_syousen_jigyoubetu_settei" & _
            " WHERE kikaku_syurui_cd = '" & strKikakuSyuruiCD & "'" & _
            "   AND nen_ren_no = '" & strNenRenNO & "'" & _
            "   AND eda_no = '" & strEdaNO & "'" & _
            "   AND jan_cd = '" & strJanCd & "'" & _
            "   AND jigyou_cd = '" & Trim(rsAdo1("jigyou_cd")) & "'"
    
                rsAdo2 = objCDBAccess.GetDCRecordset(strUdl, strSQL)
    
            ''VSWF���ƕʐݒ�e�[�u���Ƀ��R�[�h�����ɑ��݂����ꍇ�͏������I������(�G���[�ł͖���)
            If rsAdo2.EOF = False Then
                '''�f�[�^�x�[�X�����
                    rsAdo1 = Nothing
                    rsAdo2 = Nothing
                objCDBAccess = Nothing
                'GetObjectContext.SetComplete
                '''�߂�l���Z�b�g����(�G���[�ł͖����̂�TRUE��Ԃ�)
                Insert_t_syousen_jigyoubetu_settei = True
                GoTo Ext_Insert_t_syousen_jigyoubetu_settei
            End If
            
            ''VMST���ƕʃe�[�u����背�R�[�h�Z�b�g���擾����
            ''(TABLE: ���ƕʃr���[�A��L�[�}�X�^)
            ''(FIELDS: *)
            ''(KEY:JAN�R�[�h, ���ƃR�[�h)
'********** 2005/05/20 �y�c **********��
'�����I�����̃t�B�[���h���R���o�[�g��������悤�ύX
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
    
            ''VMST���ƕʃe�[�u���Ƀ��R�[�h�����݂��Ȃ������珈�����I������(�G���[�ɂ͂��Ȃ�)
            If rsAdo3.EOF Then
                '''�f�[�^�x�[�X�����
                    rsAdo1 = Nothing
                    rsAdo2 = Nothing
                    rsAdo3 = Nothing
                objCDBAccess = Nothing
                'GetObjectContext.SetComplete
                '''�߂�l���Z�b�g����(�G���[�ł͖����̂�TRUE��Ԃ�)
                Insert_t_syousen_jigyoubetu_settei = True
                GoTo Ext_Insert_t_syousen_jigyoubetu_settei
            End If
            
            ''VSWF���ƕʐݒ�e�[�u���Ƀ��R�[�h��ǉ�����
            ''(TABLE: ���ƕʐݒ�e�[�u��)
            ''(FIELDS: *)
            ''(KEY:��懂, JAN�R�[�h, ���ƃR�[�h)
            '''�ؒf���R�[�h�Z�b�g���X�V����
            With rsAdo2
                .AddNew
                .Fields("kikaku_syurui_cd") = strKikakuSyuruiCD
                .Fields("nen_ren_no") = strNenRenNO
                .Fields("eda_no") = strEdaNO
                .Fields("jan_cd") = strJanCd
                .Fields("jigyou_cd") = Trim(rsAdo1("jigyou_cd"))
                .Fields("hattyuu_kaisi_date") = rsAdo3("CONVERT_hattyuu_kaisi_date")
'********** 2005/05/20 �y�c �ύX**********��
'                .Fields("hattyuu_syuuryou_date") = rsAdo3("hattyuu_syuuryou_date")
                .Fields("hattyuu_syuuryou_date") = rsAdo3("CONVERT_hattyuu_syuuryou_date")
'********** 2005/05/20 �y�c �ύX**********��
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

            '''�f�[�^�x�[�X���X�V����
            Call objCDBAccess.UpdateDCRecordset(strUdl, rsAdo2)
                rsAdo2 = Nothing
                rsAdo3 = Nothing
        
        End If
        rsAdo1.MoveNext
    Loop

    '�߂�l���Z�b�g����
    Insert_t_syousen_jigyoubetu_settei = True
    
    '�f�[�^�x�[�X�����
        rsAdo1 = Nothing
    objCDBAccess = Nothing
    'GetObjectContext.SetComplete

    '�G���[���̏���
Ext_Insert_t_syousen_jigyoubetu_settei:
    Exit Function
Err_Insert_t_syousen_jigyoubetu_settei:
    If Not rsAdo1 Is Nothing Then Set rsAdo1 = Nothing
    If Not rsAdo2 Is Nothing Then Set rsAdo1 = Nothing
    If Not rsAdo3 Is Nothing Then Set rsAdo1 = Nothing
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    Call ErrRaise(DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description)
    Resume Ext_Insert_t_syousen_jigyoubetu_settei

End Function
'********** 2005/05/07 �k�V **********��

'********** 2005/05/07 �k�V **********��
Rem ========================================================================
Rem �@�\�����F ���z�ނ����̒l�����`�F�b�N����
Rem �������F CSV�f�[�^�z��, CSV�t�@�C���̌��݂̍s
Rem �߂�l�@�F TRUE=�G���[����, FALSE=�G���[�L��
Rem ========================================================================
Private Function FncCheck_NotZero(ByRef arryCSVRecord() As String _
                          , ByVal intNowLine As Integer _
                          , ByVal strTenHac As String) As Boolean

'�߂�l�̏����l���Z�b�g����
    FncCheck_NotZero = True
    
'------------------------------
'���z�ނ������l�����`�F�b�N����
'------------------------------
    ''�X�ʈꊇ�o�^(TEN)�̏ꍇ
    If strTenHac = "TEN" Then
        '''����
        
'********** 2005/05/18 �c�� �ύX **********��
'        If arryCSVRecord(enmTenColNo.Genka) <> "" Then
        If arryCSVRecord(enmTenColNo.Genka) <> "" And IsNumeric(arryCSVRecord(enmTenColNo.Genka)) Then
'********** 2005/05/18 �c�� �ύX **********��
            
            If CCur(arryCSVRecord(enmTenColNo.Genka)) <= 0 Then
                FncSetError (intNowLine, enmTenColNo.Genka, PARAM_ZERO
                FncCheck_NotZero = False
            End If
        End If
        '''����
        
'********** 2005/05/18 �c�� �ύX **********��
'        If arryCSVRecord(enmTenColNo.Baika) <> "" Then
        If arryCSVRecord(enmTenColNo.Baika) <> "" And IsNumeric(arryCSVRecord(enmTenColNo.Baika)) Then
'********** 2005/05/18 �c�� �ύX **********��
           
           If CCur(arryCSVRecord(enmTenColNo.Baika)) <= 0 Then
                FncSetError (intNowLine, enmTenColNo.Baika, PARAM_ZERO
                FncCheck_NotZero = False
            End If
        End If
        '''�����P��
        
'********** 2005/05/18 �c�� �ύX **********��
'        If arryCSVRecord(enmTenColNo.Hacchuutani) <> "" Then
        If arryCSVRecord(enmTenColNo.Hacchuutani) <> "" And IsNumeric(arryCSVRecord(enmTenColNo.Hacchuutani)) Then
'********** 2005/05/18 �c�� �ύX **********��
            
            If CCur(arryCSVRecord(enmTenColNo.Hacchuutani)) <= 0 Then
                FncSetError (intNowLine, enmTenColNo.Hacchuutani, PARAM_ZERO
                FncCheck_NotZero = False
            End If
        End If
    ''�������~���ꊇ�o�^(HAC)�̏ꍇ
    Else
        '''�Ώۍ��ږ���
    End If

'------------------------------
'�`�F�b�N��ǉ�����ꍇ�͂����ɋL�q����
'------------------------------
'********** 2005/06/16 �k�V **********��
'�����ύX���R�́A�X=0000������<>""�̏ꍇ�͓��͕K�{�A����ȊO�̏ꍇ�͓��͕s��
    ''�X�ʈꊇ�o�^(TEN)�̏ꍇ�̂ݎ��s
    If strTenHac = "TEN" Then
        
        '''�X=0000�̏ꍇ
        If arryCSVRecord(enmTenColNo.TenCd) = "0000" Then
            ''''���������͂���Ă�����A�����ύX���R�͓��͕K�{
            If arryCSVRecord(enmTenColNo.Baika) <> "" Then
                If arryCSVRecord(enmTenColNo.BaikaRiyuu) = "" Then
                    FncSetError (intNowLine, enmTenColNo.BaikaRiyuu, PARAM_BAIKARIYUU_HISSU
                    FncCheck_NotZero = False
                End If
            End If
            ''''�����������͂�������A�����ύX���R�͓��͕s��
            If arryCSVRecord(enmTenColNo.Baika) = "" Then
                If arryCSVRecord(enmTenColNo.BaikaRiyuu) <> "" Then
                    FncSetError (intNowLine, enmTenColNo.BaikaRiyuu, PARAM_BAIKARIYUU_FUKA
                    FncCheck_NotZero = False
                End If
            End If
        End If
        
        '''�X<>0000�̏ꍇ
        If arryCSVRecord(enmTenColNo.TenCd) <> "0000" Then
            ''''�����ύX���R�͓��͕s��
            If arryCSVRecord(enmTenColNo.BaikaRiyuu) <> "" Then
                FncSetError (intNowLine, enmTenColNo.BaikaRiyuu, PARAM_BAIKARIYUU_FUKA
                FncCheck_NotZero = False
            End If
        End If
    
    End If
'********** 2005/06/16 �k�V **********��

End Function
'********** 2005/05/07 �k�V **********��
'
'*< TOSTEM >*******************************************************************'
'*
'*�y�v���V�[�W�����z
'*!     FncBunTouroku
'*
'*�y�����T�v�z
'*?     ���ޏ��ꊇ�ύX�̃f�[�^�����i�I��WF��DB�Ɏ�荞��
'*
'*�y�p�����[�^�z
'*      Name                   Type                Comment
'*P     strErrMsg              Object             �G���[���b�Z�[�W;
'*�y�߂�l�z
'*      Name                   Type                Comment
'*R     FncBunTouroku          Boolean;
'*
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)            EDP-No      ����
'*M 2005/08/31; ���� (IC);             P-24884;    �V�K�쐬;
'*M 2006/09/14; �a�c(IC);              P-XXXXX;    ���i�I��v�e�⍇���o�ŉ��P;
'*
'*< Comment End >**************************************************************
Public Function FncBunTouroku(ByRef strErrMsg) As Boolean
    
    Dim intN As Integer
    Dim blnRet As Boolean
    Dim strJanCd As String
    Dim intNowLine As Integer
    Dim strCSVRecord() As String
    Dim flgChkError As Boolean      '�`�F�b�N���̃G���[�t���O�iTRUE:�G���[����j
    
    FncBunTouroku = False
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    
    'DB�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h�Z�b�g�̐���
    Dim strSQL          As String
    Dim objCDBAccess    As Object
        
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")

    '�ꊇ�o�^���A�������[�U�[�̂ݎ��s�ł���
    Call LockTableIkkatu(objCDBAccess)
    
    Call SetIkkatuLog("1-1", "���ޏ��ꊇ�ύX�̃f�[�^�����i�I��WF��DB�Ɏ�荞��", "FncBunTouroku", "", "", pstrTantousyaId, "INS")
    Call SetIkkatuLog("1-1-1", "�d��JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")
    '�d��JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(�d��JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����)"
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
    pChkSameJANCodeRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    pChkSameJANCodeRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    Call SetIkkatuLog("1-1-1", "", "", "", "OK", "", "UPD")
    
    Call SetIkkatuLog("1-1-2", "CSV�t�@�C����Jan�R�[�h�̃��X�g(Jan�R�[�h�̏d������)�쐬", "SQL", "", "", pstrTantousyaId, "INS")
    'CSV�t�@�C����Jan�R�[�h�̃��X�g(Jan�R�[�h�̏d������)�쐬
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(CSV�t�@�C����Jan�R�[�h�̃��X�g(Jan�R�[�h�̏d������)�쐬)"

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
        Call SetIkkatuLog("1-1-3", "CSV�t�@�C����JAN�R�[�h�R�t����L�[�̃��X�g(��L�[�d������)�쐬", "SQL", "", "", pstrTantousyaId, "INS")
    '�����}�X�^�o�^�ς�JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(�����}�X�^�o�^�ς�JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����)"

    strSQL = "SELECT jan_cd, key_cd FROM v_key_1 WITH(readcommitted)"
    strSQL = strSQL & " WHERE jan_cd IN(" & pCsvJanCodeList & ") "
    
        pChkvk1JANCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pChkvk1JANCodeRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
        Call SetIkkatuLog("1-1-3", "", "", "", "OK", "", "UPD")
        Call SetIkkatuLog("1-1-4", "CSV�t�@�C����JAN�R�[�h�R�t����L�[�̃��X�g(��L�[�d������)�쐬", "SQL", "", "", pstrTantousyaId, "INS")
    '�����}�X�^�o�^�ς�JAN�R�[�h�R�t����L�[�̃��X�g(��L�[�d������)�쐬
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(CSV�t�@�C����JAN�R�[�h�R�t����L�[�̃��X�g(��L�[�d������)�쐬)"
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
        Call SetIkkatuLog("1-1-5", "CSV�t�@�C������1���ރR�[�h�̃��X�g(��1���ރR�[�h�̏d������)�쐬", "SQL", "", "", pstrTantousyaId, "INS")
    'CSV�t�@�C������1���ރR�[�h�̃��X�g(��1���ރR�[�h�̏d������)�쐬
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(CSV�t�@�C������1���ރR�[�h�̃��X�g(��1���ރR�[�h�̏d������)�쐬)"

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
        Call SetIkkatuLog("1-1-6", "�i��\���f�[�^�o�^�ςݕ��ރR�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")
    '�i��\���f�[�^�o�^�ςݕ��ރR�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(���ރ}�X�^KEY�`�F�b�N�p�ؒf���R�[�h�Z�b�g����)"

    strSQL = "SELECT key_bnrui_cd_1, key_bnrui_cd_2, key_bnrui_cd_3, uriba_cd, old_bumon_cd"
    strSQL = strSQL & " ,old_tyuubunrui_cd, old_syoubunrui_cd, bunrui_cd_3"
    strSQL = strSQL & " FROM v_hinsyu_hyouji WITH(readcommitted)"
    strSQL = strSQL & " WHERE key_bnrui_cd_1 IN(" & pCsvBunrui1CodeList & ") "
    
        pChkBunruiCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    With pChkBunruiCodeRS
            .Fields("key_bnrui_cd_1").Properties("Optimize") = True '�I�v�e�B�}�C�U
            .Fields("key_bnrui_cd_2").Properties("Optimize") = True '�I�v�e�B�}�C�U
            .Fields("key_bnrui_cd_3").Properties("Optimize") = True '�I�v�e�B�}�C�U
    End With
    Call SetIkkatuLog("1-1-6", "", "", "", "OK", "", "UPD")
    Call SetIkkatuLog("1-1-7", "�S�Ă̔z��f�[�^�̃`�F�b�N���s���B", "SQL", "", "", pstrTantousyaId, "INS")
    '�S�Ă̔z��f�[�^�̃`�F�b�N���s���B
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(�S�Ă̔z��f�[�^�̃`�F�b�N���s���B)"
    
    flgChkError = False
    intNowLine = 0
    strErrMsg = ""
    Do Until UBound(paryCsvData) < intNowLine
        
        strCSVRecord = paryCsvData(intNowLine)
        
        '(1)DB�Ɉˑ����Ȃ��`�F�b�N�FFncCheckParameter_CorrectForm_Bun
        '(2)DB���Q�Ƃ��ă`�F�b�N:FncCheckParameter_DataBase_Bun
        If FncCheckParameter_CorrectForm_Bun(strCSVRecord, intNowLine + 1) = False Or _
            FncCheckParameter_DataBase_Bun(strCSVRecord, intNowLine + 1) = False Then
            flgChkError = True
        End If
        
        intNowLine = intNowLine + 1
    Loop
        Call SetIkkatuLog("1-1-7", "", "", "", "OK", "", "UPD")
        Call SetIkkatuLog("1-1-8", "DB�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h�Z�b�g�̊J��", "SQL", "", "", pstrTantousyaId, "INS")
    'DB�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h�Z�b�g�̊J��
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(DB�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h�Z�b�g�̊J��)"
        
    '�`�F�b�N�p�̃I�u�W�F�N�g���J������B
    If Not pChkSameJANCodeRS Is Nothing Then
        pChkSameJANCodeRS.Close
            pChkSameJANCodeRS = Nothing
    End If
    Call SetIkkatuLog("1-1-8", "", "", "", "OK", "", "UPD")
    
    '�f�[�^�ɂP�ł���肪����ꍇ�́A�����ŏI���I�I�I
    If flgChkError = True Then
        '�G���[������΃G���[���b�Z�[�W���擾����B
        strErrMsg = FncGetErrorMsg_Bun()
        
        Exit Function
    End If
    Call SetIkkatuLog("1-2", "�f�[�^�X�V�����J�n", "SQL", "", "", pstrTantousyaId, "INS")
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(�f�[�^�X�V�����J�n)"
    
    '�S�Ẵe�[�u�����X�V����B
    Call subIns_SetAllTables_Bun
    Call SetIkkatuLog("1-2", "", "", "", "OK", "", "UPD")
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncBunTouroku(�f�[�^�X�V�����I��)"
    
    FncBunTouroku = True
    
    '�I�u�W�F�N�g�̊J��
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
    
    '�I�u�W�F�N�g�̊J��
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
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description

End Function


'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckParameter_CorrectForm_Bun
'*�y�����T�v�z
'*?    CSV�t�@�C���̃p�����[�^�ɑ΂��āADB�Ɉˑ����Ȃ��`���`�F�b�N���s���܂��B
'*?    �i���ޏ��ꊇ�ύX�p�`�F�b�N�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckParameter_CorrectForm_Bun; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/08/31; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
'
Private Function FncCheckParameter_CorrectForm_Bun(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheckParameter_CorrectForm_Bun = False
    
    Dim blnIsOKRequired
    Dim blnIsOKNumeric
    Dim blnIsOKNarrow
    Dim blnIsOKLimit
    Dim blnIsOKCheckDigit
    
    'a.�K�{���ځi�ύX���AJAN�R�[�h�A��1���ރR�[�h�A��2���ރR�[�h�A��3���ރR�[�h�j�������Ă��邩�m�F����B
    blnIsOKRequired = FncCheck_RequiredParameter_Bun(arryCSVRecord(), intNowLine)
    
    'b.�ύX���AJAN�R�[�h�ɂ����āA���p�����ł��邩�m�F����B
    blnIsOKNumeric = FncCheck_IsParameterNumeric_Bun(arryCSVRecord(), intNowLine)
    
    'c.��1���ރR�[�h�A��2���ރR�[�h�A��3���ރR�[�h�ɂ����āA���p�����ł��邩�m�F����B
    blnIsOKNarrow = FncCheck_IsParameterNarrow_Bun(arryCSVRecord(), intNowLine)
    
    'd.�����������𒴂��Ă��Ȃ����m�F����B
    blnIsOKLimit = FncCheck_Limit_Bun(arryCSVRecord(), intNowLine)

    'e.JAN�R�[�h�ɐ������`�F�b�N�f�W�b�g�̒l���Z�b�g����Ă��邩�𒲂ׂ�
    blnIsOKCheckDigit = FncCheck_CheckDigit(arryCSVRecord(enmBunColNo.janCD), intNowLine, 5)
    
    If blnIsOKRequired And blnIsOKNumeric And blnIsOKNarrow And _
       blnIsOKLimit And blnIsOKCheckDigit Then
       
       FncCheckParameter_CorrectForm_Bun = True
    End If
    
End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_RequiredParameter_Bun
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�K�{���ڂ��󔒂łȂ����ǂ����`�F�b�N���܂��B
'*?    �i���ޏ��ꊇ�ύX�p�`�F�b�N�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_RequiredParameter_Bun; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/08/31; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_RequiredParameter_Bun(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_RequiredParameter_Bun = True
    
    If arryCSVRecord(enmBunColNo.HenkouDate) = "" Then   '�ύX��
        FncSetError (intNowLine, enmBunColNo.HenkouDate, PARAM_EMPTY
        FncCheck_RequiredParameter_Bun = False
    End If
    If arryCSVRecord(enmBunColNo.janCD) = "" Then        'JAN�R�[�h
        FncSetError (intNowLine, enmBunColNo.janCD, PARAM_EMPTY
        FncCheck_RequiredParameter_Bun = False
    End If
    If arryCSVRecord(enmBunColNo.Bunrui1) = "" Then      '���ރR�[�h1
        FncSetError (intNowLine, enmBunColNo.Bunrui1, PARAM_EMPTY
        FncCheck_RequiredParameter_Bun = False
    End If
    If arryCSVRecord(enmBunColNo.Bunrui2) = "" Then      '���ރR�[�h2
        FncSetError (intNowLine, enmBunColNo.Bunrui2, PARAM_EMPTY
        FncCheck_RequiredParameter_Bun = False
    End If
    If arryCSVRecord(enmBunColNo.Bunrui3) = "" Then      '���ރR�[�h3
        FncSetError (intNowLine, enmBunColNo.Bunrui3, PARAM_EMPTY
        FncCheck_RequiredParameter_Bun = False
    End If

End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_IsParameterNumeric_Bun
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�Ώۃt�B�[���h�����p���l���ǂ����`�F�b�N���܂��B
'*?    �i���ޏ��ꊇ�ύX�p�`�F�b�N�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_IsParameterNumeric_Bun; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/08/31; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterNumeric_Bun(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_IsParameterNumeric_Bun = True

    If FncChkHankakuNumber(arryCSVRecord(enmBunColNo.HenkouDate)) = False Then      '�ύX��
        FncSetError (intNowLine, enmBunColNo.HenkouDate, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Bun = False
    End If
    
    If FncChkHankakuNumber(arryCSVRecord(enmBunColNo.janCD)) = False Then           'JAN�R�[�h
        FncSetError (intNowLine, enmBunColNo.janCD, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Bun = False
    End If
    
End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_IsParameterNarrow_Bun
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�Ώۃt�B�[���h�����p�������ǂ����`�F�b�N���܂��B
'*?    �i���ޏ��ꊇ�ύX�p�`�F�b�N�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_IsParameterNarrow_Bun; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/08/31; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterNarrow_Bun(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_IsParameterNarrow_Bun = True

    If fncChkHankakuEisuu(arryCSVRecord(enmBunColNo.Bunrui1)) = False Then      '��1���ރR�[�h
        FncSetError (intNowLine, enmBunColNo.Bunrui1, PARAM_HANEISUU
        FncCheck_IsParameterNarrow_Bun = False
    End If
    
    If fncChkHankakuEisuu(arryCSVRecord(enmBunColNo.Bunrui2)) = False Then      '��2���ރR�[�h
        FncSetError (intNowLine, enmBunColNo.Bunrui2, PARAM_HANEISUU
        FncCheck_IsParameterNarrow_Bun = False
    End If
    
    If fncChkHankakuEisuu(arryCSVRecord(enmBunColNo.Bunrui3)) = False Then      '��3���ރR�[�h
        FncSetError (intNowLine, enmBunColNo.Bunrui3, PARAM_HANEISUU
        FncCheck_IsParameterNarrow_Bun = False
    End If
    
End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_Limit_Bun
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�t�B�[���h�������������z���Ă��Ȃ����`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_Limit_Bun; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/08/31; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_Limit_Bun(ByRef arryCSVRecord() As String, intNowLine As Integer)

    FncCheck_Limit_Bun = True
    
    '(a)�ύX��
    If FncCheck_Length_Kotei(arryCSVRecord(enmBunColNo.HenkouDate), LENGTH_HENKOU_DATE) = False _
        Or FncCheckDate(arryCSVRecord(enmBunColNo.HenkouDate)) = False Then
        
        FncSetError (intNowLine, enmBunColNo.HenkouDate, PARAM_WRONGDATE)
        FncCheck_Limit_Bun = False
    End If
    '(b)JAN�R�[�h
    If FncCheck_JANCode(arryCSVRecord(enmBunColNo.janCD)) = False Then
        FncSetError (intNowLine, enmBunColNo.janCD, PARAM_SHORTDISIT)
        FncCheck_Limit_Bun = False
    End If
    '(c)��1���ރR�[�h
    If FncCheck_Length_Less(arryCSVRecord(enmBunColNo.Bunrui1), LENGTH_BUNRUI) = False Then
        FncSetError (intNowLine, enmBunColNo.Bunrui1, PARAM_SHORTDISIT)
        FncCheck_Limit_Bun = False
    End If
    '(d)��2���ރR�[�h
    If FncCheck_Length_Less(arryCSVRecord(enmBunColNo.Bunrui2), LENGTH_BUNRUI) = False Then
        FncSetError (intNowLine, enmBunColNo.Bunrui2, PARAM_SHORTDISIT)
        FncCheck_Limit_Bun = False
    End If
    '(e)��3���ރR�[�h
    If FncCheck_Length_Less(arryCSVRecord(enmBunColNo.Bunrui3), LENGTH_BUNRUI) = False Then
        FncSetError (intNowLine, enmBunColNo.Bunrui3, PARAM_SHORTDISIT)
        FncCheck_Limit_Bun = False
    End If
    
End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckParameter_DataBase_Bun
'*�y�����T�v�z
'*?    CSV�t�@�C���̃p�����[�^�ɑ΂��āADB�ɃA�N�Z�X���ă`�F�b�N���s���܂��B
'*?    (���ޏ���)
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVRecord() ; String CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckParameter_DataBase_Bun; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/08/31; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheckParameter_DataBase_Bun(ByRef arryCSVRecord() As String, intNowLine As Integer)

    Dim blnIsOKJANCode
    Dim blnIsOKNaibuCode
    Dim blnIsOKBunruiCode
    
    Dim strKeyCode As String
    
    FncCheckParameter_DataBase_Bun = "False"
    
    'a.JAN�R�[�h���L�����m�F����B
    blnIsOKJANCode = FncCheckDB_JANCode(arryCSVRecord(enmBunColNo.janCD), strKeyCode, intNowLine, RET_FLD_BUN)
    
    'b.KEY�R�[�h���L�����m�F����B
    blnIsOKNaibuCode = FncCheckDB_NaibuCode(strKeyCode, intNowLine, RET_FLD_BUN)
    
    'c.��1���ރR�[�h�A��2���ރR�[�h�A��3���ރR�[�h���L�����m�F����B
    blnIsOKBunruiCode = FncCheckDB_BunruiCode(arryCSVRecord(enmBunColNo.Bunrui1), arryCSVRecord(enmBunColNo.Bunrui2) _
                                            , arryCSVRecord(enmBunColNo.Bunrui3), intNowLine, RET_FLD_BUN)
    
    If blnIsOKJANCode And blnIsOKBunruiCode And blnIsOKNaibuCode Then
        FncCheckParameter_DataBase_Bun = "True"
    End If
       
End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckDB_NaibuCode
'*�y�����T�v�z
'*?    ���O���i�̓����R�[�h�d���`�F�b�N�����܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strBunrui1 ; String ��1���ރR�[�h������킷������
'*P    strBunrui2 ; String ��2���ރR�[�h������킷������
'*P    strBunrui3 ; String ��3���ރR�[�h������킷������
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*P    intFieldType; Integer; RET_FLD_BUN:bunrui���
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    strKeyCode ; String JAN�R�[�h�ɑΉ����铝���}�X�^�̃L�[�R�[�h���i�[����
'*R    FncCheckDB_NaibuCode; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/09/02; �����@����(IC); �V�K�쐬;
'*M    2006/09/12; �a�c(IC); ���i�I��v�e�⍇���o�ŉ��P;
'*< Comment End >**************************************************'
Private Function FncCheckDB_NaibuCode(ByVal strKeyCode As String _
                                , intNowLine As Integer, intFieldType As Integer) As Boolean
    
    Dim strKikakuNo As String   'KEY�R�[�h�����ɓo�^����Ă����懂
    
    'a.���ޏ�񂪗L�����m�F����B
    FncCheckDB_NaibuCode = True
    
'���ږ��̍쐬�i�G���[���b�Z�[�W�p�j
    Dim intField As Integer
    Select Case intFieldType
        Case RET_FLD_HACH:      intField = enmHacColNo.janCD
        Case RET_FLD_TENB:      intField = enmTenColNo.janCD
        Case RET_FLD_SYO:       intField = enmSyoColNo.janCD
        Case RET_FLD_BUN:       intField = enmBunColNo.janCD
    End Select

    '�@(a)WF�Ǘ��e�[�u���ɑ��݂��Ȃ����Ƃ��m�F����B
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
'*�y�v���V�[�W�����z
'*!    FncCheckDB_BunruiCode
'*�y�����T�v�z
'*?    ��1���ރR�[�h�A��2���ރR�[�h�A��3���ރR�[�h���L�������`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strBunrui1 ; String ��1���ރR�[�h������킷������
'*P    strBunrui2 ; String ��2���ރR�[�h������킷������
'*P    strBunrui3 ; String ��3���ރR�[�h������킷������
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*P    intFieldType; Integer; RET_FLD_BUN:bunrui���
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    strKeyCode ; String JAN�R�[�h�ɑΉ����铝���}�X�^�̃L�[�R�[�h���i�[����
'*R    FncCheckDB_BunruiCode; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/09/02; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheckDB_BunruiCode(ByVal strBunrui1 As String, ByVal strBunrui2 As String, ByVal strBunrui3 As String _
                              , intNowLine As Integer, intFieldType As Integer) As Boolean
    
    'a.���ޏ�񂪗L�����m�F����B
    FncCheckDB_BunruiCode = True
    
'���ږ��̍쐬�i�G���[���b�Z�[�W�p�j
    Dim intField As Integer
    Select Case intFieldType
        Case RET_FLD_BUN:       intField = enmBunColNo.Bunrui1
    End Select
    
    '�@(a)�i��R�t�}�X�^�ɑ��݂��鎖���m�F����B
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
'*�y�v���V�[�W�����z
'*!     fncChkKinsiMoji
'*
'*�y�����T�v�z
'*?     ���p�`�F�b�N(�p���̂�)
'*
'*�y�p�����[�^�z
'*      Name                   Type                Comment
'*P     strValue               String              �l
'*P     strErrMsg              String              �G���[���b�Z�[�W
'*�y�߂�l�z
'*      Name                   Type                Comment
'*R     fncChkHankaku          Boolean;
'*
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)            EDP-No      ����
'*M 2005/09/20; ���� (IC);             #-#####;    �V�K�쐬;
'*M 2006/06/15; �a�c(IC);               P-30109;   �v�e�d�l���P�ȈՈꊇ�@�\�̒ǉ�;
'*
'*< Comment End >**************************************************************
Private Function fncChkKinsiMoji(ByVal strValue As String) As Boolean
        
    Dim lngCount As Long
    Dim strCheck As String
    
    fncChkKinsiMoji = False
    
    Const PC_KINSOKU_STRING As String = "[,,,"",\,<,>,']"
    
    For lngCount = 1 To Len(strValue)
        strCheck = Mid(strValue, lngCount, 1)
        '1�o�C�g�����̏ꍇ�Ƀ`�F�b�N
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
'*�y�v���V�[�W�����z
'*!     fncChkKinsiMojiReceipt
'*
'*�y�����T�v�z
'*?     �֎~�����`�F�b�N
'*
'*�y�p�����[�^�z
'*      Name                   Type                Comment
'*P     strValue               String              �l
'*P     strErrMsg              String              �G���[���b�Z�[�W
'*�y�߂�l�z
'*      Name                   Type                Comment
'*R     fncChkHankaku          Boolean;
'*
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)            EDP-No      ����
'*M 2014/05/01; �a�c(LIS);              #-#####;    �V�K�쐬;
'*
'*< Comment End >**************************************************************
Private Function fncChkKinsiMojiReceipt(ByVal strValue As String) As Boolean
        
    Dim lngCount As Long
    Dim strCheck As String
    
    fncChkKinsiMojiReceipt = False
    
    Const PC_KINSOKU_STRING As String = "[�P�Q�R�S�T�U�V�W�X�Y�Z�[�\�]�^�_�`�p�q�r�s�t�u�v�w�x�y�z�{�|�}�~�����������������������������������������������������������������������������������������������������@�A�B�C�D�E�F�G�H�I�J�K�L�M�N�O�P�Q�R�S�T�U�V�W�X�Y�Z�[�\�]�_�`�a�b�c�d�e�f�g�h�i�j�k�l�m�n�o�p�q�r�s�t�u�~�����������������������������������߁燓����ہڇ����恿��]"
    
    For lngCount = 1 To Len(strValue)
        strCheck = Mid(strValue, lngCount, 1)
        If strCheck Like PC_KINSOKU_STRING Then
            Exit Function
        End If
    Next

    fncChkKinsiMojiReceipt = True

End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    subIns_SetAllTables_Bun
'*�y�����T�v�z
'*?    CSV�t�@�C���̃p�����[�^�ɑ΂��āADB�ɃA�N�Z�X���ă`�F�b�N���s���܂��B
'*?    (���ޏ���)
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    subIns_SetAllTables_Bun; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/09/05; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Sub subIns_SetAllTables_Bun()

    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    
    '�ϐ�
    Dim strMethodCallInfo As String
    Dim strWhereSQL As String               '�����쐬�p
    Dim strSQL  As String                   'RS�쐬�pSQL
    Dim lstrNow As Object                  '�������t�i�[�ϐ�
    Dim i As Long
    
    '���R�[�h�Z�b�g�p�ϐ�
    Dim iRecord As Long
    Dim rstHyoujiSiyou As ADODB.Recordset   '�\��/�d�l���
    Dim rstWfKanri As ADODB.Recordset       'WF�Ǘ����
    Dim rstRiekiKanri As ADODB.Recordset    '���v�Ǘ����
    Dim rstIppanKihon As ADODB.Recordset    '��ʊ�{���
    
    'WF�X�V�p�t���O(WF�X�V���̌������̈�String��"0"����"1"������)
    Dim strRiekiKanriAdded As String        '���v�Ǘ�RS���X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    Dim strHyoujiAdded As String            '�\��/�d�lRS���X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    
    '���t���n�߂Ɏ擾
    lstrNow = Format(Now(), "yyyy/mm/dd hh:nn:ss")
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun()"

    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    '---------------
    '   ��������
    '---------------
    On Error GoTo DefectErrorHandler

    ' ************** RS�쐬�pSQL�̏��������i����ރR�[�h�A�N�A�ԁA�}�Ԃ̎w��j **************
    strWhereSQL = "WHERE kikaku_syurui_cd = '" & pstrKikakuSyurui & "'" & vbCrLf
    strWhereSQL = strWhereSQL & "AND nen_ren_no ='" & pstrKikakuNenrenban & "'" & vbCrLf
    strWhereSQL = strWhereSQL & "AND eda_no ='" & pstrKikakuEdaban & "'" & vbCrLf
    
    ' ************** �\��/�d�l���RS�쐬 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(�\��/�d�l���RS�쐬)"
    
    strSQL = "SELECT * FROM t_syousen_hyouji_siyou with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "

        rstHyoujiSiyou = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstHyoujiSiyou.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    ' ************** WF�Ǘ����RS�쐬 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(WF�Ǘ����RS�쐬)"
    strSQL = "SELECT * FROM t_syousen_wf_kanri with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstWfKanri = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstWfKanri.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
         
    ' ************** ���v�Ǘ����RS�쐬 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(WF�Ǘ����RS�쐬)"
    strSQL = "SELECT * FROM t_syousen_rieki_kanri WITH(UPDLOCK)" & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstRiekiKanri = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstRiekiKanri.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    ' ************** VMST��ʊ�{���RS�쐬 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(VMST��ʊ�{���RS�쐬)"
    strSQL = "SELECT vk1.jan_cd, vik.* FROM v_ippan_kihon as vik with (READCOMMITTED) " & vbCrLf
    strSQL = strSQL & " INNER JOIN v_key_1 as vk1 ON (vik.key_cd = vk1.key_cd) " & vbCrLf
    strSQL = strSQL & " WHERE vk1.jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstIppanKihon = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstIppanKihon.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    ' ************** 1���R�[�h����RS���X�V *******************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(1���R�[�h����RS���X�V �J�n)"
    
    For iRecord = LBound(paryCsvData) To UBound(paryCsvData)
    
        'WF�X�V�p�t���O�̏�����
        strRiekiKanriAdded = 0
        strHyoujiAdded = 0
        
        '�f�[�^�X�V�ΏۓX�R�[�h�̎擾
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(��ʊ�{���̎擾)"
        
        'VMST��ʊ�{���̎擾
        With rstIppanKihon
            If Not .BOF Then
                .MoveFirst
                .Find ("jan_cd = '" & paryCsvData(iRecord)(enmBunColNo.janCD) & "'")
            End If
        End With
        
        '�f�[�^�X�V�ΏۓX�R�[�h�̎擾
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(�\��/�d�l���̎擾)"
        
        '�\��/�d�l���
        With rstHyoujiSiyou
            '�\��/�d�l���̎擾
            If Not .BOF Then
                .MoveFirst
                .Find ("jan_cd = '" & paryCsvData(iRecord)(enmBunColNo.janCD) & "'")
            End If
            
            '�f�[�^�X�V�ΏۓX�R�[�h�̎擾
            'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(�\��/�d�l���̍X�V)"
        
            If .EOF Then
                '�\��/�d�l���̒ǉ�
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
        
        '�f�[�^�X�V�ΏۓX�R�[�h�̎擾
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(���v�Ǘ����̎擾)"
        
        '���v�Ǘ����
        '���v�Ǘ��e�[�u���Ɉ�����JAN�R�[�h�̃f�[�^�����݂����ꍇ�͍폜����B
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
        
        '�f�[�^�X�V�ΏۓX�R�[�h�̎擾
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(�i��\�����̎擾)"
        
        '�i��\���e�[�u���Ɉ�����KEY���ރR�[�h�P�AKEY���ރR�[�h�Q�AKEY���ރR�[�h�R���
        '�R�t�������擾����B
        With pChkBunruiCodeRS
            If Not .BOF Then
                .MoveFirst
                .Filter = "key_bnrui_cd_1 = '" & paryCsvData(iRecord)(enmBunColNo.Bunrui1) & "' " _
                        & "AND key_bnrui_cd_2 = '" & paryCsvData(iRecord)(enmBunColNo.Bunrui2) & "'" _
                        & "AND key_bnrui_cd_3 = '" & paryCsvData(iRecord)(enmBunColNo.Bunrui3) & "'"
            End If
        End With
        
        '�f�[�^�X�V�ΏۓX�R�[�h�̎擾
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(���v�Ǘ����̍X�V)"
        
        '���v�Ǘ��e�[�u���Ƀf�[�^��ǉ�����B
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
        
        '�i��\���e�[�u����Filter����������B
        pChkBunruiCodeRS.Filter = adFilterNone


        '�f�[�^�X�V�ΏۓX�R�[�h�̎擾
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(WF�Ǘ����̍X�V)"
        
        'WF�Ǘ����ڂ��ŐV���ɍX�V����
        With rstWfKanri
            If Not .BOF Then
                .MoveFirst
                .Find ("jan_cd = '" & paryCsvData(iRecord)(enmBunColNo.janCD) & "'")
            End If
            'WF�Ǘ����ڂɖ����ꍇ�́A�V�K�ǉ�
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
    
            'WF�Ǘ����ڂɂ���ꍇ�́A�X�V
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
    
    'DB�̍X�V���s���B
    Call objCDBAccess.UpdateDCRecordset("VSWF", rstHyoujiSiyou)
    Call objCDBAccess.UpdateDCRecordset("VSWF", rstRiekiKanri)
    Call objCDBAccess.UpdateDCRecordset("VSWF", rstWfKanri)
    
    'RS�̃o�b�`�X�V
    'GetObjectContext.SetComplete
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Bun(�I�u�W�F�N�g�J��)"
    
    '�X�V�Ŏg�p�����I�u�W�F�N�g���J������B
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

    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
            'Err.Number, Err.Source, Err.Description
End Sub
'
'*< TOSTEM >*******************************************************************'
'*
'*�y�v���V�[�W�����z
'*!     FncSyoTouroku
'*
'*�y�����T�v�z
'*?     ���i��{���ꊇ�ύX�̃f�[�^�����i�I��WF��DB�Ɏ�荞��
'*
'*�y�p�����[�^�z
'*      Name                   Type                Comment
'*P     strErrMsg              Object             �G���[���b�Z�[�W;
'*�y�߂�l�z
'*      Name                   Type                Comment
'*R     FncSyoTouroku          Boolean;
'*
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)            EDP-No      ����
'*M 2005/09/05; ���� (IC);             P-24884;    �V�K�쐬;
'*M 2006/09/12; �a�c(IC);              P-XXXXX;    ���i�I��v�e�⍇���o�ŉ��P;
'*
'*< Comment End >**************************************************************
Public Function FncSyoTouroku(ByRef strErrMsg) As Boolean
    
    Dim intN As Integer
    Dim blnRet As Boolean
    Dim strJanCd As String
    Dim intNowLine As Integer
    Dim strCSVRecord() As String
    Dim flgChkError As Boolean      '�`�F�b�N���̃G���[�t���O�iTRUE:�G���[����j
    
    FncSyoTouroku = False
    
    Dim strMethodCallInfo As String
    On Error GoTo DefectErrorHandler
    
    'DB�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h�Z�b�g�̐���
    Dim strSQL          As String
    Dim objCDBAccess    As Object
        
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    '�ꊇ�o�^���A�������[�U�[�̂ݎ��s�ł���
    Call LockTableIkkatu(objCDBAccess)
    
    Call SetIkkatuLog("1-1", "���i��{���ꊇ�ύX�̃f�[�^�����i�I��WF��DB�Ɏ�荞��", "FncSyoTouroku", "", "", pstrTantousyaId, "INS")
    Call SetIkkatuLog("1-1-1", "�d��JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")
    '�d��JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(�d��JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����)"
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
    pChkSameJANCodeRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    pChkSameJANCodeRS.Fields("key_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U

        Call SetIkkatuLog("1-1-1", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-1-2", "CSV�t�@�C����Jan�R�[�h�̃��X�g(Jan�R�[�h�̏d������)�쐬", "SQL", "", "", pstrTantousyaId, "INS")
    'CSV�t�@�C����Jan�R�[�h�̃��X�g(Jan�R�[�h�̏d������)�쐬
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(CSV�t�@�C����Jan�R�[�h�̃��X�g(Jan�R�[�h�̏d������)�쐬)"

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
        Call SetIkkatuLog("1-1-3", "�����}�X�^�o�^�ς�JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")
    '�����}�X�^�o�^�ς�JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(�����}�X�^�o�^�ς�JAN�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����)"

    strSQL = "SELECT jan_cd, key_cd FROM v_key_1 WITH(readcommitted)"
    strSQL = strSQL & " WHERE jan_cd IN(" & pCsvJanCodeList & ") "
    
        pChkvk1JANCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    pChkvk1JANCodeRS.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
        Call SetIkkatuLog("1-1-3", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-1-4", "CSV�t�@�C����JAN�R�[�h�R�t����L�[�̃��X�g(��L�[�d������)�쐬", "SQL", "", "", pstrTantousyaId, "INS")
    '�����}�X�^�o�^�ς�JAN�R�[�h�R�t����L�[�̃��X�g(��L�[�d������)�쐬
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(CSV�t�@�C����JAN�R�[�h�R�t����L�[�̃��X�g(��L�[�d������)�쐬)"
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
    'CSV�t�@�C�������v�R�[�h�̃��X�g(���v�R�[�h�̏d������)�쐬
'    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(CSV�t�@�C�������v�R�[�h�̃��X�g(���v�R�[�h�̏d������)�쐬)"
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
        Call SetIkkatuLog("1-1-5", "���ރ}�X�^KEY�`�F�b�N�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")
    '����̌n�f�[�^�o�^�ςݎ��v�R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(���ރ}�X�^KEY�`�F�b�N�p�ؒf���R�[�h�Z�b�g����)"

    strSQL = "SELECT jigyou_cd, jyuyou_cd"
    strSQL = strSQL & " FROM v_uriba_taikei_hyouji WITH(readcommitted)"
    'strSQL = strSQL & " WHERE jyuyou_cd IN(" & pCsvJyuyouCodeList & ") "
    
        pChkJyuyouCodeRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    With pChkJyuyouCodeRS
            .Fields("jigyou_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
            .Fields("jyuyou_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    End With
        Call SetIkkatuLog("1-1-5", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-1-6", "�����p�^�[���R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����", "SQL", "", "", pstrTantousyaId, "INS")
    '�����p�^�[���R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(�����p�^�[���R�[�h�`�F�b�N�p�ؒf���R�[�h�Z�b�g����)"

    strSQL = "SELECT atukai_ptn_cd "
    strSQL = strSQL & " FROM v_atukai_ptn WITH(READCOMMITTED)"
    
        pChkAtukaiPtnCdRS = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    With pChkAtukaiPtnCdRS
            .Fields("atukai_ptn_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    End With
        Call SetIkkatuLog("1-1-6", "", "", "", "OK", "", "UPD")
        
        Call SetIkkatuLog("1-1-7", "�S�Ă̔z��f�[�^�̃`�F�b�N���s���B", "SQL", "", "", pstrTantousyaId, "INS")
    '�S�Ă̔z��f�[�^�̃`�F�b�N���s���B
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(�S�Ă̔z��f�[�^�̃`�F�b�N���s���B)"
    
    flgChkError = False
    intNowLine = 0
    strErrMsg = ""
    Do Until UBound(paryCsvData) < intNowLine
        
        strCSVRecord = paryCsvData(intNowLine)
        
        '(1)DB�Ɉˑ����Ȃ��`�F�b�N�FFncCheckParameter_CorrectForm_Syo
        '(2)DB���Q�Ƃ��ă`�F�b�N:FncCheckParameter_DataBase_Syo
        If FncCheckParameter_CorrectForm_Syo(strCSVRecord, intNowLine + 1) = False Or _
            FncCheckParameter_DataBase_Syo(strCSVRecord, intNowLine + 1) = False Then
            flgChkError = True
        End If
        
        intNowLine = intNowLine + 1
    Loop
    Call SetIkkatuLog("1-1-7", "", "", "", "OK", "", "UPD")
    
    Call SetIkkatuLog("1-1-8", "DB�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h�Z�b�g�̊J��", "SQL", "", "", pstrTantousyaId, "INS")
    'DB�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h�Z�b�g�̊J��
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(DB�Q�ƃG���[�`�F�b�N�p�ؒf���R�[�h�Z�b�g�̊J��)"
    
    '�`�F�b�N�Ŏg�p�����I�u�W�F�N�g���J������B
    If Not pChkSameJANCodeRS Is Nothing Then
        pChkSameJANCodeRS.Close
            pChkSameJANCodeRS = Nothing
    End If
        
    '�f�[�^�ɂP�ł���肪����ꍇ�́A�����ŏI���I�I�I
    If flgChkError = True Then
        '�G���[������΃G���[���b�Z�[�W���擾����B
        strErrMsg = FncGetErrorMsg_Syo()
        
        Exit Function
    End If
    Call SetIkkatuLog("1-1-8", "", "", "", "OK", "", "UPD")
    
    Call SetIkkatuLog("1-2", "�f�[�^�X�V����", "subIns_SetAllTables_Syo", "", "", pstrTantousyaId, "INS")
    
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncSyoTouroku(�f�[�^�X�V�����J�n)"
    
    '�S�Ẵe�[�u�����X�V����B
    Call subIns_SetAllTables_Syo
    Call SetIkkatuLog("1-2", "", "", "", "OK", "", "UPD")
    FncSyoTouroku = True
    
    '�I�u�W�F�N�g�̊J��
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
    
    '�I�u�W�F�N�g�̊J��
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
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", 'Err.Number, Err.Source, Err.Description

End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckParameter_CorrectForm_Syo
'*�y�����T�v�z
'*?    CSV�t�@�C���̃p�����[�^�ɑ΂��āADB�Ɉˑ����Ȃ��`���`�F�b�N���s���܂��B
'*?    �i���i��{���ꊇ�ύX�p�`�F�b�N�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckParameter_CorrectForm_Syo; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/09/05; �����@����(IC); �V�K�쐬;
'*M    2006/06/15; �a�c(IC);       P-30109; �v�e�d�l���P�ȈՈꊇ�@�\�̒ǉ�;
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
    
    'a.�K�{���ځiJAN�R�[�h�j�������Ă��邩�m�F����B
    blnIsOKRequired = FncCheck_RequiredParameter_Syo(arryCSVRecord(), intNowLine)
    
    'b.JAN�R�[�h�A�����p�^�[���A�W�������A�d�،����A�ʏ퉿�i�AҰ����]�������i�ɂ����āA���p�����ł��邩�m�F����B
    blnIsOKNumeric = FncCheck_IsParameterNumeric_Syo(arryCSVRecord(), intNowLine)
    
    'c1.�J�i���i��(TA�`�[�p)�A���V�[�g��(��)�A
    '  VH���v�R�[�h�AR&D���v�R�[�h�AVCS���v�R�[�h�ASVH���v�R�[�h�ɂ����āA���p�����ł��邩�m�F����B
    blnIsOKNarrow = FncCheck_IsParameterNarrow_Syo(arryCSVRecord(), intNowLine)
    
    'c2.�J�i���i��(TA�`�[�p)�A���V�[�g��(��)�APOP���C�A�E�g�P�APOP���C�A�E�g�Q�APOP�T�C�Y�ɂ����āA���p�֎~�����ł��邩�m�F����B
    blnIsOKKinsi = FncCheck_IsParameterKinsi_Syo(arryCSVRecord(), intNowLine)
    
    'd.�������i��(SA/��ײ�����)�A���V�[�g��(����)�A���i��(�̑��p)�A���[�J�[���A�K�i/�T�C�Y�A�i��/�^�ԁA�F/���A
    '�e�[�X�g����Y�n�����������i�R�s�[�ɂ����đS�p�����ł��邩�m�F����
    blnIsOKWide = FncCheck_IsParameterWide_Syo(arryCSVRecord(), intNowLine)

    'e.�����������𒴂��Ă��Ȃ����m�F����B
    blnIsOKLimit = FncCheck_Limit_Syo(arryCSVRecord(), intNowLine)

    'f.JAN�R�[�h�ɐ������`�F�b�N�f�W�b�g�̒l���Z�b�g����Ă��邩�𒲂ׂ�
    blnIsOKCheckDigit = FncCheck_CheckDigit(arryCSVRecord(enmSyoColNo.janCD), intNowLine, 4)
    
    If blnIsOKRequired And blnIsOKNumeric And blnIsOKNarrow And blnIsOKKinsi And _
       blnIsOKWide And blnIsOKLimit And blnIsOKCheckDigit Then
       
       FncCheckParameter_CorrectForm_Syo = True
    End If
    
End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_RequiredParameter_Syo
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�K�{���ڂ��󔒂łȂ����ǂ����`�F�b�N���܂��B
'*?    �i���i��{���ꊇ�ύX�p�`�F�b�N�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_RequiredParameter_Syo; Boolean;�`�F�b�NOK:True,NG:False
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/09/05; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'

Private Function FncCheck_RequiredParameter_Syo(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_RequiredParameter_Syo = True
    
    If arryCSVRecord(enmSyoColNo.janCD) = "" Then        'JAN�R�[�h
        FncSetError (intNowLine, enmSyoColNo.janCD, PARAM_EMPTY
        FncCheck_RequiredParameter_Syo = False
    End If

End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_IsParameterNumeric_Syo
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�S�t�B�[���h�����p���l���ǂ����`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_IsParameterNumeric_Syo; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/09/05; �����@����(IC); �V�K�쐬;
'*M    2006/06/15; �a�c(IC);       P-30109 �v�e�d�l���P�ȈՈꊇ�@�\�̒ǉ�;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterNumeric_Syo(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_IsParameterNumeric_Syo = True
    
    If FncChkHankakuNumber(arryCSVRecord(enmSyoColNo.janCD)) = False Then     'JAN�R�[�h
        FncSetError (intNowLine, enmSyoColNo.janCD, PARAM_NOTDISIT)
        FncCheck_IsParameterNumeric_Syo = False
    End If
    
    If arryCSVRecord(enmSyoColNo.TuujyouKakaku) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmSyoColNo.TuujyouKakaku)) = False Then     '�ʏ퉿�i
            FncSetError (intNowLine, enmSyoColNo.TuujyouKakaku, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Syo = False
        End If
    End If
    
    If arryCSVRecord(enmSyoColNo.MakerKakaku) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmSyoColNo.MakerKakaku)) = False Then     'Ұ����]�������i
            FncSetError (intNowLine, enmSyoColNo.MakerKakaku, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Syo = False
        End If
    End If
    
    If arryCSVRecord(enmSyoColNo.AtukaiPtn) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmSyoColNo.AtukaiPtn)) = False Then     '�����p�^�[��
            FncSetError (intNowLine, enmSyoColNo.AtukaiPtn, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Syo = False
        End If
    End If
    
    If arryCSVRecord(enmSyoColNo.SikiriGenka) <> "" Then
        If IsNumeric(arryCSVRecord(enmSyoColNo.SikiriGenka)) = False Then   '�d�،���
            FncSetError (intNowLine, enmSyoColNo.SikiriGenka, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Syo = False
        End If
    End If
    
    If arryCSVRecord(enmSyoColNo.HyoujyunBaika) <> "" Then
        If FncChkHankakuNumber(arryCSVRecord(enmSyoColNo.HyoujyunBaika)) = False Then   '�W������
            FncSetError (intNowLine, enmSyoColNo.HyoujyunBaika, PARAM_NOTDISIT)
            FncCheck_IsParameterNumeric_Syo = False
        End If
    End If
    
End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_IsParameterNarrow_Syo
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�Ώۃt�B�[���h�����p�������ǂ����`�F�b�N���܂��B
'*?    �i���i��{���ꊇ�ύX�p�`�F�b�N�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_IsParameterNarrow_Syo; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/09/05; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterNarrow_Syo(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_IsParameterNarrow_Syo = True

    If FncChkHankakuNarrow(arryCSVRecord(enmSyoColNo.SyouhinmeiKana)) = False Then      '�J�i���i��(TA�`�[�p)
        FncSetError (intNowLine, enmSyoColNo.SyouhinmeiKana, PARAM_NONARROW
        FncCheck_IsParameterNarrow_Syo = False
    End If
    
    If FncChkHankakuNarrow(arryCSVRecord(enmSyoColNo.ReceiptmeiKana)) = False Then      '���V�[�g��(��)
        FncSetError (intNowLine, enmSyoColNo.ReceiptmeiKana, PARAM_NONARROW
        FncCheck_IsParameterNarrow_Syo = False
    End If
    
    If fncChkHankakuEisuu(arryCSVRecord(enmSyoColNo.VhJyuyouCd)) = False Then      'VH���v�R�[�h
        FncSetError (intNowLine, enmSyoColNo.VhJyuyouCd, PARAM_HANEISUU
        FncCheck_IsParameterNarrow_Syo = False
    End If
    
    If fncChkHankakuEisuu(arryCSVRecord(enmSyoColNo.RdJyuyouCd)) = False Then      'R&D���v�R�[�h
        FncSetError (intNowLine, enmSyoColNo.RdJyuyouCd, PARAM_HANEISUU
        FncCheck_IsParameterNarrow_Syo = False
    End If
    
    If fncChkHankakuEisuu(arryCSVRecord(enmSyoColNo.VcsJyuyouCd)) = False Then      'VCS���v�R�[�h
        FncSetError (intNowLine, enmSyoColNo.VcsJyuyouCd, PARAM_HANEISUU
        FncCheck_IsParameterNarrow_Syo = False
    End If
    
    If fncChkHankakuEisuu(arryCSVRecord(enmSyoColNo.SvhJyuyouCd)) = False Then      'SVH���v�R�[�h
        FncSetError (intNowLine, enmSyoColNo.SvhJyuyouCd, PARAM_HANEISUU
        FncCheck_IsParameterNarrow_Syo = False
    End If
    
End Function
'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_IsParameterKinsi_Syo
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�Ώۃt�B�[���h�ɔ��p�֎~�������܂܂�Ă��邩�ǂ����`�F�b�N���܂��B
'*?    �i���i��{���ꊇ�ύX�p�`�F�b�N�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_IsParameterKinsi_Syo; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/09/20; �����@����(IC); �V�K�쐬;
'*M    2006/06/15; �a�c(IC);       P-30109 �v�e�d�l���P�ȈՈꊇ�@�\�̒ǉ�;
'*M    2014/05/02; �a�c(LIS);      P-xxxxx ���V�[�g��(����)�֎~�����`�F�b�N�ǉ�;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterKinsi_Syo(ByRef arryCSVRecord() As String, intNowLine As Integer) As Boolean

    FncCheck_IsParameterKinsi_Syo = True

    If fncChkKinsiMoji(arryCSVRecord(enmSyoColNo.SyouhinmeiKana)) = False Then      '�J�i���i��(TA�`�[�p)
        FncSetError (intNowLine, enmSyoColNo.SyouhinmeiKana, PARAM_KINSIMOJI
        FncCheck_IsParameterKinsi_Syo = False
    End If
    
    '2014/05/02 add wadak7
    If fncChkKinsiMojiReceipt(arryCSVRecord(enmSyoColNo.ReceiptmeiKanji)) = False Then      '���V�[�g��(����)
        FncSetError (intNowLine, enmSyoColNo.ReceiptmeiKanji, PARAM_KINSIMOJI_RECEIPT
        FncCheck_IsParameterKinsi_Syo = False
    End If
    
    If fncChkKinsiMoji(arryCSVRecord(enmSyoColNo.ReceiptmeiKana)) = False Then      '���V�[�g��(��)
        FncSetError (intNowLine, enmSyoColNo.ReceiptmeiKana, PARAM_KINSIMOJI
        FncCheck_IsParameterKinsi_Syo = False
    End If
    
    If fncChkKinsiMoji(arryCSVRecord(enmSyoColNo.POPLayout1)) = False Then          'POP���C�A�E�g�P
        FncSetError (intNowLine, enmSyoColNo.POPLayout1, PARAM_KINSIMOJI
        FncCheck_IsParameterKinsi_Syo = False
    End If
    
    If fncChkKinsiMoji(arryCSVRecord(enmSyoColNo.POPLayout2)) = False Then          'POP���C�A�E�g�Q
        FncSetError (intNowLine, enmSyoColNo.POPLayout2, PARAM_KINSIMOJI
        FncCheck_IsParameterKinsi_Syo = False
    End If
    
    If fncChkKinsiMoji(arryCSVRecord(enmSyoColNo.PopSize)) = False Then             'POP�T�C�Y
        FncSetError (intNowLine, enmSyoColNo.PopSize, PARAM_KINSIMOJI
        FncCheck_IsParameterKinsi_Syo = False
    End If
    
End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheck_IsParameterWide_Syo
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�t�B�[���h�������������z���Ă��Ȃ����`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_IsParameterWide_Syo; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/09/05; �����@����(IC); �V�K�쐬;
'*M    2006/06/15; �a�c(IC);       P-30109 �v�e�d�l���P�ȈՈꊇ�@�\�̒ǉ�;
'*< Comment End >**************************************************'
Private Function FncCheck_IsParameterWide_Syo(ByRef arryCSVRecord() As String, intNowLine As Integer)

    FncCheck_IsParameterWide_Syo = True
    
    '(a)�������i��(TA/��ײ�����)
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.SyouhinmeiKanji)) = False Then
        FncSetError (intNowLine, enmSyoColNo.SyouhinmeiKanji, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    '(b)���V�[�g��(����)
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.ReceiptmeiKanji)) = False Then
        FncSetError (intNowLine, enmSyoColNo.ReceiptmeiKanji, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    '(c)���i��(�̑��p)
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.SyouhihinmeiHansoku)) = False Then
        FncSetError (intNowLine, enmSyoColNo.SyouhihinmeiHansoku, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
    '���[�J�[��
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.MakerMei)) = False Then
        FncSetError (intNowLine, enmSyoColNo.MakerMei, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
    '�K�i/�T�C�Y
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.KikakuSize)) = False Then
        FncSetError (intNowLine, enmSyoColNo.KikakuSize, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
    '�i��/�^��
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.HinbanKataban)) = False Then
        FncSetError (intNowLine, enmSyoColNo.HinbanKataban, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
    '�F/��
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.IroGara)) = False Then
        FncSetError (intNowLine, enmSyoColNo.IroGara, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
    '�e�[�X�g
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.Taste)) = False Then
        FncSetError (intNowLine, enmSyoColNo.Taste, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
    '���Y�n��
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.Seisantimei)) = False Then
        FncSetError (intNowLine, enmSyoColNo.Seisantimei, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
    '����
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.Tokutyou)) = False Then
        FncSetError (intNowLine, enmSyoColNo.Tokutyou, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
    '���i�R�s�[
    If FncChkHankakuWide(arryCSVRecord(enmSyoColNo.SyouhinCopy)) = False Then
        FncSetError (intNowLine, enmSyoColNo.SyouhinCopy, PARAM_ZENKAKU
        FncCheck_IsParameterWide_Syo = False
    End If
    
End Function


'*< TOSTEM >*******************************************************************'
'*
'*�y�v���V�[�W�����z
'*!     fncChkHankakuEisuu
'*
'*�y�����T�v�z
'*?     ���p�`�F�b�N(�p���̂�) �S�p�����̓G���[
'*
'*�y�p�����[�^�z
'*      Name                   Type                Comment
'*P     strValue               String              �l
'*�y�߂�l�z
'*      Name                   Type                Comment
'*R     fncChkHankaku          Boolean;
'*
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)            EDP-No      ����
'*M 2005/09/20; ���� (IC);             #-#####;    �V�K�쐬;
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
'*�y�v���V�[�W�����z
'*!    FncCheck_Limit_Syo
'*�y�����T�v�z
'*?    CSV�f�[�^�̂P�s�ɑ΂��A�t�B�[���h�������������z���Ă��Ȃ����`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheck_Limit_Syo; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/08/31; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheck_Limit_Syo(ByRef arryCSVRecord() As String, intNowLine As Integer)

    FncCheck_Limit_Syo = True
    
    '(a)JAN�R�[�h
    If FncCheck_JANCode(arryCSVRecord(enmSyoColNo.janCD)) = False Then
        FncSetError (intNowLine, enmSyoColNo.janCD, PARAM_SHORTDISIT)
        FncCheck_Limit_Syo = False
    End If
    '(b)�J�i���i��(TA�`�[�p)
    If arryCSVRecord(enmSyoColNo.SyouhinmeiKana) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.SyouhinmeiKana), LENGTH_SYOKANA) = False Then
            FncSetError (intNowLine, enmSyoColNo.SyouhinmeiKana, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(c)�������i��(SA/��ײ����ޗp)
    If arryCSVRecord(enmSyoColNo.SyouhinmeiKanji) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.SyouhinmeiKanji), LENGTH_SYOKANJI) = False Then
            FncSetError (intNowLine, enmSyoColNo.SyouhinmeiKanji, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(d)���V�[�g��(��)
    If arryCSVRecord(enmSyoColNo.ReceiptmeiKana) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.ReceiptmeiKana), LENGTH_RESKANA) = False Then
            FncSetError (intNowLine, enmSyoColNo.ReceiptmeiKana, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(e)���V�[�g��(����)
    If arryCSVRecord(enmSyoColNo.ReceiptmeiKanji) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.ReceiptmeiKanji), LENGTH_RESKANJI) = False Then
            FncSetError (intNowLine, enmSyoColNo.ReceiptmeiKanji, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(f)���i��(�̑��p)
    If arryCSVRecord(enmSyoColNo.SyouhihinmeiHansoku) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.SyouhihinmeiHansoku), LENGTH_SYOHANSOKU) = False Then
            FncSetError (intNowLine, enmSyoColNo.SyouhihinmeiHansoku, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
    
     '���[�J�[��
    If arryCSVRecord(enmSyoColNo.MakerMei) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.MakerMei), 10) = False Then
            FncSetError (intNowLine, enmSyoColNo.MakerMei, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     '�K�i/�T�C�Y
    If arryCSVRecord(enmSyoColNo.KikakuSize) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.KikakuSize), 20) = False Then
            FncSetError (intNowLine, enmSyoColNo.KikakuSize, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     '�i��/�^��
    If arryCSVRecord(enmSyoColNo.HinbanKataban) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.HinbanKataban), 20) = False Then
            FncSetError (intNowLine, enmSyoColNo.HinbanKataban, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     '�F/��
    If arryCSVRecord(enmSyoColNo.IroGara) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.IroGara), 20) = False Then
            FncSetError (intNowLine, enmSyoColNo.IroGara, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     '�e�[�X�g
    If arryCSVRecord(enmSyoColNo.Taste) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.Taste), 20) = False Then
            FncSetError (intNowLine, enmSyoColNo.Taste, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     '���Y�n��
    If arryCSVRecord(enmSyoColNo.Seisantimei) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.Seisantimei), 20) = False Then
            FncSetError (intNowLine, enmSyoColNo.Seisantimei, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     '����
    If arryCSVRecord(enmSyoColNo.Tokutyou) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.Tokutyou), 150) = False Then
            FncSetError (intNowLine, enmSyoColNo.Tokutyou, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     '���i�R�s�[
    If arryCSVRecord(enmSyoColNo.SyouhinCopy) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.SyouhinCopy), 50) = False Then
            FncSetError (intNowLine, enmSyoColNo.SyouhinCopy, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     '�ʏ퉿�i
    If arryCSVRecord(enmSyoColNo.TuujyouKakaku) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.TuujyouKakaku), 7) = False Then
            FncSetError (intNowLine, enmSyoColNo.TuujyouKakaku, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     '���[�J�[��]�������i
    If arryCSVRecord(enmSyoColNo.MakerKakaku) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.MakerKakaku), 7) = False Then
            FncSetError (intNowLine, enmSyoColNo.MakerKakaku, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     'POP���C�A�E�g�P
    If arryCSVRecord(enmSyoColNo.POPLayout1) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.POPLayout1), 30) = False Then
            FncSetError (intNowLine, enmSyoColNo.POPLayout1, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     'POP���C�A�E�g�Q
    If arryCSVRecord(enmSyoColNo.POPLayout2) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.POPLayout2), 30) = False Then
            FncSetError (intNowLine, enmSyoColNo.POPLayout2, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
     'POP�T�C�Y
    If arryCSVRecord(enmSyoColNo.PopSize) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.PopSize), 30) = False Then
            FncSetError (intNowLine, enmSyoColNo.PopSize, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
    
    '(g)�����p�^�[��
    If arryCSVRecord(enmSyoColNo.AtukaiPtn) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.AtukaiPtn), LENGTH_ATUPTN) = False Then
            FncSetError (intNowLine, enmSyoColNo.AtukaiPtn, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(h)�W������
    If arryCSVRecord(enmSyoColNo.HyoujyunBaika) <> "" Then
        If FncCheck_Length_Less(arryCSVRecord(enmSyoColNo.HyoujyunBaika), LENGTH_BAIKA) = False Then
            FncSetError (intNowLine, enmSyoColNo.HyoujyunBaika, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(i)�d�،���
    If arryCSVRecord(enmSyoColNo.SikiriGenka) <> "" Then
        If FncCheck_Genka(arryCSVRecord(enmSyoColNo.SikiriGenka)) = False Then
            FncSetError (intNowLine, enmSyoColNo.SikiriGenka, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(j)VH���v�R�[�h
    If arryCSVRecord(enmSyoColNo.VhJyuyouCd) <> "" Then
        If FncCheck_Length_Kotei(arryCSVRecord(enmSyoColNo.VhJyuyouCd), LENGTH_JYUYOU) = False Then
            FncSetError (intNowLine, enmSyoColNo.VhJyuyouCd, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(k)R&D���v�R�[�h
    If arryCSVRecord(enmSyoColNo.RdJyuyouCd) <> "" Then
        If FncCheck_Length_Kotei(arryCSVRecord(enmSyoColNo.RdJyuyouCd), LENGTH_JYUYOU) = False Then
            FncSetError (intNowLine, enmSyoColNo.RdJyuyouCd, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(l)VCS���v�R�[�h
    If arryCSVRecord(enmSyoColNo.VcsJyuyouCd) <> "" Then
        If FncCheck_Length_Kotei(arryCSVRecord(enmSyoColNo.VcsJyuyouCd), LENGTH_JYUYOU) = False Then
            FncSetError (intNowLine, enmSyoColNo.VcsJyuyouCd, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    '(m)SVH���v�R�[�h
    If arryCSVRecord(enmSyoColNo.SvhJyuyouCd) <> "" Then
        If FncCheck_Length_Kotei(arryCSVRecord(enmSyoColNo.SvhJyuyouCd), LENGTH_JYUYOU) = False Then
            FncSetError (intNowLine, enmSyoColNo.SvhJyuyouCd, PARAM_SHORTDISIT)
            FncCheck_Limit_Syo = False
        End If
    End If
    
End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckParameter_DataBase_Syo
'*�y�����T�v�z
'*?    CSV�t�@�C���̃p�����[�^�ɑ΂��āADB�ɃA�N�Z�X���ă`�F�b�N���s���܂��B
'*?    (���ޏ���)
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckParameter_DataBase_Syo; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/08/31; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheckParameter_DataBase_Syo(ByRef arryCSVRecord() As String, intNowLine As Integer)

    Dim blnIsOKJANCode
    Dim blnIsOKNaibuCode
    Dim blnIsOKJyuyouCode
    Dim blnIsOKAtukaiPtnCd
    
    Dim strKeyCode As String
    
    FncCheckParameter_DataBase_Syo = "False"
    
    'a.JAN�R�[�h���L�����m�F����B
    blnIsOKJANCode = FncCheckDB_JANCode(arryCSVRecord(enmSyoColNo.janCD), strKeyCode, intNowLine, RET_FLD_SYO)
    
    'b.KEY�R�[�h���L�����m�F����B
    blnIsOKNaibuCode = FncCheckDB_NaibuCode(strKeyCode, intNowLine, RET_FLD_SYO)
    
    'c.���v�R�[�h���L�����m�F����B
    blnIsOKJyuyouCode = FncCheckDB_JyuyouCode(arryCSVRecord(), intNowLine, RET_FLD_SYO)
    
    'd.�����p�^�[���R�[�h���L�����m�F����B
    blnIsOKAtukaiPtnCd = FncCheckDB_AtukaiPtnCd(arryCSVRecord(enmSyoColNo.AtukaiPtn), intNowLine, RET_FLD_SYO)
    
    If blnIsOKJANCode And blnIsOKJyuyouCode And blnIsOKNaibuCode And blnIsOKAtukaiPtnCd Then
        FncCheckParameter_DataBase_Syo = "True"
    End If
       
End Function

'******************************************************************'
'*�y�v���V�[�W�����z
'*!    subIns_SetAllTables_Syo
'*�y�����T�v�z
'*?    CSV�t�@�C���̃p�����[�^�ɑ΂��āADB�ɃA�N�Z�X���ă`�F�b�N���s���܂��B
'*?    (���ޏ���)
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckParameter_DataBase_Syo; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/09/05; �����@����(IC); �V�K�쐬;
'*M    2006/06/15; �a�c(IC);       P-30109; �v�e�d�l���P�ȈՈꊇ�@�\�̒ǉ�;
'*< Comment End >**************************************************'
Private Sub subIns_SetAllTables_Syo()

    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    
    '�ϐ�
    Dim strMethodCallInfo As String
    Dim strWhereSQL As String               '�����쐬�p
    Dim strSQL  As String                   'RS�쐬�pSQL
    Dim lstrNow As Object                  '�������t�i�[�ϐ�
    Dim i As Long
    
    '���R�[�h�Z�b�g�p�ϐ�
    Dim iRecord As Long
    Dim rstHyoujiSiyou As ADODB.Recordset   '�\��/�d�l���
    Dim rstWfKanri As ADODB.Recordset       'WF�Ǘ����
    Dim rstHanbai As ADODB.Recordset        '�̔����
    Dim rstHanki As ADODB.Recordset         '�����ύX���
    Dim rstUriba As ADODB.Recordset         '������
    Dim rstIppanKihon As ADODB.Recordset    'VMST��ʊ�{���
    Dim rstHankiHenkou As ADODB.Recordset   'VMST�����ύX���
    Dim rstPOP As ADODB.Recordset           'POP���
    Dim rstSyanaiKihon As ADODB.Recordset   'VMST�Г���{���
    
    'WF�X�V�p�t���O(WF�X�V���̌������̈�String��"0"����"1"������)
    Dim strHyoujiAdded As String            '�����ύXRS���X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    Dim strHanbaiAdded As String            '�d���ύXRS���X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    Dim strHankiAdded As String             '�Z���^�[�R�[�h�ύXRS���X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    Dim strUribaAdded As String             '�Z���^�[�R�[�h�ύXRS���X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    Dim strPOPAdded As String               'POP���RS���X�V(�V�K�ǉ����͍X�V)���ꂽ��("0"�F�X�V�����@"1"�F�X�V�݂�)
    
    '���t���n�߂Ɏ擾
    lstrNow = Format(Now(), "yyyy/mm/dd hh:nn:ss")
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo()"

    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    '---------------
    '   ��������
    '---------------
    On Error GoTo DefectErrorHandler

    ' ************** RS�쐬�pSQL�̏��������i����ރR�[�h�A�N�A�ԁA�}�Ԃ̎w��j **************
    strWhereSQL = "WHERE kikaku_syurui_cd = '" & pstrKikakuSyurui & "'" & vbCrLf
    strWhereSQL = strWhereSQL & "AND nen_ren_no ='" & pstrKikakuNenrenban & "'" & vbCrLf
    strWhereSQL = strWhereSQL & "AND eda_no ='" & pstrKikakuEdaban & "'" & vbCrLf
    
    ' ************** �\��/�d�l���RS�쐬 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(�\��/�d�l���RS�쐬)"
    
    strSQL = "SELECT * FROM t_syousen_hyouji_siyou with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "

        rstHyoujiSiyou = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstHyoujiSiyou.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    ' ************** WF�Ǘ����RS�쐬 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(WF�Ǘ����RS�쐬)"
    strSQL = "SELECT * FROM t_syousen_wf_kanri with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstWfKanri = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstWfKanri.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
         
    ' ************** �̔����RS�쐬 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(�̔����RS�쐬)"
    strSQL = "SELECT * FROM t_syousen_hanbai with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstHanbai = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstHanbai.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    ' ************** �������RS�쐬 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(�������RS�쐬)"
    strSQL = "SELECT * FROM t_syousen_hanki_henkou with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstHanki = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstHanki.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    ' ************** ������RS�쐬 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(������RS�쐬)"
    strSQL = "SELECT * FROM t_syousen_uriba with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstUriba = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstUriba.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
        
    ' ************** POP���RS�쐬 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(POP���RS�쐬)"
    strSQL = "SELECT * FROM t_syousen_pop with (UPDLOCK) " & strWhereSQL
    strSQL = strSQL & " AND jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstPOP = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstPOP.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
        
    ' ************** VMST��ʊ�{���RS�쐬 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(VMST��ʊ�{���RS�쐬)"
    strSQL = "SELECT vk1.jan_cd, vik.* FROM v_ippan_kihon as vik with (READCOMMITTED) " & vbCrLf
    strSQL = strSQL & " INNER JOIN v_key_1 as vk1 ON (vik.key_cd = vk1.key_cd) " & vbCrLf
    strSQL = strSQL & " WHERE vk1.jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstIppanKihon = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstIppanKihon.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    ' ************** VMST�����ύX���RS�쐬 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(VMST�����ύX���RS�쐬)"
    strSQL = "SELECT vk1.jan_cd, vhh.* FROM v_hanki_henkou as vhh with (READCOMMITTED) " & vbCrLf
    strSQL = strSQL & " INNER JOIN v_key_1 as vk1 with (READCOMMITTED) ON (vhh.key_cd = vk1.key_cd)" & vbCrLf
    strSQL = strSQL & " WHERE vk1.jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstHankiHenkou = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstHankiHenkou.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    ' ************** VMST�Г���{���RS�쐬 **************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(VMST�Г���{���RS�쐬)"
    strSQL = "SELECT vk1.jan_cd, vsk.* FROM v_syanai_kihon as vsk with (READCOMMITTED) " & vbCrLf
    strSQL = strSQL & " INNER JOIN v_key_1 as vk1 ON (vsk.key_cd = vk1.key_cd) " & vbCrLf
    strSQL = strSQL & " WHERE vk1.jan_cd IN(" & pCsvJanCodeList & ") "
    
        rstSyanaiKihon = objCDBAccess.GetDCRecordset("VSWF", strSQL)
    rstSyanaiKihon.Fields("jan_cd").Properties("Optimize") = True '�I�v�e�B�}�C�U
    
    
    ' ************** 1���R�[�h����RS���X�V *******************
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(1���R�[�h����RS���X�V �J�n)"
    
    For iRecord = LBound(paryCsvData) To UBound(paryCsvData)
    
        'WF�X�V�p�t���O�̏�����
        strHyoujiAdded = 0
        strHanbaiAdded = 0
        strHankiAdded = 0
        strUribaAdded = 0
        strPOPAdded = 0
        
        '�f�[�^�X�V�ΏۓX�R�[�h�̎擾
        'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(��ʊ�{���̎擾)"
        
        'VMST��ʊ�{���̎擾
        With rstIppanKihon
            If Not .BOF Then
                .MoveFirst
                .Find ("jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "'")
            End If
        End With
        
        '�\��/�d�l���
        With rstHyoujiSiyou
             '�\��/�d�l���̎擾
             If Not .BOF Then
                 .MoveFirst
                 .Filter = "jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "' "
                 If .RecordCount <> 0 Then
                     .Delete
                 End If
             End If
         
             '�\��/�d�l���̒ǉ�
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
             '�����[�J�[��
             If paryCsvData(iRecord)(enmSyoColNo.MakerMei) <> "" Then
                 .Fields("maker_mei_kanji").Value = paryCsvData(iRecord)(enmSyoColNo.MakerMei)
             Else
                 .Fields("maker_mei_kanji").Value = rstIppanKihon.Fields("maker_mei_kanji").Value
             End If
             '���K�i/�T�C�Y
             If paryCsvData(iRecord)(enmSyoColNo.KikakuSize) <> "" Then
                 .Fields("kikaku_size_kanji").Value = paryCsvData(iRecord)(enmSyoColNo.KikakuSize)
             Else
                 .Fields("kikaku_size_kanji").Value = rstIppanKihon.Fields("kikaku_size_kanji").Value
             End If
             '���i��/�^��
             If paryCsvData(iRecord)(enmSyoColNo.HinbanKataban) <> "" Then
                 .Fields("sina_no_kata_no_kanji").Value = paryCsvData(iRecord)(enmSyoColNo.HinbanKataban)
             Else
                 .Fields("sina_no_kata_no_kanji").Value = rstIppanKihon.Fields("sina_no_kata_no_kanji").Value
             End If
             '���F/��
             If paryCsvData(iRecord)(enmSyoColNo.IroGara) <> "" Then
                 .Fields("iro_gara_kanji").Value = paryCsvData(iRecord)(enmSyoColNo.IroGara)
             Else
                 .Fields("iro_gara_kanji").Value = rstIppanKihon.Fields("iro_gara_kanji").Value
             End If
             '���e�[�X�g
             If paryCsvData(iRecord)(enmSyoColNo.Taste) <> "" Then
                 .Fields("taste_kanji").Value = paryCsvData(iRecord)(enmSyoColNo.Taste)
             Else
                 .Fields("taste_kanji").Value = rstIppanKihon.Fields("taste_kanji").Value
             End If
             '�����Y�n��
             If paryCsvData(iRecord)(enmSyoColNo.Seisantimei) <> "" Then
                 .Fields("seisanti_mei_kanji").Value = paryCsvData(iRecord)(enmSyoColNo.Seisantimei)
             Else
                 .Fields("seisanti_mei_kanji").Value = rstIppanKihon.Fields("seisanti_mei_kanji").Value
             End If
             '������
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
        
        '�̔����
        With rstHanbai
            If paryCsvData(iRecord)(enmSyoColNo.AtukaiPtn) <> "" Then
                '�̔����̎擾
                If Not .BOF Then
                    .MoveFirst
                    .Filter = "jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "' "
                    If .RecordCount <> 0 Then
                        .Delete
                    End If
                End If
            
                '�̔����̒ǉ�
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
        
        
        'VMST�����ύX���̎擾
        With rstHankiHenkou
            If Not .BOF Then
                .MoveFirst
                .Filter = "jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "' "
            End If
        
        End With
        
        '�������
        With rstHanki
            If paryCsvData(iRecord)(enmSyoColNo.HyoujyunBaika) <> "" _
                Or paryCsvData(iRecord)(enmSyoColNo.SikiriGenka) <> "" Then
                '�������̎擾
                If Not .BOF Then
                    .MoveFirst
                    .Filter = "jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "' "
                    If .RecordCount <> 0 Then
                        .Delete
                    End If
                End If
                
                    '�������̒ǉ�
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
                        .Fields("touroku_kbn").Value = 1 'VMST�����ύX���ɑΏۃf�[�^�����݂����ꍇ�͂P��ݒ肷��B
                    Else
                        .Fields("touroku_kbn").Value = 2 'VMST�����ύX���ɑΏۃf�[�^�����݂��Ȃ��ꍇ�͂Q��ݒ肷��B
                    End If
                    .Fields("touroku_user_id").Value = pstrTantousyaId
                    .Fields("touroku_ymd_hms").Value = lstrNow
                    .Update
                    .Filter = adFilterNone
                    strHankiAdded = 1
            End If
        End With
        
        rstHankiHenkou.Filter = adFilterNone
        '����̌n���
        With rstUriba
            '�ϐ�
            Dim allJigyouCd As Object
            Dim allJuyouCd As Object
            '�����ݒ�
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
            
            '����̌n���̎擾
            If allJigyouCd <> "" Then
                allJigyouCd = Left(allJigyouCd, Len(allJigyouCd) - 1)
                
                '����̌n���̒ǉ�
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

        'VMST�Г���{���̎擾
        With rstSyanaiKihon
            If Not .BOF Then
                .MoveFirst
                .Find ("jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "'")
            End If
        End With
        
        'POP���
        If paryCsvData(iRecord)(enmSyoColNo.SyouhinCopy) <> "" Or _
            paryCsvData(iRecord)(enmSyoColNo.TuujyouKakaku) <> "" Or _
            paryCsvData(iRecord)(enmSyoColNo.MakerKakaku) <> "" Or _
            paryCsvData(iRecord)(enmSyoColNo.POPLayout1) <> "" Or _
            paryCsvData(iRecord)(enmSyoColNo.POPLayout2) <> "" Or _
            paryCsvData(iRecord)(enmSyoColNo.PopSize) <> "" Then
            With rstPOP
                 'POP���̎擾
                 If Not .BOF Then
                     .MoveFirst
                     .Filter = "jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "' "
                     If .RecordCount <> 0 Then
                         .Delete
                     End If
                 End If
             
                 'POP���̒ǉ�
                 .AddNew
                 .Fields("kikaku_syurui_cd").Value = pstrKikakuSyurui
                 .Fields("nen_ren_no").Value = pstrKikakuNenrenban
                 .Fields("eda_no").Value = pstrKikakuEdaban
                 .Fields("jan_cd").Value = paryCsvData(iRecord)(enmSyoColNo.janCD)
                 '�����i�R�s�[
                 If paryCsvData(iRecord)(enmSyoColNo.SyouhinCopy) <> "" Then
                     .Fields("syouhin_copy").Value = paryCsvData(iRecord)(enmSyoColNo.SyouhinCopy)
                 Else
                     .Fields("syouhin_copy").Value = rstSyanaiKihon.Fields("syouhin_copy").Value
                 End If
                 '���ʏ퉿�i
                 If paryCsvData(iRecord)(enmSyoColNo.TuujyouKakaku) <> "" Then
                     .Fields("tuujyou_kakaku").Value = paryCsvData(iRecord)(enmSyoColNo.TuujyouKakaku)
                 Else
                     .Fields("tuujyou_kakaku").Value = rstSyanaiKihon.Fields("tuujyou_kakaku").Value
                 End If
                 '�����[�J�[��]�������i
                 If paryCsvData(iRecord)(enmSyoColNo.MakerKakaku) <> "" Then
                     .Fields("maker_kiboukouri_kakaku").Value = paryCsvData(iRecord)(enmSyoColNo.MakerKakaku)
                 Else
                     .Fields("maker_kiboukouri_kakaku").Value = rstSyanaiKihon.Fields("maker_kiboukouri_kakaku").Value
                 End If
                 '��POP���C�A�E�g�P
                 If paryCsvData(iRecord)(enmSyoColNo.POPLayout1) <> "" Then
                     .Fields("pop_layout_cd_1").Value = paryCsvData(iRecord)(enmSyoColNo.POPLayout1)
                 Else
                     .Fields("pop_layout_cd_1").Value = rstSyanaiKihon.Fields("pop_layout_cd_1").Value
                 End If
                 '��POP���C�A�E�g�Q
                 If paryCsvData(iRecord)(enmSyoColNo.POPLayout2) <> "" Then
                     .Fields("pop_layout_cd_2").Value = paryCsvData(iRecord)(enmSyoColNo.POPLayout2)
                 Else
                     .Fields("pop_layout_cd_2").Value = rstSyanaiKihon.Fields("pop_layout_cd_2").Value
                 End If
                 '��POP�T�C�Y
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

        'WF�Ǘ����ڂ��ŐV���ɍX�V����
        With rstWfKanri
            If Not .BOF Then
                .MoveFirst
                .Find ("jan_cd = '" & paryCsvData(iRecord)(enmSyoColNo.janCD) & "'")
            End If
            'WF�Ǘ����ڂɖ����ꍇ�́A�V�K�ǉ�
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
                'POP���o�^�敪
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
    
            'WF�Ǘ����ڂɂ���ꍇ�́A�X�V
            Else
                .Fields("touroku_kbn").Value = WF_KANRI_TOUROKU_KBN_SYUUSEI
                .Fields("touroku_tantousya").Value = pstrTantousyaId
                .Fields("kikaku_no").Value = pstrKikakuSyurui & pstrKikakuNenrenban & "-" & pstrKikakuEdaban
                .Fields("naibu_cd").Value = rstIppanKihon.Fields("key_cd").Value
                If strHyoujiAdded = 0 Then
                Else
                    .Fields("hyouji_siyou_touroku_kbn").Value = 1
                End If
                'POP���o�^�敪
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
    
    'RS�̃o�b�`�X�V
    'GetObjectContext.SetComplete
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(�I�u�W�F�N�g�J��)"
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
    
    'strMethodCallInfo = App.Title & ".VswfIkatuTou.subIns_SetAllTables_Syo(�o�^�����p�����̂ł����ŊJ������)"
    
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

    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
            'Err.Number, Err.Source, Err.Description
End Sub



'*< UNISYSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckDB_JyuyouCode
'*�y�����T�v�z
'*?    ��1���ރR�[�h�A��2���ރR�[�h�A��3���ރR�[�h���L�������`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*P    intFieldType; Integer; RET_FLD_BUN:bunrui���
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    strKeyCode ; String JAN�R�[�h�ɑΉ����铝���}�X�^�̃L�[�R�[�h���i�[����
'*R    FncCheckDB_JyuyouCode; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2005/09/02; �����@����(IC); �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheckDB_JyuyouCode(ByRef arryCSVRecord() As String, intNowLine As Integer, intFieldType As Integer) As Boolean
    
    'a.���ޏ�񂪗L�����m�F����B
    FncCheckDB_JyuyouCode = True
    
    '�@(a)����̌n�}�X�^�ɓ���̎��v�R�[�h�����݂��鎖���m�F����B
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
'*�y�v���V�[�W�����z
'*!    FncCheckDB_AtukaiPtnCd
'*�y�����T�v�z
'*?    �����p�^�[���R�[�h���L�������`�F�b�N���܂��B
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intNowLine; Integer; CSV�t�@�C���̌��݂̍s
'*P    intFieldType; Integer; RET_FLD_SYO:���i���
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    FncCheckDB_AtukaiPtnCd; Boolean;�`�F�b�NOK:True,NG:False
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2006/09/22; �a�c(IC); ���i�I��v�e�⍇���o�ŉ��P;
'*< Comment End >**************************************************'
Private Function FncCheckDB_AtukaiPtnCd(ByVal strAtukaiPtnCd As String, _
                                      ByVal intNowLine As Integer, _
                                      ByVal intFieldType As Integer) As Boolean
    
    'a.�����p�^�[����񂪗L�����m�F����B
    FncCheckDB_AtukaiPtnCd = True
    
    '�@(a)�����p�^�[���}�X�^�ɑ��݂��鎖���m�F����B
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
'*�y�v���V�[�W�����z
'*!     GetSaibanNo
'*
'*�y�����T�v�z
'*?     �̔�NO�̎擾�B
'*
'*�y�p�����[�^�z
'*      Name                �@�@  Type                Comment
'*P     strSaibanKbn;             string              �̔ԋ敪
'*
'*�y�߂�l�z
'*      Name                �@�@  Type                Comment
'*R     GetSaibanNo;              integer;            �̔�NO;
'*
'******************************************************************************
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)       EDP-No      ����
'*M 2008/02/29; �a�c(IC);         P-32685;    ���i�I��v�e�V�X�e���̌����E�����ύX�o�^���̒l�����`�F�b�N�@�\�ǉ�;
'*< Comment End >**************************************************************
Public Function GetSaibanNo(ByVal strSaibanKbn As String) As Integer
                               
    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    Dim objCDBAccess    As COMDBUtil.CDBAccess
    Dim strSQL          As String
    Dim rsSaiban        As ADODB.Recordset
    
    On Error GoTo DefectErrorHandler

    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.GetSaibanNo(" & strSaibanKbn & _
                                                            ")"
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL��
    strSQL = ""
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     saiban_no + 1 AS saiban_no"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     w_syousen_neire_keikoku_saiban WITH(UPDLOCK)"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     saiban_kbn = ?"

    '�ؒfRecordset�擾
        rsSaiban = objCDBAccess.GetDCRecordset("VSWF", strSQL, MakeParam("@saiban_kbn", DT_Char, 1, strSaibanKbn))
    
    GetSaibanNo = rsSaiban.Fields("saiban_no").Value
    If GetSaibanNo >= 1000 Then
        GetSaibanNo = 0
    End If
    rsSaiban.Close
        rsSaiban = Nothing
    
     'SQL��
    strSQL = ""
    strSQL = strSQL & " UPDATE"
    strSQL = strSQL & "     w_syousen_neire_keikoku_saiban"
    strSQL = strSQL & " SET saiban_no = ?"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     saiban_kbn = ?"
   
    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no", DT_Int, 2, GetSaibanNo), _
                        MakeParam("@saiban_kbn", DT_Char, 1, strSaibanKbn) _
                        )
    
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    '�C���X�^���X�j��
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Function

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
             'Err.Number, Err.Source, Err.Description
End Function

'*< TOSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    NeirerituCheck
'*�y�����T�v�z
'*?    �l�����̃`�F�b�N
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    aryData;             CSV�t�@�C�����
'*P    strKikakuSyuruiCd;   ����ރR�[�h
'*P    strNenRenNo;         �N�A��
'*P    strEdaNo;            �}��
'*P    strLoginUser;        ���O�C�����[�U�[
'*P    intSaibanNo;         �̔�NO
'*P    sngSyouhizeiRitu;    ����ŗ�(%)
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    �Ȃ�
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2008/03/03; �a�c(IC); P-32685 ���i�I��WF�V�X�e���̌����E�����ύX�o�^���̒l�����`�F�b�N�@�\�ǉ�;
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
   
    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.NeirerituCheck(aryData" & _
                                                        "," & strKikakuSyuruiCD & _
                                                        "," & strNenRenNO & _
                                                        "," & strEdaNO & _
                                                        "," & strLoginUser & _
                                                        "," & intSaibanNo & _
                                                        "," & sngSyouhizeiRitu & _
                                                        ")"
    '�����[�N�e�[�u���̏�����
    Call DeleteWork(intSaibanNo)
    
    '��JAN�R�[�h�̃_�u��r���A�o�^
    Call DistinctJanCd(aryData, strKikakuSyuruiCD, strNenRenNO, strEdaNO, intSaibanNo)
   
    '���l�����x�����̏�����
    Call DeleteNeireKeikoku(intSaibanNo)
   
    '���l�����x�����̓o�^
    Call InsertNeireKeikoku(strLoginUser, intSaibanNo, sngSyouhizeiRitu)
    
    '�����[�N�e�[�u���̏�����
    'Call DeleteWork(intSaibanNo)
Exit Sub

DefectErrorHandler:
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    DeleteWork
'*�y�����T�v�z
'*?    ���[�N�e�[�u���̏�����
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intSaibanNo;         �̔�NO
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    �Ȃ�
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2008/02/27; �a�c(IC); P-32685 ���i�I��WF�V�X�e���̌����E�����ύX�o�^���̒l�����`�F�b�N�@�\�ǉ�;
'*< Comment End >**************************************************'
Private Sub DeleteWork(ByVal intSaibanNo As Integer)
    
    Dim strMethodCallInfo As String
    Dim strSQL As String
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    
    On Error GoTo DefectErrorHandler
   
    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.DeleteWork(" & intSaibanNo & _
                                                            ")"
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
       
    '�����i�I��WF(�l�����x��)WORK
    'SQL��
    strSQL = ""
    strSQL = strSQL & " DELETE FROM"
    strSQL = strSQL & "     w_syousen_neire_keikoku"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     saiban_no = ?"

    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, MakeParam("@saiban_no", DT_Int, 2, intSaibanNo))
  
    '�����i�I��WF(�l�����x���L�[)WORK
    'SQL��
    strSQL = ""
    strSQL = strSQL & " DELETE FROM"
    strSQL = strSQL & "     w_syousen_neire_keikoku_key"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     saiban_no = ?"

    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, MakeParam("@saiban_no", DT_Int, 2, intSaibanNo))
  
    '�����i�I��WF(�l�����x��JAN)WORK
    'SQL��
    strSQL = ""
    strSQL = strSQL & " DELETE FROM"
    strSQL = strSQL & "     w_syousen_neire_keikoku_jan"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     saiban_no = ?"

    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, MakeParam("@saiban_no", DT_Int, 2, intSaibanNo))
  
    '�C���X�^���X�j��
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    DistinctJanCd
'*�y�����T�v�z
'*?    CSV����JAN�R�[�h��DISTINCT�A�o�^
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    aryData;                  CSV�t�@�C�����
'*     strKikakuSyuruiCd;        ����ރR�[�h
'*     strNenRenNo;              �N�A��
'*     strEdaNo;                 �}��
'*P    intSaibanNo;              �̔�NO
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    �Ȃ�
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2008/02/27; �a�c(IC); P-32685 ���i�I��WF�V�X�e���̌����E�����ύX�o�^���̒l�����`�F�b�N�@�\�ǉ�;
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
   
    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.DistinctJanCd(aryData" & _
                                                            "," & strKikakuSyuruiCD & _
                                                            "," & strNenRenNO & _
                                                            "," & strEdaNO & _
                                                            "," & intSaibanNo & _
                                                            ")"
    For i = 0 To UBound(aryData)
        'JAN�R�[�h���݃`�F�b�N
            rsJan = GetJanInfo(strKikakuSyuruiCD _
                                , strNenRenNO _
                                , strEdaNO _
                                , aryData(i)(enmTenColNo.janCD) _
                                , aryData(i)(enmTenColNo.HenkouDate) _
                                , intSaibanNo)
        If rsJan.EOF Then
            '���݂��Ȃ��ꍇ�A�o�^
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
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >*******************************************************************
'*
'*�y�v���V�[�W�����z
'*!     GetJanInfo
'*
'*�y�����T�v�z
'*?     ���i�I��WF(�l�����x��JAN)WORK���R�[�h�Z�b�g�̎擾�B
'*
'*�y�p�����[�^�z
'*      Name                �@�@  Type                Comment
'*      strKikakuSyuruiCd;        String;             ����ރR�[�h;
'*      strNenRenNo;              String;             �N�A��;
'*      strEdaNo;                 String;             �}��;
'*      strJanCd;                 String;             JAN�R�[�h;
'*      strHenkouDate;            String;             �ύX��;
'*      intSaibanNo;              integer;            �̔�NO;
'*
'*�y�߂�l�z
'*      Name                �@�@  Type                Comment
'*R     GetBaikaHenkou;           ADODB.Recordset;    ���i�I��WF(�����ύX���)���R�[�h�Z�b�g;
'*
'******************************************************************************
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)       EDP-No      ����
'*M 2008/02/29; �a�c(IC);         P-32685;    ���i�I��v�e�V�X�e���̌����E�����ύX�o�^���̒l�����`�F�b�N�@�\�ǉ�;
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

    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.GetJanInfo(" & strKikakuSyuruiCD & _
                                                            "," & strNenRenNO & _
                                                            "," & strEdaNO & _
                                                            "," & strJanCd & _
                                                            "," & strHenkouDate & _
                                                            "," & intSaibanNo & _
                                                            ")"
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL��
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

    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@kikaku_syurui_cd", DT_Char, 3, strKikakuSyuruiCD), _
                        MakeParam("@nen_ren_no", DT_Char, 10, strNenRenNO), _
                        MakeParam("@eda_no", DT_Char, 3, strEdaNO), _
                        MakeParam("@jan_cd", DT_VarChar, 20, strJanCd), _
                        MakeParam("@henkou_date", DT_Char, 8, strHenkouDate), _
                        MakeParam("@saiban_no", DT_Int, 2, intSaibanNo) _
                        )
    
    '�ؒfRecordset�擾
        GetJanInfo = objCDBAccess.GetDCRecordset("VSWF", strSQL, vntAryParam)
    
    '�C���X�^���X�j��
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Function

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
             'Err.Number, Err.Source, Err.Description
End Function

'*< TOSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    InsertJan
'*�y�����T�v�z
'*?    ���i�I��WF(�l�����x��JAN)WORK�̓o�^
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strKikakuSyuruiCd;   ����ރR�[�h
'*P    strNenRenNo;         �N�A��
'*P    strEdaNo;            �}��
'*     strJanCd;            JAN�R�[�h
'*     strHenkouDate;       �ύX��
'*     intSaibanNo;         �̔�NO
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    �Ȃ�
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2008/02/27; �a�c(IC); P-32685 ���i�I��WF�V�X�e���̌����E�����ύX�o�^���̒l�����`�F�b�N�@�\�ǉ�;
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
   
    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.InsertJan(" & strKikakuSyuruiCD & _
                                                            "," & strNenRenNO & _
                                                            "," & strEdaNO & _
                                                            "," & strJanCd & _
                                                            "," & strHenkouDate & _
                                                            "," & intSaibanNo & _
                                                            ")"
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
       
    'SQL��
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

    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@kikaku_syurui_cd", DT_Char, 3, strKikakuSyuruiCD), _
                        MakeParam("@nen_ren_no", DT_Char, 10, strNenRenNO), _
                        MakeParam("@eda_no", DT_Char, 3, strEdaNO), _
                        MakeParam("@jan_cd", DT_VarChar, 20, strJanCd), _
                        MakeParam("@henkou_date", DT_Char, 8, strHenkouDate), _
                        MakeParam("@saiban_no", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
  
    '�C���X�^���X�j��
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    DeleteNeireKeikoku
'*�y�����T�v�z
'*?    �l�����x�����̏�����
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intSaibanNo;         �̔�NO
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    �Ȃ�
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2008/02/27; �a�c(IC); P-32685 ���i�I��WF�V�X�e���̌����E�����ύX�o�^���̒l�����`�F�b�N�@�\�ǉ�;
'*< Comment End >**************************************************'
Private Sub DeleteNeireKeikoku(ByVal intSaibanNo As Integer)
    
    Dim strMethodCallInfo As String
    Dim strSQL As String
    Dim vntAryParam As Object
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    
    On Error GoTo DefectErrorHandler
   
    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.DeleteNeireKeikoku(" & intSaibanNo & _
                                                                    ")"
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
       
    'SQL��
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

    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, MakeParam("@saiban_no", DT_Int, 2, intSaibanNo))
  
    '�C���X�^���X�j��
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    InsertNeireKeikoku
'*�y�����T�v�z
'*?    �l�����x���̓o�^
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strLoginUser;        ���O�C�����[�U�[
'*P    intSaibanNo;         �̔�NO
'*P    sngSyouhizeiRitu;    ����ŗ�(%)
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    �Ȃ�
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2008/02/27; �a�c(IC); P-32685 ���i�I��WF�V�X�e���̌����E�����ύX�o�^���̒l�����`�F�b�N�@�\�ǉ�;
'*< Comment End >**************************************************'
Private Sub InsertNeireKeikoku(ByVal strLoginUser As String _
                          , ByVal intSaibanNo As Integer _
                          , ByVal sngSyouhizeiRitu As Single)
    
    Dim strMethodCallInfo As String
    
    On Error GoTo DefectErrorHandler
   
    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.InsertNeireKeikoku(" & strLoginUser & _
                                                            "," & intSaibanNo & _
                                                            "," & sngSyouhizeiRitu & _
                                                            ")"
    '�������ύX���̃`�F�b�N
        '���߉ߋ�����
        '���i�I��WF(�l�����x��)WORK
    Call InsertBaikaWork(strLoginUser _
                    , intSaibanNo _
                    , sngSyouhizeiRitu _
                    )

        '���i�I��WF(�l�����x��KEY)WORK
    Call InsertBaikaKeyWork(intSaibanNo)
        
        '���i�I��WF(�l�����x��)
    Call InsertNeire(intSaibanNo, "1")
    
    '���d���ύX���̃`�F�b�N
        '���߉ߋ�����
        '���i�I��WF(�l�����x��)WORK
    Call InsertGenkaWork(strLoginUser _
                    , intSaibanNo _
                    , sngSyouhizeiRitu _
                    )

        '���i�I��WF(�l�����x��KEY)WORK
    Call InsertGenkaKeyWork(intSaibanNo)
        
        '���i�I��WF(�l�����x��)
    Call InsertNeire(intSaibanNo, "2")
    
    
    'Commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >*******************************************************************
'*
'*�y�v���V�[�W�����z
'*!     InsertBaikaWork
'*
'*�y�����T�v�z
'*?     ���i�I��WF(�l�����x��)WORK�ɒl�����ُ핪��o�^
'*
'*�y�p�����[�^�z
'*      Name                �@�@  Type                Comment
'*      strLoginUser;             String;             ���O�C�����[�U�[;
'*      intSaibanNo;              Integer;            �̔�NO;
'*      sngSyouhizeiRitu;         Single;             ����ŗ�(%);
'*
'*�y�߂�l�z
'*      Name                �@�@  Type                Comment
'*R     �Ȃ�
'*
'******************************************************************************
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)       EDP-No      ����
'*M 2008/02/29; �a�c(IC);         P-32685;    ���i�I��v�e�V�X�e���̌����E�����ύX�o�^���̒l�����`�F�b�N�@�\�ǉ�;
'*M 2009/04/22; Sureerat P.;      P-35471;    Mod check �l�������,�l��������;
'*M 2014/01/22; ������(��A) ;�@P-46854; �Ő������ɔ����Ή�(���i�I��WF�E�V���i���);
'*M 2014/03/04; �a�c(IC);         P-46854;    �����𕪊�(�^�C���A�E�g�Ή�);
'*M 2014/04/10; �a�c(IC);         P-46854;    KEY�o�^�ȈՉ�;
'*M 2014/04/18; �a�c(IC);         P-46854;    �d���ύX���(�^�C���A�E�g�Ή�);
'*< Comment End >**************************************************************
Private Sub InsertBaikaWork(ByVal strLoginUser As String _
                      , ByVal intSaibanNo As Integer _
                      , ByVal sngSyouhizeiRitu As Single)

    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim strSQL As String
    
    On Error GoTo DefectErrorHandler

    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.InsertBaikaWork(" & strLoginUser & _
                                                            "," & intSaibanNo & _
                                                            "," & sngSyouhizeiRitu & _
                                                             ")"
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    Dim syouhizeiRitu '����ŗ�
    Dim zeiHanteiKbn  '�Ŕ���敪
    Dim objVswf
    Dim objStr
    
    '���i���o�^COM�̃C���X�^���X����
        objVswf = CreateObject("VSWF01BUS005.CVswfSyouhizei")
    
    '����ŗ��ƐŔ���敪�f�[�^���擾����
    objStr = objVswf.GetKaisibiJyouhou()
    syouhizeiRitu = Split(objStr, "|")(0)
    zeiHanteiKbn = Split(objStr, "|")(1)
    
        objVswf = Nothing
    'SQL��
    '�����i�I��WF(�d���ύX���)�Ƃ̃`�F�b�N
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
   
    If zeiHanteiKbn = "0" Then  ' �ō��� �̏ꍇ
        '|--------------------(����(�Ŕ���)------------------------||--����---||--------------------(����(�Ŕ���)--------------|
        '((����(�ō���) �~ ����ŗ� �� (����ŗ� �{ 100)�| ����Ŋz) - ����)  /  (����(�ō���) �~ ����ŗ� �� (����ŗ� �{ 100)
        strSQL = strSQL & "         ,((sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) - sel.genka) / (sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) * 100 AS neire_ritu"     '�l����
    Else                        ' �Ŕ��� �̏ꍇ
        '(����(�Ŕ���) �| ����) �� ����(�Ŕ���)
        strSQL = strSQL & "         ,round((sel.baika- sel.genka) / sel.baika,3) * 100 AS neire_ritu "      '�l����
    End If
    
    strSQL = strSQL & "         ,'1' AS touroku_naiyou"         '�����ύX(�Œ�l)
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
                                    '���i�I��WF(�d���ύX���)
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

    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@saiban_no2", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)


    'SQL��
    '���d���ύX���Ƃ̃`�F�b�N
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
   
    If zeiHanteiKbn = "0" Then  ' �ō��� �̏ꍇ
        '|--------------------(����(�Ŕ���)------------------------||--����---||--------------------(����(�Ŕ���)--------------|
        '((����(�ō���) �~ ����ŗ� �� (����ŗ� �{ 100)�| ����Ŋz) - ����)  /  (����(�ō���) �~ ����ŗ� �� (����ŗ� �{ 100)
        strSQL = strSQL & "         ,((sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) - sel.genka) / (sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) * 100 AS neire_ritu"     '�l����
    Else                        ' �Ŕ��� �̏ꍇ
        '(����(�Ŕ���) �| ����) �� ����(�Ŕ���)
        strSQL = strSQL & "         ,round((sel.baika- sel.genka) / sel.baika,3) * 100 AS neire_ritu "      '�l����
    End If
    
    strSQL = strSQL & "         ,'1' AS touroku_naiyou"         '�����ύX(�Œ�l)
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
                                    '�d���ύX���
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

    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@saiban_no2", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)


    'SQL��
    '���X�ʏ��Ƃ̃`�F�b�N
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
   
    If zeiHanteiKbn = "0" Then  ' �ō��� �̏ꍇ
        '|--------------------(����(�Ŕ���)------------------------||--����---||--------------------(����(�Ŕ���)--------------|
        '((����(�ō���) �~ ����ŗ� �� (����ŗ� �{ 100)�| ����Ŋz) - ����)  /  (����(�ō���) �~ ����ŗ� �� (����ŗ� �{ 100)
        strSQL = strSQL & "         ,((sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) - sel.genka) / (sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) * 100 AS neire_ritu"     '�l����
    Else                        ' �Ŕ��� �̏ꍇ
        '(����(�Ŕ���) �| ����) �� ����(�Ŕ���)
        strSQL = strSQL & "         ,round((sel.baika- sel.genka) / sel.baika,3) * 100 AS neire_ritu "      '�l����
    End If
    
    strSQL = strSQL & "         ,'1' AS touroku_naiyou"         '�����ύX(�Œ�l)
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
    strSQL = strSQL & "             ,'19900101' AS genka_henkou_date"            'V�}�X�^�̕ύX����1990/01/01���Z�b�g
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
                                    '�X�ʏ��
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

    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@saiban_no2", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)


    'SQL��
    '���Г���{���Ƃ̃`�F�b�N
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
   
    If zeiHanteiKbn = "0" Then  ' �ō��� �̏ꍇ
        '|--------------------(����(�Ŕ���)------------------------||--����---||--------------------(����(�Ŕ���)--------------|
        '((����(�ō���) �~ ����ŗ� �� (����ŗ� �{ 100)�| ����Ŋz) - ����)  /  (����(�ō���) �~ ����ŗ� �� (����ŗ� �{ 100)
        strSQL = strSQL & "         ,((sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) - sel.genka) / (sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) * 100 AS neire_ritu"     '�l����
    Else                        ' �Ŕ��� �̏ꍇ
        '(����(�Ŕ���) �| ����) �� ����(�Ŕ���)
        strSQL = strSQL & "         ,round((sel.baika- sel.genka) / sel.baika,3) * 100 AS neire_ritu "      '�l����
    End If
    
    strSQL = strSQL & "         ,'1' AS touroku_naiyou"         '�����ύX(�Œ�l)
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
    strSQL = strSQL & "             ,'19900101' AS genka_henkou_date"            'V�}�X�^�̕ύX����1990/01/01���Z�b�g
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
                                    '�Г���{���
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

    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@saiban_no2", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
 
    '�C���X�^���X�j��
    objCDBAccess = Nothing
    
    'commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
             'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >*******************************************************************
'*
'*�y�v���V�[�W�����z
'*!     InsertBaikaKeyWork
'*
'*�y�����T�v�z
'*?     ���i�I��WF(�l�����x��KEY)WORK�ɑΏۃL�[��o�^
'*
'*�y�p�����[�^�z
'*      Name                �@�@  Type                Comment
'*      intSaibanNo;              Integer;            �̔�NO;
'*
'*�y�߂�l�z
'*      Name                �@�@  Type                Comment
'*R     �Ȃ�
'*
'******************************************************************************
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)       EDP-No      ����
'*M 2008/02/29; �a�c(IC);         P-32685;    ���i�I��v�e�V�X�e���̌����E�����ύX�o�^���̒l�����`�F�b�N�@�\�ǉ�;
'*M 2014/03/04; �a�c(IC);         P-46854;    �����𕪊�(�^�C���A�E�g�Ή�);
'*< Comment End >**************************************************************
Private Sub InsertBaikaKeyWork(ByVal intSaibanNo As Integer)

    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim strSQL As String
    
    On Error GoTo DefectErrorHandler

    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.InsertBaikaKeyWork(" & intSaibanNo & _
                                                             ")"
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL��
    '�����i�I��WF(�d���ύX���)�̃L�[�o�^
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
    strSQL = strSQL & "     ,w_siire.henkou_date + '2' AS genka_henkou_date"    '�����ύX���{�e�[�u���敪
    strSQL = strSQL & "     ,'2'"      '�e�[�u���敪"
    strSQL = strSQL & "     ,'1' AS syori_kbn"         '�����ύX(�Œ�l)
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
                            '���i�I��WF(�d���ύX���)
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

    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)


    'SQL��
    '���d���ύX���̃L�[�o�^
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
    strSQL = strSQL & "     ,v_siire.henkou_date + '1' AS genka_henkou_date"    '�����ύX���{�e�[�u���敪
    strSQL = strSQL & "     ,'1'"      '�e�[�u���敪"
    strSQL = strSQL & "     ,'1' AS syori_kbn"         '�����ύX(�Œ�l)
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
                            '�d���ύX���
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

    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)


    'SQL��
    '���X�ʏ��̃L�[�o�^
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
    strSQL = strSQL & "     ,'19900101' + '0'AS genka_henkou_date"            'V�}�X�^�̕ύX����1990/01/01���Z�b�g �����ύX���{�e�[�u���敪
    strSQL = strSQL & "     ,'0'"      '�e�[�u���敪"
    strSQL = strSQL & "     ,'1' AS syori_kbn"         '�����ύX(�Œ�l)
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
                            '�X�ʏ��
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

    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)


    'SQL��
    '���Г���{���̃L�[�o�^
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
    strSQL = strSQL & "     ,'19900101' + '0'"            'V�}�X�^�̕ύX����1990/01/01���Z�b�g �����ύX���{�e�[�u���敪
    strSQL = strSQL & "     ,'0'"      '�e�[�u���敪"
    strSQL = strSQL & "     ,'1' AS syori_kbn"         '�����ύX(�Œ�l)
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
                            '�Г���{���
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

    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)


    'SQL��
    '�����߂̃L�[��o�^
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
    strSQL = strSQL & "             AND syori_kbn = '1'"    '�����ύX(�Œ�l)
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

    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)

    '�C���X�^���X�j��
    objCDBAccess = Nothing
    
    'commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
             'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    InsertNeire
'*�y�����T�v�z
'*?    �l�����x�����̏�����
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    intSaibanNo;         �̔�NO
'*P    strSyoriKbn;         1:���� 2:����
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    �Ȃ�
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2008/02/27; �a�c(IC); P-32685 ���i�I��WF�V�X�e���̌����E�����ύX�o�^���̒l�����`�F�b�N�@�\�ǉ�;
'*< Comment End >**************************************************'
Private Sub InsertNeire(ByVal intSaibanNo As Integer _
                          , ByVal strSyoriKbn As String)
    
    Dim strMethodCallInfo As String
    Dim strSQL As String
    Dim vntAryParam As Object
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    
    On Error GoTo DefectErrorHandler
   
    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.InsertNeire(" & intSaibanNo & _
                                                             "," & strSyoriKbn & _
                                                                    ")"
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
       
    'SQL��
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

    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@syori_kbn", DT_Char, 1, strSyoriKbn) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
  
    '�C���X�^���X�j��
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >*******************************************************************
'*
'*�y�v���V�[�W�����z
'*!     InsertGenkaWork
'*
'*�y�����T�v�z
'*?     ���i�I��WF(�l�����x��)WORK�ɒl�����ُ핪��o�^
'*
'*�y�p�����[�^�z
'*      Name                �@�@  Type                Comment
'*      strLoginUser;             String;             ���O�C�����[�U�[;
'*      intSaibanNo;              Integer;            �̔�NO;
'*      sngSyouhizeiRitu;         Single;             ����ŗ�(%);
'*
'*�y�߂�l�z
'*      Name                �@�@  Type                Comment
'*R     �Ȃ�
'*
'******************************************************************************
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)       EDP-No      ����
'*M 2008/02/29; �a�c(IC);         P-32685;    ���i�I��v�e�V�X�e���̌����E�����ύX�o�^���̒l�����`�F�b�N�@�\�ǉ�;
'*M 2009/04/22; Sureerat P.;      P-35471;    Mod check �l�������,�l��������;
'*M 2014/01/22; ������(��A) ;�@P-46854; �Ő������ɔ����Ή�(���i�I��WF�E�V���i���);
'*M 2014/03/04; �a�c(IC);         P-46854;    �����𕪊�(�^�C���A�E�g�Ή�);
'*M 2014/04/10; �a�c(IC);         P-46854;    KEY�o�^�ȈՉ�;
'*M 2014/04/18; �a�c(IC);         P-46854;    �����ύX���(�^�C���A�E�g�Ή�);
'*< Comment End >**************************************************************
Private Sub InsertGenkaWork(ByVal strLoginUser As String _
                      , ByVal intSaibanNo As Integer _
                      , ByVal sngSyouhizeiRitu As Single)

    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim strSQL As String
    
    On Error GoTo DefectErrorHandler

    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.InsertGenkaWork(" & strLoginUser & _
                                                            "," & intSaibanNo & _
                                                            "," & sngSyouhizeiRitu & _
                                                             ")"
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")

    Dim syouhizeiRitu '����ŗ�
    Dim zeiHanteiKbn  '�Ŕ���敪
    Dim objVswf
    Dim objStr
    
    '���i���o�^COM�̃C���X�^���X����
        objVswf = CreateObject("VSWF01BUS005.CVswfSyouhizei")
    
    '����ŗ��ƐŔ���敪�f�[�^���擾����
    objStr = objVswf.GetKaisibiJyouhou()
    syouhizeiRitu = Split(objStr, "|")(0)
    zeiHanteiKbn = Split(objStr, "|")(1)
    
        objVswf = Nothing

    'SQL��
    '�����i�I��WF(�����ύX���)�Ƃ̃`�F�b�N
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

    If zeiHanteiKbn = "0" Then  ' �ō��� �̏ꍇ
        '|--------------------(����(�Ŕ���)------------------------||--����---||--------------------(����(�Ŕ���)--------------|
        '((����(�ō���) �~ ����ŗ� �� (����ŗ� �{ 100)�| ����Ŋz) - ����)  /  (����(�ō���) �~ ����ŗ� �� (����ŗ� �{ 100)
        strSQL = strSQL & "         ,((sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) - sel.genka) / (sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) * 100 AS neire_ritu"     '�l����
    Else                        ' �Ŕ��� �̏ꍇ
        '(����(�Ŕ���) �| ����) �� ����(�Ŕ���)
        strSQL = strSQL & "         ,round((sel.baika- sel.genka) / sel.baika,3) * 100 AS neire_ritu "      '�l����
    End If

    strSQL = strSQL & "         ,'2' AS touroku_naiyou"   '�d���ύX(�Œ�l)
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
                                    '���i�I��WF(�����ύX���)
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
    
    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@saiban_no2", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    
    'SQL��
    '�������ύX���Ƃ̃`�F�b�N
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

    If zeiHanteiKbn = "0" Then  ' �ō��� �̏ꍇ
        '|--------------------(����(�Ŕ���)------------------------||--����---||--------------------(����(�Ŕ���)--------------|
        '((����(�ō���) �~ ����ŗ� �� (����ŗ� �{ 100)�| ����Ŋz) - ����)  /  (����(�ō���) �~ ����ŗ� �� (����ŗ� �{ 100)
        strSQL = strSQL & "         ,((sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) - sel.genka) / (sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) * 100 AS neire_ritu"     '�l����
    Else                        ' �Ŕ��� �̏ꍇ
        '(����(�Ŕ���) �| ����) �� ����(�Ŕ���)
        strSQL = strSQL & "         ,round((sel.baika- sel.genka) / sel.baika,3) * 100 AS neire_ritu "      '�l����
    End If

    strSQL = strSQL & "         ,'2' AS touroku_naiyou"   '�d���ύX(�Œ�l)
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
                                    '�����ύX���
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
    
    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@saiban_no2", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    'SQL��
    '���X�ʏ��Ƃ̃`�F�b�N
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

    If zeiHanteiKbn = "0" Then  ' �ō��� �̏ꍇ
        '|--------------------(����(�Ŕ���)------------------------||--����---||--------------------(����(�Ŕ���)--------------|
        '((����(�ō���) �~ ����ŗ� �� (����ŗ� �{ 100)�| ����Ŋz) - ����)  /  (����(�ō���) �~ ����ŗ� �� (����ŗ� �{ 100)
        strSQL = strSQL & "         ,((sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) - sel.genka) / (sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) * 100 AS neire_ritu"     '�l����
    Else                        ' �Ŕ��� �̏ꍇ
        '(����(�Ŕ���) �| ����) �� ����(�Ŕ���)
        strSQL = strSQL & "         ,round((sel.baika- sel.genka) / sel.baika,3) * 100 AS neire_ritu "      '�l����
    End If

    strSQL = strSQL & "         ,'2' AS touroku_naiyou"   '�d���ύX(�Œ�l)
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
    strSQL = strSQL & "             ,'19900101' baika_henkou_date"    'V�}�X�^�̕ύX����1990/01/01���Z�b�g
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
                                    '�X�ʏ��
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
    
    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@saiban_no2", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    'SQL��
    '���Г���{���Ƃ̃`�F�b�N
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

    If zeiHanteiKbn = "0" Then  ' �ō��� �̏ꍇ
        '|--------------------(����(�Ŕ���)------------------------||--����---||--------------------(����(�Ŕ���)--------------|
        '((����(�ō���) �~ ����ŗ� �� (����ŗ� �{ 100)�| ����Ŋz) - ����)  /  (����(�ō���) �~ ����ŗ� �� (����ŗ� �{ 100)
        strSQL = strSQL & "         ,((sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) - sel.genka) / (sel.baika - ROUND(sel.baika * " & syouhizeiRitu & " / (100 + " & syouhizeiRitu & "),0,1)) * 100 AS neire_ritu"     '�l����
    Else                        ' �Ŕ��� �̏ꍇ
        '(����(�Ŕ���) �| ����) �� ����(�Ŕ���)
        strSQL = strSQL & "         ,round((sel.baika- sel.genka) / sel.baika,3) * 100 AS neire_ritu "      '�l����
    End If

    strSQL = strSQL & "         ,'2' AS touroku_naiyou"   '�d���ύX(�Œ�l)
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
    strSQL = strSQL & "             ,'19900101' AS baika_henkou_date"    'V�}�X�^�̕ύX����1990/01/01���Z�b�g
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
                                    '�Г���{���
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
    
    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        , MakeParam("@saiban_no2", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    '�C���X�^���X�j��
    objCDBAccess = Nothing
    
    'commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
             'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >*******************************************************************
'*
'*�y�v���V�[�W�����z
'*!     InsertGenkaKeyWork
'*
'*�y�����T�v�z
'*?     ���i�I��WF(�l�����x��)WORK�ɒl�����ُ핪��o�^
'*
'*�y�p�����[�^�z
'*      Name                �@�@  Type                Comment
'*      intSaibanNo;              Integer;            �̔�NO;
'*
'*�y�߂�l�z
'*      Name                �@�@  Type                Comment
'*R     �Ȃ�
'*
'******************************************************************************
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)       EDP-No      ����
'*M 2008/02/29; �a�c(IC);         P-32685;    ���i�I��v�e�V�X�e���̌����E�����ύX�o�^���̒l�����`�F�b�N�@�\�ǉ�;
'*M 2014/03/04; �a�c(IC);         P-46854;    �����𕪊�(�^�C���A�E�g�Ή�);
'*< Comment End >**************************************************************
Private Sub InsertGenkaKeyWork(ByVal intSaibanNo As Integer)

    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim strSQL As String
    
    On Error GoTo DefectErrorHandler

    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.InsertGenkaKeyWork(" & intSaibanNo & _
                                                             ")"
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL��
    '�����i�I��WF(�����ύX���)�̃L�[�o�^
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
    strSQL = strSQL & "     ,w_baika.henkou_date + '2' AS baika_henkou_date"    '�����ύX���{�e�[�u���敪
    strSQL = strSQL & "     ,jan.henkou_date AS genka_henkou_date"
    strSQL = strSQL & "     ,'2'"      '�e�[�u���敪"
    strSQL = strSQL & "     ,'2' AS syori_kbn"         '�����ύX(�Œ�l)
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
                            '���i�I��WF(�����ύX���)
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
    
    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    
    'SQL��
    '�������ύX���̃L�[�o�^
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
    strSQL = strSQL & "     ,v_baika.henkou_date + '1' "    '�����ύX���{�e�[�u���敪
    strSQL = strSQL & "     ,jan.henkou_date"
    strSQL = strSQL & "     ,'1'"      '�e�[�u���敪"
    strSQL = strSQL & "     ,'2' AS syori_kbn"         '�����ύX(�Œ�l)
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
                            '�����ύX���
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
    
    
    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    
    'SQL��
    '���X�ʏ��̃L�[�o�^
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
    strSQL = strSQL & "     ,'19900101' + '0'"    'V�}�X�^�̕ύX����1990/01/01���Z�b�g �����ύX���{�e�[�u���敪
    strSQL = strSQL & "     ,jan.henkou_date"
    strSQL = strSQL & "     ,'0'"      '�e�[�u���敪"
    strSQL = strSQL & "     ,'2' AS syori_kbn"         '�����ύX(�Œ�l)
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
                            '�X�ʏ��
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
    
    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    
    'SQL��
    '���Г���{���̃L�[�o�^
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
    strSQL = strSQL & "     ,'19900101' + '0'"    'V�}�X�^�̕ύX����1990/01/01���Z�b�g �����ύX���{�e�[�u���敪
    strSQL = strSQL & "     ,jan.henkou_date"
    strSQL = strSQL & "     ,'0'"      '�e�[�u���敪"
    strSQL = strSQL & "     ,'2' AS syori_kbn"         '�����ύX(�Œ�l)
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
                            '�Г���{���
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
    
    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    
    '�����߂̃L�[��o�^
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
    strSQL = strSQL & "             AND syori_kbn = '2'"    '�����ύX(�Œ�l)
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

    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no1", DT_Int, 2, intSaibanNo) _
                        )
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    '�C���X�^���X�j��
    objCDBAccess = Nothing
    
    'commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
             'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >*******************************************************************
'*
'*�y�v���V�[�W�����z
'*!     GetNVHInfo
'*
'*�y�����T�v�z
'*?     NVH�X�܂̎擾�B
'*
'*�y�p�����[�^�z
'*      Name                �@�@  Type                Comment
'*      strJanCd;                 String;             JAN�R�[�h;
'*
'*�y�߂�l�z
'*      Name                �@�@  Type                Comment
'*R     GetNVHInfo;               ADODB.Recordset;   ���R�[�h�Z�b�g;
'*
'******************************************************************************
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)       EDP-No      ����
'*M 2008/09/19; �a�c(IC);         P-34312;    �m�u�g�X�܂ւ̏��i�}�X�^�[�ύX�o�^��Ƃ̊ȑf��;
'*< Comment End >**************************************************************
Private Function GetNVHInfo(ByVal strJanCd As String) As ADODB.Recordset
                               
    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    Dim objCDBAccess    As COMDBUtil.CDBAccess
    Dim strSQL          As String
    
    On Error GoTo DefectErrorHandler

    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.GetNVHInfo(" & strJanCd & _
                                                            ")"
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL��
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

    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@strJanCd", DT_Char, 13, strJanCd))
    
    '�ؒfRecordset�擾
        GetNVHInfo = objCDBAccess.GetDCRecordset("VSWF", strSQL, vntAryParam)
    
    '�C���X�^���X�j��
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Function

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
             'Err.Number, Err.Source, Err.Description
End Function

'*< TOSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    FncCheckDB_Group
'*�y�����T�v�z
'*?    �O���[�v�i�`�F�b�N(�O���[�v�i�̏ꍇ�G���[)
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strJanCd;         JAN�R�[�h
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    �Ȃ�
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2009/05/12; �a�c; P-35471 �V�K�쐬;
'*< Comment End >**************************************************'
Private Function FncCheckDB_Group(ByVal strJanCd As String) As Boolean
    
    Dim strMethodCallInfo As String
    Dim strSQL As String
    Dim vntAryParam As Object
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim rsGrp As ADODB.Recordset
    On Error GoTo DefectErrorHandler
   
    FncCheckDB_Group = True
    
    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.FncCheckDB_Group(" & strJanCd & _
                                                                    ")"
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
       
    'SQL��
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

    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@strJanCd", DT_Char, 13, strJanCd))
    
    '�ؒfRecordset�擾
        rsGrp = objCDBAccess.GetDCRecordset("VSWF", strSQL, vntAryParam)
    
    If Not rsGrp.EOF Then
        If rsGrp.Fields("tanpin_kbn").Value = "1" Then
            FncCheckDB_Group = False
        End If
    End If
    
    '�C���X�^���X�j��
    rsGrp.Close
        rsGrp = Nothing
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Function

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
             'Err.Number, Err.Source, Err.Description
End Function
'******************************************************************'
'*�y�v���V�[�W�����z
'*!    UpdGenka_HattuTani
'*�y�����T�v�z
'*?    INSERT���쐬�i���i�I��WF(�߂����i)�j
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    arryCSVData() ; Object CSV�f�[�^���P�s���i�[����z��ւ̃|�C���^
'*�y�߂�l�z
'*    < Name Type Comment >
'*R
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2009/06/08; Supattra T(Tostem Thai); �V�K�쐬;
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
    
    '���Ə��}�X�^(v_jigyousyo)���A�X�܂��擾����B
    Dim strSQL  As String

    'strMethodCallInfo = App.Title & ".VswfIkatuTou.UpdGenka_HattuTani(�ꊇ�o�^)"
    
    
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

    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >****************************************************'
'*�y�v���V�[�W�����z
'*!    DeleteJyogaiTenpo
'*�y�����T�v�z
'*?    �ύX���O�ō쐬���ꂽ�X�܂�NULL�����폜����
'*?xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'*�y�p�����[�^�z
'*    < Name Type Comment >
'*P    strKikakuNo;         ���No
'*�y�߂�l�z
'*    < Name Type Comment >
'*R    �Ȃ�
'******************************************************************'
'*�y�ύX�����z
'*    < ���t ���O(��Ж�) ���� >
'*M    2009/07/23; �a�c(IC); P-36097 ���Ԍ���̑����i�i�h�܂链�h�Ȃǁj�̃��[�N�t���[�ւ̓o�^�V�X�e���\�z;
'*< Comment End >**************************************************'
Public Sub DeleteJyogaiTenpo(ByVal strKikakuNo As String)
    
    Dim strMethodCallInfo As String
    Dim strSQL As String
    Dim objCDBAccess As Object = CreateObject("COMDBUtil.CDBAccess")
    Dim vntAryParam As Object
    
    On Error GoTo DefectErrorHandler
   
    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.DeleteJyogaiTenpo(" & strKikakuNo & ")"
    
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
       
    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@KikakuSyuruiCd", DT_Char, 3, Left(strKikakuNo, 3)) _
                        , MakeParam("@NenRenNo", DT_Char, 10, Mid(strKikakuNo, 4, 10)) _
                        , MakeParam("@EdaNo", DT_Char, 3, Right(strKikakuNo, 3)) _
                        )
    
    '�����i�I��WF(�����ύX���)�폜
    'SQL��
    strSQL = ""
    strSQL = strSQL & " DELETE FROM"
    strSQL = strSQL & "     t_syousen_baika_henkou"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     kikaku_syurui_cd = ?"
    strSQL = strSQL & "     AND nen_ren_no = ?"
    strSQL = strSQL & "     AND eda_no = ?"
    strSQL = strSQL & "     AND baika IS NULL"

    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
  
    '�����i�I��WF(�d���ύX���)�폜
    'SQL��
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

    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
  
    '�����i�I��DB(�Z���^�[�R�[�h�ύX���)�폜
    'SQL��
    strSQL = ""
    strSQL = strSQL & " DELETE FROM"
    strSQL = strSQL & "     t_syousen_center_cd_henkou"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     kikaku_syurui_cd = ?"
    strSQL = strSQL & "     AND nen_ren_no = ?"
    strSQL = strSQL & "     AND eda_no = ?"
    strSQL = strSQL & "     AND center_kbn IS NULL"

    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    '���߂����i�폜
    'SQL��
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

    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
  
    '�C���X�^���X�j��
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
            'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >*******************************************************************'
'*�y�v���V�[�W�����z
'*!     UpdateNVHModosi
'*�y�����T�v�z
'*?     NVH�X�ܕύX���o�^����
'*�y�p�����[�^�z
'*      Name                Type            Comment
'*P     strKikakuNo;        <String>;       ���NO;
'*P     strJanCd;�@         <String>;       JAN�R�[�h;
'*P     strModosiGenka;     <String>;       �߂�����;
'*P     strModosiBaika;     <String>;       �߂�����;
'*P     strModosiHattyuuTani;     <String>;       �߂������P��;
'*P     strPOPSize;         <String>;       POP�T�C�Y;
'*P     strTenCd;           <String>;       �X�R�[�h;
'*P     strTantousyaId;     <String>;       �S����ID;
'*�y�߂�l�z
'*      Name                    Type        Comment
'*R     <UpdateNVHModosi>;            <Boolean>;  ����/���s
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)            EDP-No      ����
'*M 2009/08/06; �a�c(IC);              P-36097;    �V�K�쐬;
'*M 2010/03/26; �a�c(IC);              P-37515;    �ύX���O�X�ܓo�^���Ȃ��ꍇ�̑Ή�;
'*< Comment End >**************************************************************
Public Function UpdateNVHModosi(ByVal strKikakuNo As String, ByVal strJanCd As String, _
                    ByVal strModosiGenka As String, ByVal strModosiBaika As String, _
                    ByVal strModosiHattyuuTani As String, ByVal strPOPSize As String, _
                    ByVal strTenCd As String, ByVal strTantousyaId As String) As Boolean

    Dim objCDBAccess            As COMDBUtil.CDBAccess  'DB�I�u�W�F�N�g
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
    Dim strGenkaJyogaiKbn       As String   '�������O�敪
    Dim strbaikaJyogaiKbn       As String   '�������O�敪
    Dim strHattyuuTaniJyogaiKbn     As String   '�����P�ʏ��O�敪

    On Error GoTo DefectErrorHandler

    'strMethodCallInfo = App.Title & ".CHacTenCom.UpdateNVHModosi(" & _
                        strKikakuNo & "," & strJanCd & "," & _
                        strModosiGenka & "," & strModosiBaika & "," & _
                        strModosiHattyuuTani & "," & strPOPSize & "," & _
                        strTenCd & "," & strTantousyaId & "," & ")"
                            
    '���O�X�܏��̎擾
        objJyogai = GetJyogaiKbn(strKikakuNo, strTenCd)
    
    If Not objJyogai.EOF Then
        '�������O�敪
        strGenkaJyogaiKbn = objJyogai.Fields("genka_jyogai_kbn").Value
        '�������O�敪
        strbaikaJyogaiKbn = objJyogai.Fields("baika_jyogai_kbn").Value
        '�����P�ʏ��O�敪
        strHattyuuTaniJyogaiKbn = objJyogai.Fields("hattyuu_tani_jyogai_kbn").Value
    Else
        '�������O�敪
        strGenkaJyogaiKbn = "0"
        '�������O�敪
        strbaikaJyogaiKbn = "0"
        '�����P�ʏ��O�敪
        strHattyuuTaniJyogaiKbn = "0"
    End If

    '�I�u�W�F�N�g�j��
    objJyogai.Close
        objJyogai = Nothing

    If strGenkaJyogaiKbn = "1" Then strModosiGenka = ""
    If strbaikaJyogaiKbn = "1" Then strModosiBaika = ""
    If strHattyuuTaniJyogaiKbn = "1" Then strModosiHattyuuTani = ""
    
    '���NO�������ރR�[�h�A�N�A�ԁA�}�Ԃ��擾
    strKikakuSyuruiCD = Left(strKikakuNo, 3)
    strNenRenNO = Mid(strKikakuNo, 4, 10)
    strEdaNO = Right(strKikakuNo, 3)
       
    '�C���X�^���X�쐬
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
                
    '�N���̎擾
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
    
    '�I�u�W�F�N�g�j��
    objNengetu.Close
        objNengetu = Nothing
    
    '���݃`�F�b�N
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
    
    'SQL���s
    lngSqlRtn = objCDBAccess.ExecuteSql("VSWF", strSQL)
    
    '�C���X�^���X�j��
    objRs.Close
        objRs = Nothing
    objCDBAccess = Nothing
    
    'GetObjectContext.SetComplete
    
    UpdateNVHModosi = True
    
Exit Function
DefectErrorHandler:

    UpdateNVHModosi = False
    
    '�G���[����
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
      
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
                  'Err.Number, Err.Source, Err.Description
End Function

'*< TOSTEM >*******************************************************************'
'*�y�v���V�[�W�����z
'*!     GetModosiCount
'*�y�����T�v�z
'*?     �߂����i�e�[�u���̑��݃`�F�b�N
'*�y�p�����[�^�z
'*      Name                Type            Comment
'*P     strNengetu;         <String>;       �N��;
'*P     strJanCd;           <String>;       JAN�R�[�h;
'*P     strTenCd;           <String>;       �X�R�[�h;
'*�y�߂�l�z
'*      Name                    Type                Comment
'*R     <GetModosiCount>;       <ADODB.Recordset>;  ���݃`�F�b�N
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)            EDP-No      ����
'*M 2009/08/06; �a�c(IC);              P-36097;    �V�K�쐬;

'*< Comment End >**************************************************************
Private Function GetModosiCount(ByVal strNengetu As String, ByVal strJanCd As String, _
                           ByVal strTenCd As String) As ADODB.Recordset

    Dim strSQL                  As String
    Dim objCDBAccess            As COMDBUtil.CDBAccess  'DB�I�u�W�F�N�g
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
    
    '�C���X�^���X�쐬
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'Recordset�擾
        GetModosiCount = objCDBAccess.GetDCRecordset("VSWF", strSQL, vntAryParam)
    '�C���X�^���X�j��
    objCDBAccess = Nothing
    
Exit Function
DefectErrorHandler:

    '�G���[����
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
      
    'If Not GetObjectContext Is Nothing Then
        'GetObjectContext.SetAbort
    End If
      
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
                  'Err.Number, Err.Source, Err.Description
End Function

'*< TOSTEM >*******************************************************************'
'*�y�v���V�[�W�����z
'*!     GetJyogaiKbn
'*�y�����T�v�z
'*?     ���i�I��DB(�ύX���O�X��)�����擾����
'*�y�p�����[�^�z
'*      Name                Type            Comment
'*P     strKikakuNo;        <String>;       ���NO;
'*P     strTenCd;           <String>;       �X�R�[�h;
'*�y�߂�l�z
'*      Name                    Type                 Comment
'*R     <GetSiireJyogaiKbn>;    <ADODB.Recordset>;   �d�����O�敪
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)            EDP-No      ����
'*M 2009/08/06; �a�c(IC);              P-36097;    �V�K�쐬;
'*< Comment End >**************************************************************
Public Function GetJyogaiKbn(ByVal strKikakuNo As String, ByVal strTenCd As String) As ADODB.Recordset

    Dim objCDBAccess            As COMDBUtil.CDBAccess  'DB�I�u�W�F�N�g
    Dim strSQL                  As String   'SQL
    Dim strMethodCallInfo       As String
    Dim strKikakuSyuruiCD       As String
    Dim strNenRenNO             As String
    Dim strEdaNO                As String
    Dim vntAryParam             As Object
    
    On Error GoTo DefectErrorHandler
    
    'strMethodCallInfo = App.Title & ".CHacTenCom.GetJyogaiKbn(" & _
                        strKikakuNo & "," & strTenCd & "," & ")"
                                                
    '���NO�������ރR�[�h�A�N�A�ԁA�}�Ԃ��擾
    strKikakuSyuruiCD = Left(strKikakuNo, 3)
    strNenRenNO = Mid(strKikakuNo, 4, 10)
    strEdaNO = Right(strKikakuNo, 3)

    '�C���X�^���X�쐬
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
                
    '�X�V�pSQL�쐬
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
    
    'Recordset�擾
        GetJyogaiKbn = objCDBAccess.GetDCRecordset("VSWF", strSQL, vntAryParam)

    '�C���X�^���X�j��
    objCDBAccess = Nothing
    
    'GetObjectContext.SetComplete
    
Exit Function
DefectErrorHandler:

    '�G���[����
    If Not objCDBAccess Is Nothing Then
        objCDBAccess = Nothing
    End If
      
    'If Not GetObjectContext Is Nothing Then
        'GetObjectContext.SetAbort
    End If
      
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
                  'Err.Number, Err.Source, Err.Description
End Function

'***** 2018/11/16 ���O���ׂ����o�������ׁA���C ***********************************************************************************
''*< TOSTEM >*******************************************************************'
''*�y�v���V�[�W�����z
''*!     SetLog
''*�y�����T�v�z
''*?     �ꊇ�ύX�o�^�̏����J�n���ƏI�����Ƀ��O��o�^����
''*�y�p�����[�^�z
''*      Name                Type            Comment
''*P     strSyoriKbn;        <String>;       �ꊇ�ύX�o�^����;
''*�y�߂�l�z
''*      Name                    Type                 Comment
''*R     �Ȃ�;
''******************************************************************************
''*�y�ύX�����z
''*  ���t        ���O(��Ж�)            EDP-No      ����
''*M 2015/09/03; �a�c(IC);               P-51672;    �V�K�쐬;
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
'    '�G���[�\���p
'    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.SetLog(" & strSyoriKbn & ")"
'
'    'CDBAccess�C���X�^���X����
'    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
'
'    'SQL��
'    strSQL = ""
'    strSQL = strSQL & " INSERT INTO"
'    strSQL = strSQL & "     t_ikkatu_henkou_touroku_log"
'    strSQL = strSQL & " SELECT ?,?,?,GETDATE()"
'
'    'MakeParam�̍쐬
'    vntAryParam = Array(MakeParam("@renban", DT_Long, 4, lngRenban), _
'                        MakeParam("@syori_kbn", DT_String, 2, strSyoriKbn), _
'                        MakeParam("@user_id", DT_String, 40, strLoginUser) _
'                        )
'
'    'SQL���s
'    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
'
'    '�C���X�^���X�j��
'    objCDBAccess = Nothing
'
'    'Commit
'    'GetObjectContext.SetComplete
'Exit Sub
'
'DefectErrorHandler:
'    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
'    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
'    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
'             'Err.Number, Err.Source, Err.Description
'End Sub

'*< TOSTEM >*******************************************************************'
'*�y�v���V�[�W�����z
'*!     SetLog
'*�y�����T�v�z
'*?     �ꊇ�ύX�o�^�������Ń��O��o�^�E�X�V����
'*�y�p�����[�^�z
'*      Name                Type            Comment
'*P1    lngRenban;          <String>;       �A��;
'*P2    strSyoriFile;       <String>;       �����t�@�C��(�������~���A���i��{���A���ޏ��A�X�ʏ��);
'*P3    lngDataCnt;         <Long>;         �f�[�^��;
'*P5    strSyoriMethod;     <String>;       �������\�b�h;
'*P6    strSyoriSyousai;    <String>;       �����ڍ�;
'*P7    strLoginUser;       <String>;       ���[�U�[ID;
'*P8    strUpdateKbn;       <String>;       �f�[�^�X�V�敪(1�F�o�^�J�n�A2�F�����J�n�A�R�F�����I���A�S�F�o�^�I��);
'*�y�߂�l�z
'*      Name                    Type                 Comment
'*R     �Ȃ�;
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)            EDP-No      ����
'*M 2018/11/19; �R�c(IC);               P-XXXXX;    �V�K�쐬;
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

    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.SetLog(" & strSyoriFile & ")"
    
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL��
    strSQL = ""
    
    '�X�V�敪���ƂɎ��s����SQL��ݒ�
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

            
'            'MakeParam�̍쐬
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
    
    'SQL���s
    'Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    '�C���X�^���X�j��
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
             'Err.Number, Err.Source, Err.Description
End Sub

'*< TOSTEM >*******************************************************************
'*
'*�y�v���V�[�W�����z
'*!     GetSaibanNo2
'*
'*�y�����T�v�z
'*?     �̔�NO�̎擾�B
'*
'*�y�p�����[�^�z
'*      Name                �@�@  Type                Comment
'*P     strSaibanKbn;             string              �̔ԋ敪
'*
'*�y�߂�l�z
'*      Name                �@�@  Type                Comment
'*R     GetSaibanNo;              integer;            �̔�NO;
'*
'******************************************************************************
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)       EDP-No      ����
'*M 2015/09/03; �a�c(IC);          P-51672;    �V�K�쐬;
'*< Comment End >**************************************************************
Public Function GetSaibanNo2(ByVal strSaibanKbn As String) As Long
                               
    Dim strMethodCallInfo As String
    Dim vntAryParam As Object
    Dim objCDBAccess    As COMDBUtil.CDBAccess
    Dim strSQL          As String
    Dim rsSaiban        As ADODB.Recordset
    
    On Error GoTo DefectErrorHandler

    '�G���[�\���p
    'strMethodCallInfo = "VswfIkatuTou.CHacTenCom.GetSaibanNo2(" & strSaibanKbn & ")"
    
    'CDBAccess�C���X�^���X����
    objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL��
    strSQL = ""
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     saiban_no + 1 AS saiban_no"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     w_syousen_neire_keikoku_saiban WITH(UPDLOCK)"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     saiban_kbn = ?"

    '�ؒfRecordset�擾
        rsSaiban = objCDBAccess.GetDCRecordset("VSWF", strSQL, MakeParam("@saiban_kbn", DT_Char, 1, strSaibanKbn))
    
    GetSaibanNo2 = rsSaiban.Fields("saiban_no").Value
    rsSaiban.Close
        rsSaiban = Nothing
    
     'SQL��
    strSQL = ""
    strSQL = strSQL & " UPDATE"
    strSQL = strSQL & "     w_syousen_neire_keikoku_saiban"
    strSQL = strSQL & " SET"
    strSQL = strSQL & "     saiban_no = ?"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     saiban_kbn = ?"
   
    'MakeParam�̍쐬
    vntAryParam = Array(MakeParam("@saiban_no", DT_Long, 4, GetSaibanNo2), _
                        MakeParam("@saiban_kbn", DT_Char, 1, strSaibanKbn) _
                        )
    
    'SQL���s
    Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    '�C���X�^���X�j��
    objCDBAccess = Nothing
    
    'Commit
    'GetObjectContext.SetComplete
Exit Function

DefectErrorHandler:
    If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    'If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
             'Err.Number, Err.Source, Err.Description
End Function


'*< TOSTEM >*******************************************************************'
'*�y�v���V�[�W�����z
'*!     LockTableIkkatu
'*�y�����T�v�z
'*?     �ꊇ�ύX�o�^�������Ń��O��o�^�E�X�V����
'*�y�p�����[�^�z
'*      Name                Type            Comment
'*�y�߂�l�z
'*      Name                    Type                 Comment
'*R     �Ȃ�;
'******************************************************************************
'*�y�ύX�����z
'*  ���t        ���O(��Ж�)            EDP-No      ����
'*M 2018/11/19; ������(��A);               P-XXXXX;    �V�K�쐬;
'*< Comment End >**************************************************************
Public Sub LockTableIkkatu(ByRef objCDBAccess)
        '2018/12/14; ������(��A) �ꊇ�o�^���A�������[�U�[�̂ݎ��s�ł��� �ǉ���
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
        'SQL���s
        Call objCDBAccess.ExecuteSql("VSWF", strSQL)
        Dim mydt
        mydt = Now()
        pStrRenban = Year(mydt) & Month(mydt) & Day(mydt) & "_" & Hour(mydt) & Minute(mydt) & Second(mydt) & "." & Right("000" & ((Timer() * 1000) Mod 1000), 3)
        
        '2018/12/14; ������(��A) �ꊇ�o�^���A�������[�U�[�̂ݎ��s�ł��� �ǉ���
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

    '�G���[�\���p
    ''strMethodCallInfo = "VswfIkatuTou.CHacTenCom.SetIkkatuLog()"
    
    'CDBAccess�C���X�^���X����
    'objCDBAccess = CreateObject("COMDBUtil.CDBAccess")
    
    'SQL��
    strSQL = ""
    
    syori_step = Right("                    " & syori_step, 20)
    

    Dim mydt
    
    If pStrRenban = "" Then
        mydt = Now()
        pStrRenban = Year(mydt) & Month(mydt) & Day(mydt) & "_" & Hour(mydt) & Minute(mydt) & Second(mydt) & "." & Right("000" & ((Timer() * 1000) Mod 1000), 3)
    End If
    
    '�X�V�敪���ƂɎ��s����SQL��ݒ�
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
    'SQL���s
    'Call objCDBAccess.ExecuteSql("VSWF", strSQL, vntAryParam)
    
    '�C���X�^���X�j��
    'objCDBAccess = Nothing
    
    'Commit
    ''GetObjectContext.SetComplete
Exit Sub

DefectErrorHandler:
    'If Not objCDBAccess Is Nothing Then objCDBAccess = Nothing
    ''If Not GetObjectContext Is Nothing Then 'GetObjectContext.SetAbort
    'ErrRaise DError.AbnormalDError, strMethodCallInfo, "�\�����Ȃ��G���[�ł��B", _
             'Err.Number, Err.Source, Err.Description
End Sub


    Private Function IsNull(ByVal v As Object) As Boolean
        Return IsDBNull.vaue = True
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub
End Class