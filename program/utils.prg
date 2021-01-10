#INCLUDE	INCLUDE\MAIN.H

FUNCTION IsN_VMP(para1)
 PRIVATE m.is_ok, m.n_vmp
 m.is_ok = .F.
 m.n_vmp = para1
 IF LEN(m.n_vmp)<>17
  RETURN .F.
 ENDIF 
 IF !(SUBSTR(m.n_vmp,3,1)='.' AND SUBSTR(m.n_vmp,8,1)='.' AND SUBSTR(m.n_vmp,14,1)='.')
  RETURN .F.
 ENDIF 
 IF !(ISDIGIT(SUBSTR(m.n_vmp,1,1)) AND ISDIGIT(SUBSTR(m.n_vmp,2,1)) AND ;
 	ISDIGIT(SUBSTR(m.n_vmp,4,1)) AND ISDIGIT(SUBSTR(m.n_vmp,5,1)) AND ISDIGIT(SUBSTR(m.n_vmp,6,1)) AND ISDIGIT(SUBSTR(m.n_vmp,7,1)) AND ;
 	ISDIGIT(SUBSTR(m.n_vmp,9,1)) AND ISDIGIT(SUBSTR(m.n_vmp,10,1)) AND ISDIGIT(SUBSTR(m.n_vmp,11,1)) AND ISDIGIT(SUBSTR(m.n_vmp,12,1)) AND ISDIGIT(SUBSTR(m.n_vmp,13,1)) AND ;
 	ISDIGIT(SUBSTR(m.n_vmp,15,1)) AND ISDIGIT(SUBSTR(m.n_vmp,16,1)) AND ISDIGIT(SUBSTR(m.n_vmp,17,1)))
  RETURN .F.
 ENDIF 
RETURN .T.

FUNCTION GoNWrkDays(stday, ndays)
 PRIVATE m.stday, m.ndays
 
 IF m.ndays=0
  RETURN m.stday
 ENDIF 
 
 m.IsUsedHolidays = .T.
 IF !USED('holidays')
  m.IsUsedHolidays = .F.
  IF OpenFile(pcommon+'\holidays', 'holidays', 'shar', 'day')>0
   IF USED('holidays')
    USE IN holidays
    RETURN {}
   ENDIF 
  ENDIF 
 ENDIF 
  
 m.curday  = m.stday
 m.wrkdays = 0
 DO WHILE m.wrkdays < m.ndays
  m.curday = m.curday + 1
  m.wrkdays = m.wrkdays + IIF(INLIST(DOW(m.curday,2),6,7) or ;
   SEEK(m.curday, 'holidays'),0,1)
 ENDDO 

 IF m.IsUsedHolidays = .F.
  USE IN holidays
 ENDIF 
 
RETURN m.curday

FUNCTION FStraf(para1, para2) && ��� ������, ���� ���������
 IF PARAMETERS()<2
  RETURN -1
 ENDIF 
 IF VARTYPE(para1)!='N'
  RETURN -1
 ENDIF 
 IF VARTYPE(para2)!='D'
  RETURN -1
 ENDIF 

 LOCAL m.cod, m.d_exp, m.s_s

 m.cod   = para1
 m.d_exp = para2
 m.s_s   = 0
 
 oal = ALIAS()
 
 m.lWasUsed = .T.
 IF !USED('pnyear')
  m.lWasUsed = .F.
  IF OpenFile(pCommon+'\pnyear', 'pnyear', 'shar', 'period')>0
   IF USED('pnyear')
    USE IN pnyear
   ENDIF 
   IF !EMPTY(oal)
    SELECT &oal
   ENDIF 
   RETURN -1
  ENDIF 
 ENDIF 
 
 IF m.d_exp < {01.06.2019}
  m.s_s = IIF(SEEK(LEFT(DTOS(m.d_exp),6), 'pnyear'), pnyear.pnorm, 0)
  RETURN m.s_s
 ENDIF 
 
 DO CASE 
  CASE INLIST(m.cod, 26210,26281,40040,40041,40042)
   m.s_s = IIF(SEEK(LEFT(DTOS(DATE()),6), 'pnyear'), pnyear.prd, 0)

  CASE IsUsl(m.cod)
   m.s_s = IIF(SEEK(LEFT(DTOS(DATE()),6), 'pnyear'), pnyear.app, 0)

  CASE IsMes(m.cod)
   m.s_s = IIF(SEEK(LEFT(DTOS(DATE()),6), 'pnyear'), pnyear.st, 0)

  CASE IsKD(m.cod)
   m.s_s = IIF(SEEK(LEFT(DTOS(DATE()),6), 'pnyear'), pnyear.dst, 0)

  CASE FLOOR(m.cod/1000)=96
   m.s_s = IIF(SEEK(LEFT(DTOS(DATE()),6), 'pnyear'), pnyear.smp, 0)

  CASE IsVMP(m.cod)
   m.s_s = IIF(SEEK(LEFT(DTOS(DATE()),6), 'pnyear'), pnyear.vmp, 0)
 ENDCASE 

 IF m.lWasUsed = .F.
  IF USED('pnyear')
   USE IN pnyear
  ENDIF 
 ENDIF 

 IF !EMPTY(oal)
  SELECT &oal
 ENDIF 

RETURN m.s_s

FUNCTION IsOkNVmpForEco(ppp)
 LOCAL m.isok, m.n_vmp

 m.isok = .T.
 m.n_vmp = ALLTRIM(ppp)

 IF LEN(m.n_vmp)<>12
 	m.isok = .f.
 ENDIF 
 IF m.isok = .t.
  IF !ISALPHA(SUBSTR(m.n_vmp,1,1))
   m.isok = .f.
  ENDIF 
 ENDIF 
 IF m.isok = .t.
  IF !ISDIGIT(SUBSTR(m.n_vmp,2,2))
   m.isok = .f.
  ENDIF 
 ENDIF 
 IF m.isok = .t.
  IF !ISALPHA(SUBSTR(m.n_vmp,4,1))
   m.isok = .f.
  ENDIF 
 ENDIF 
 IF m.isok = .t.
  IF !ISALPHA(SUBSTR(m.n_vmp,5,1))
   m.isok = .f.
  ENDIF 
 ENDIF 
 IF m.isok = .t.
  IF !ISALPHA(SUBSTR(m.n_vmp,6,1))
   m.isok = .f.
  ENDIF 
 ENDIF 
 IF m.isok = .t.
  IF !ISDIGIT(SUBSTR(m.n_vmp,7,6))
   m.isok = .f.
  ENDIF 
 ENDIF 
RETURN m.isok

FUNCTION saytipofexp(para1)
 m.et   = para1
 m.name = ''
 DO CASE 
  CASE m.et = '2'
   m.name = '�������� ������-������������� ����������'
  CASE m.et = '3'
   m.name = '������� ������-������������� ����������'
  CASE m.et = '7'
   m.name = '������������ ������-������������� ����������'
  CASE m.et = '8'
   m.name = '������-������������� ���������� �� ������'

  CASE m.et = '4'
   m.name = '�������� ���������� �������� ����������� ������'
  CASE m.et = '5'
   m.name = '������� ���������� �������� ����������� ������'
  CASE m.et = '6'
   m.name = '������������ ���������� �������� ����������� ������'
  CASE m.et = '9'
   m.name = '���������� �������� ����������� ������ �� ������'
 ENDCASE 
RETURN m.name

FUNCTION SayBriefTipOfExp(para1)
 PRIVATE m.et
 m.et   = para1
 m.name = ''
 DO CASE 
  CASE m.et = '2'
   m.name = '����'
  CASE m.et = '3'
   m.name = '����'
  CASE m.et = '7'
   m.name = '����'
  CASE m.et = '8'
   m.name = '����'

  CASE m.et = '4'
   m.name = '�����'
  CASE m.et = '5'
   m.name = '�����'
  CASE m.et = '6'
   m.name = '�����'
  CASE m.et = '9'
   m.name = '�����'
 ENDCASE 
RETURN m.name

FUNCTION NumActOfExp(para1, para2, para3, para4) && para1 - lpuid, para2 - et, para3 - reason, para4 - �����
* LOCAL para1, para2, para3, para4
 IF PARAMETERS()<4
  para4 = 0
 ENDIF 
 IF PARAMETERS()<3
  para3 = ' '
 ENDIF 
 IF PARAMETERS()<2
  RETURN ''
 ENDIF 

 LOCAL m.lpuid, m.IsMeeOrEkmp, m.PlanOrCel, m.reason, m.num
 
 m.lpuid       = STR(para1,4)
 m.IsMeeOrEkmp = IIF(INLIST(para2,'2','3','7','8'), '1', '2') && ���/����
 m.PlanOrCel   = IIF(INLIST(para2,'3','5'), '2', '1') && �������/��������

 DO CASE 
  CASE INLIST(para2,'2','4') && �������� ����������
   m.reason = '0'
  CASE INLIST(para2,'6','7') && ������������ ����������
   m.reason = '�'
  OTHERWISE 
   m.reason = IIF(INLIST(para3,'1','2','3','4','5','6'), para3, '0')
   *m.reason = para3
 ENDCASE 
 
 m.num = IIF(para4>0, PADL(para4,6,'0'), '')
 
 m.NumOfAct = m.qcod + m.lpuid + m.IsMeeOrEkmp + m.PlanOrCel + m.reason + m.num

RETURN m.NumOfAct

FUNCTION IsStac(m.mcod)
 m.IsStac = IIF(VAL(SUBSTR(m.mcod,3,2))>40,.t.,.f.)
RETURN m.IsStac

FUNCTION flmindate(m.flcod)
 m.result = {}
 IF LEN(m.flcod)!=12
  RETURN m.result
 ENDIF 
  MESSAGEBOX('!',0+64,'')
 FOR m.nm=0 TO 12
  m.result = IIF(SUBSTR(m.flcod,m.nm,1)='1', GOMONTH(m.tdat1,-(12-m.nm)), m.result)
  IF !EMPTY(m.result)
   EXIT 
  ENDIF 
 ENDFOR 
RETURN m.result

FUNCTION flmaxdate(m.flcod)
 m.result = {}
 IF LEN(m.flcod)!=12
  RETURN m.result
 ENDIF 
  MESSAGEBOX('!!',0+64,'')
 FOR m.nm=12 TO 0 STEP -1
  m.result = IIF(SUBSTR(m.flcod,m.nm,1)='1', GOMONTH(m.tdat2,-(12-m.nm)), m.result)
  IF !EMPTY(m.result)
   EXIT 
  ENDIF 
 ENDFOR 
RETURN m.result

FUNCTION IsVz
 PARAMETERS m.tipofpr, m.cod
RETURN 

FUNCTION TipOfPr
 PARAMETERS m.lpuobr, m.lpuprikl
 IF !USED('pilot')
  RETURN 0
 ENDIF 
 IF EMPTY(m.lpuprikl)
  RETURN 0
 ENDIF 
 IF m.lpuobr=m.lpuprikl
  RETURN 3
 ENDIF 
 IF SEEK(m.lpuprikl, 'pilot', 'mcod')
  RETURN 2
 ELSE 
  RETURN 1
 ENDIF 
RETURN 0

FUNCTION TipOfPrS
 PARAMETERS m.lpuobr, m.lpuprikl
 IF !USED('pilots')
  RETURN 0
 ENDIF 
 IF EMPTY(m.lpuprikl)
  RETURN 0
 ENDIF 
 IF m.lpuobr=m.lpuprikl
  RETURN 3
 ENDIF 
 IF SEEK(m.lpuprikl, 'pilots', 'mcod')
  RETURN 2
 ELSE 
  RETURN 1
 ENDIF 
RETURN 0

FUNCTION TipOfPaz(amcod,bmcod)
 PRIVATE lcmcod, lcprmcod, IsPilot, m.paztip
 m.lcmcod   = amcod
 m.lcprmcod = bmcod
 m.IsPilot = IIF(SEEK(m.lcprmcod, 'pilot', 'mcod'), .t., .f.)

 m.paztip = 0

 DO CASE 
  CASE EMPTY(m.lcprmcod) && �� ����������
   m.paztip = 0
  CASE m.lcmcod = m.lcprmcod && ���������� �� ����� ���������
   m.paztip = 1
  CASE m.lcmcod != m.lcprmcod AND m.IsPilot=.t. && ���������� � ������ �� �� ����� ���������
   m.paztip = 2
  CASE m.lcmcod != m.lcprmcod AND m.IsPilot=.f. && ���������� � �� ������ �� �� ����� ���������
   m.paztip = 3
  OTHERWISE 
   m.paztip = 0
 ENDCASE 

RETURN m.paztip

FUNCTION TipOfPazS(amcod,bmcod)
 PRIVATE lcmcod, lcprmcod, IsPilot, m.paztip
 m.lcmcod   = amcod
 m.lcprmcod = bmcod
 m.IsPilot = IIF(SEEK(m.lcprmcod, 'pilots', 'mcod'), .t., .f.)

 m.paztip = 0

 DO CASE 
  CASE EMPTY(m.lcprmcod) && �� ����������
   m.paztip = 0
  CASE m.lcmcod = m.lcprmcod && ���������� �� ����� ���������
   m.paztip = 1
  CASE m.lcmcod != m.lcprmcod AND m.IsPilot=.t. && ���������� � ������ �� �� ����� ���������
   m.paztip = 2
  CASE m.lcmcod != m.lcprmcod AND m.IsPilot=.f. && ���������� � �� ������ �� �� ����� ���������
   m.paztip = 3
  OTHERWISE 
   m.paztip = 0
 ENDCASE 

RETURN m.paztip

FUNCTION CloseAllTables
CLOSE TABLES ALL
CLOSE DATABASES ALL 

FUNCTION OpenFile
 LPARAMETERS lcFile, lcAlias, lcMode, lcOrder, lcAgain
 LOCAL loError AS Exception
 *lcFile = IIF(OCCURS('.', lcFile)>0, lcFile, lcFile+".dbf")
 lcFile  = IIF(LOWER(RIGHT(ALLTRIM(lcFile),4))='.', lcFile, lcFile+".dbf")
 lcMode  = IIF(!EMPTY(lcMode), UPPER(lcMode), "SHARED")
 lcOrder = IIF(!EMPTY(lcOrder), "ORDER "+UPPER(lcOrder), "")
 lcAgain = IIF(!EMPTY(lcAgain), UPPER(lcAgain), "")
 IF !FILE(lcFile)
  MESSAGEBOX("����������� ���� "+lcFile,0+16,"")
  RETURN 1
 ENDIF 
 
 TRY
  USE (lcFile) IN 0 ALIAS (lcAlias) &lcMode &lcOrder &lcAgain
 CATCH TO loError
  IF loError.ErrorNo!=12
  MESSAGEBOX("������ ��� �������� ����� "+lcFile+"!"+CHR(13)+;
  loError.Message + "," + ALLTRIM(STR(loError.ErrorNo)), 0, "������ �������� �����!")
  ENDIF 
 ENDTRY

 IF	VARTYPE(m.loError) == "O"
  RETURN 1
 ELSE
  RETURN 0
 ENDIF
 


FUNCTION RChar
 PARAMETERS rc_par
 IF LEN(ALLTRIM(rc_par)) > 0
  FOR i=1 TO LEN(ALLTRIM(rc_par))
   sub_tmp = SUBSTR(ALLTRIM(rc_par),i,1)
    IF !Lower(sub_tmp) $ '��������������������������������- '
     RETURN .f.
    ENDIF 
   NEXT  
  ENDIF 
RETURN .t.


FUNCTION EngToRus
 PARAMETERS string
 lowers = [qwertyuiop]+chr(91)+chr(93)+[asdfghjkl;'zxcvbnm,.QWERTYUIOP{}ASDFGHJKL:"ZXCVBNM<>]
 uppers = [����������������������������������������������������������������]
RETURN (chrtran(string,lowers,uppers))


FUNCTION RusToEng
 Para string
 lowers = [qwertyuiop]+chr(91)+chr(93)+[asdfghjkl;'zxcvbnm,.QWERTYUIOP{}ASDFGHJKL:"ZXCVBNM<>]
 uppers = [����������������������������������������������������������������]
* lowers = [qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM] + chr(241)
* uppers = [����������������������������������������������������] + chr(240)
RETURN (chrtran(string,uppers,lowers))


FUNCTION  cpr
PARAMETERS  _c
**  _c - ����� �� �����. �����������: �� 0 �� 999 ����������.
if _c<0
   return ''
endif
if _c=0
   return '���� ������'
endif
if _c>=1000000000000000
   return ''
endif
**  ����������� �������� � ����������:
**  m1(20,4), m2(5,3)
**  _p - �������� ����� �� ������.
set talk off
_c=1000*int(_c)+100*(round(_c,2)-int(_c))
_p=''
dime m1(20,4), m2(6,3)

m1(1,1)=''
m1(2,1)='����'
m1(3,1)='���'
m1(4,1)='���'
m1(5,1)='������'
m1(6,1)='����'
m1(7,1)='�����'
m1(8,1)='����'
m1(9,1)='������'
m1(10,1)='������'
m1(11,1)=''
m1(12,1)=''
m1(13,1)=''
m1(14,1)=''
m1(15,1)=''
m1(16,1)=''
m1(17,1)=''
m1(18,1)=''
m1(19,1)=''
m1(20,1)=''

m1(1,2)=''
m1(2,2)=''
m1(3,2)='��������'
m1(4,2)='��������'
m1(5,2)='�����'
m1(6,2)='���������'
m1(7,2)='����������'
m1(8,2)='���������'
m1(9,2)='�����������'
m1(10,2)='���������'
m1(11,2)='������'
m1(12,2)='�����������'
m1(13,2)='����������'
m1(14,2)='����������'
m1(15,2)='������������'
m1(16,2)='����������'
m1(17,2)='�����������'
m1(18,2)='����������'
m1(19,2)='������������'
m1(20,2)='������������'

m1(1,3)=''
m1(2,3)='���'
m1(3,3)='������'
m1(4,3)='������'
m1(5,3)='���������'
m1(6,3)='�������'
m1(7,3)='��������'
m1(8,3)='�������'
m1(9,3)='���������'
m1(10,3)='���������'
m1(11,3)=''
m1(12,3)=''
m1(13,3)=''
m1(14,3)=''
m1(15,3)=''
m1(16,3)=''
m1(17,3)=''
m1(18,3)=''
m1(19,3)=''
m1(20,3)=''

m1(1,4)=''
m1(2,4)='����'
m1(3,4)='���'
m1(4,4)='���'
m1(5,4)='������'
m1(6,4)='����'
m1(7,4)='�����'
m1(8,4)='����'
m1(9,4)='������'
m1(10,4)='������'
m1(11,4)=''
m1(12,4)=''
m1(13,4)=''
m1(14,4)=''
m1(15,4)=''
m1(16,4)=''
m1(17,4)=''
m1(18,4)=''
m1(19,4)=''
m1(20,4)=''

m2(1,1)='�������'
m2(2,1)='�����'
m2(3,1)='������'
m2(4,1)='�������'
m2(5,1)='��������'
m2(6,1)='��������'

m2(1,2)='�������'
m2(2,2)='�����'
m2(3,2)='������'
m2(4,2)='��������'
m2(5,2)='���������'
m2(6,2)='���������'

m2(1,3)='������'
m2(2,3)='������'
m2(3,3)='�����'
m2(4,3)='���������'
m2(5,3)='����������'
m2(6,3)='����������'

**  ��������!!!
for _i=1 to 6
   _tri=_c-1000*int(_c/1000)
   _c=int(_c/1000)
   _dva=_tri-100*int(_tri/100)
   if _dva>9.and._dva<20
      _sl2=m1(1+_dva,y(3*_i-1))+chr(32)+m2(_i,3)+chr(32)
      _sl1=m1(1+int(_tri/100),y(3*_i))+chr(32)
      _p=ltrim(_sl1)+ltrim(_sl2)+ltrim(_p)
   else
      _odin=_tri-10*int(_tri/10)
      _sl3=m1(1+_odin,y(3*_i-2))+chr(32)
      _sl2=m1(1+int(_dva/10),y(3*_i-1))+chr(32)
      _sl1=m1(1+int(_tri/100),y(3*_i))+chr(32)
      if len(alltrim(_sl3+_sl2+_sl1))>0
         _p=ltrim(_sl1)+ltrim(_sl2)+ltrim(_sl3);
            +m2(_i,iif(_odin=1,1,iif(_odin>1.and._odin<5,2,3)))+chr(32)+ltrim(_p)
      else
         if _i=2
            _p=m2(2,3)+chr(32)+ltrim(_p)
         endif
      endif
   endif
   if _c=0
      exit
   endif
endfor
*****wait wind _p
* set talk on
RETURN _p

FUNCTION  Y
PARAMETERS  _a
RETURN iif(_a=1 or _a=7,4,_a-3*int((_a-1)/3))


FUNCTION chk_nkms
 PARAMETERS pnkms
 n1  = val( subs( pnkms, 1, 1 ) )
 n2  = val( subs( pnkms, 2, 1 ) )
 n3  = val( subs( pnkms, 3, 1 ) )
 n4  = val( subs( pnkms, 4, 1 ) )
 n5  = val( subs( pnkms, 5, 1 ) )
 n6  = val( subs( pnkms, 6, 1 ) )
 n7  = val( subs( pnkms, 7, 1 ) )
 n8  = val( subs( pnkms, 8, 1 ) )
 n9  = val( subs( pnkms, 9, 1 ) )
 n10 = val( subs( pnkms,10, 1 ) )
 r1 = ( n2*2 + n3*8 + n4*6 + n5*3 + n6*5 + n7*4 + n8*2 + n9*1 + n10*7 ) % 10
RETURN iif(n1 = r1, .t., .f.)

FUNCTION chk_nenp
PARAMETERS lnENP

DO CASE 
	CASE VARTYPE(lnENP) = 'N'
		lcENP = STRTRAN(STR(lnENP,16,0),' ','0')
	CASE VARTYPE(lnENP) = 'C'
		lcENP = PADL(ALLTRIM(lnENP),16,'0')
	OTHERWISE
		lcENP = STRTRAN(SPACE(16),' ','0')
ENDCASE

VLast = RIGHT(lcENP,1)
V2 = LEFT(lcENP,LEN(lcENP)-1)

VA1 = ''
VB1 = ''

lMod = MOD(LEN(V2),2)

IF lMod = 1
	lModa = 0
	lModb = -1
ELSE
	lModb = 0
	lModa = -1
ENDIF

FOR NPos = LEN(V2) + lModa TO 1 STEP -2
	VA1 = VA1 + SUBSTR(V2,NPos,1)
ENDFOR

FOR NPos = LEN(V2) + lModb TO 1 STEP -2
	VB1 = VB1 + SUBSTR(V2,NPos,1)
ENDFOR

VA = ALLTRIM(STR(INT(VAL(VA1)) * 2))
VBA = VB1 + VA

VC = 0
FOR NPos = 1 TO LEN(VBA)
	VC = VC + VAL(SUBSTR(VBA,NPos,1))
ENDFOR

VD = (CEILING(VC/10) * 10) - VC

IF NOT VD = VAL(VLast)
	RETURN (.T.)
ELSE
	RETURN (.F.)
ENDIF


FUNCTION  pause
 PARAMETERS  arg
 wer = seco()
 DO  while seco() - wer <= arg
 ENDDO  
RETURN 

FUNCTION FLS(para1, para2, para3)

 PRIVATE m.dd_sid, m.dt_d, m.r_up, m.gd_sid, m.oms, m.s_all, m.m_value, m.m_unit, m.v_value, m.v_unit,;
  m.edizm, m.tarif, m.en

 m.dd_sid = para1 && sid
 m.dt_d   = para2 && dt_d, �������� (�������) ���� � �������� ����������!
 m.r_up   = para3 && ALLTRIM(r_up) �������� ��������

 IF !USED('medx')
  RETURN 0
 ENDIF 
 IF !USED('medpack')
  RETURN 0
 ENDIF 
 IF !USED('tarion')
  RETURN 0
 ENDIF 
    
 m.s_all = 0

 IF EMPTY(m.r_up)
  RETURN m.s_all
 ENDIF 
     
 m.gd_sid  = IIF(SEEK(m.dd_sid, 'medx'), ALLTRIM(medx.gd_sid), '')
 IF EMPTY(m.gd_sid)
  RETURN m.s_all
 ENDIF 
 
 m.m_value = IIF(SEEK(m.dd_sid, 'medx'), medx.mass_value, 0) && strength_mass_value form medicamet - ������� ����������!
 m.m_unit  = IIF(SEEK(m.dd_sid, 'medx'), ALLTRIM(medx.mass_unit), '')
 m.v_value = IIF(SEEK(m.dd_sid, 'medx'), medx.vol_value, 0)
 m.v_unit  = IIF(SEEK(m.dd_sid, 'medx'), ALLTRIM(medx.vol_unit), '')
     
 IF EMPTY(m.v_unit) AND EMPTY(m_unit)
  RETURN m.s_all
 ENDIF 
     
 m.edizm   = IIF(!EMPTY(m.v_unit), '��', '��')
 *IF m.gd_sid='GD012768'
 * m.edizm   = IIF(SEEK(m.r_up, 'medpack'), IIF(!EMPTY(medpack.vol_unit), '��', IIF(!EMPTY(medpack.mass_unit), '��', '')), '')
 *ENDIF 
 
 IF EMPTY(m.edizm)
  RETURN m.s_all
 ENDIF 
     
 IF m.edizm = '��'

   m.tarif  = IIF(!EMPTY(m.gd_sid) AND SEEK(m.gd_sid, 'tarion', 'cod'), tarion.ston, 0)
   m.en     = IIF(SEEK(m.r_up, 'medpack'), medpack.vol_value, 0) && ��������, 4 ��
   IF tarion.pr_v_value>0 AND tarion.pr_v_value <> m.en
    SKIP IN tarion 
    IF tarion.pr_v_value>0 AND tarion.pr_v_value = m.en
     m.tarif  = tarion.ston
    ELSE 
     m.tarif  = 0
    ENDIF 
   ENDIF 
   IF m.en>0
    *m.s_all  = (CEILING(m.dt_d/m.en) * m.en) * m.tarif
    *m.s_all  = ROUND((CEILING(m.dt_d/m.en) * m.en) * m.tarif,2)
    m.s_all  = CEILING(m.dt_d/m.en) * ROUND(m.en * m.tarif,2)
   ENDIF 

  ELSE && ��

   m.tarif  = IIF(!EMPTY(m.gd_sid) AND SEEK(m.gd_sid, 'tarion', 'cod'), tarion.ston, 0)
   m.en     = IIF(SEEK(m.r_up, 'medpack'), IIF(medpack.mass_value>0, medpack.mass_value, 1), 0) && ��������, 440 ��
   IF m.en>0
    *m.s_all  = (CEILING(m.dt_d)*m.en) * m.tarif
    m.s_all  = ROUND((CEILING(m.dt_d)*m.en) * m.tarif,2)
   ENDIF 

  ENDIF 
   
RETURN m.s_all

&& ������ �� 12 ������� 2018 ����
FUNCTION FSumm(Usl, STip, Kol, IsVed, pdatt)
 IF PARAMETERS()<5
  m.pdatt=m.tdat1
 ENDIF 

 PRIVATE m.sstip
 m.sstip = STip
 
 IF BETWEEN(Usl, 97107, 97999)
  IF EMPTY(m.sstip)
   m.sstip = '0'
  ENDIF 
 ENDIF 

 m.kulimit = IIF(m.tdat1<{01.01.2016}, 30, 9999)

 IF SEEK(Usl, 'Tarif')
   m.n_kd     = Tarif.n_kd
   m.price    = IIF(IsVed==.F., Tarif.Tarif, Tarif.Tarif_V)
   m.doplata  = IIF(FIELD('doplata', 'tarif')='DOPLATA', Tarif.Doplata, 0)
   m.kd_price = IIF(IsVed==.F., Tarif.stkd, Tarif.stkdv)
   *m.n_kd     = IIF(INLIST(INT(Usl/1000),92,192), Tarif.tarif/Tarif.stkd, m.n_kd)
   
   *IF !EMPTY(Tarif.n_kd)
   IF !EMPTY(m.n_kd)
    IF !EMPTY(m.sstip)
     DO CASE
      CASE INT(Usl/1000) = 83
        IF Usl!=83050
         DO CASE
          CASE kol < m.n_kd
           summa = round(m.kd_price * kol,2)
          OTHERWISE 
           summa = m.price
         ENDCASE  
        ELSE
         DO CASE
          CASE kol < m.n_kd
           summa = round(m.kd_price * kol,2)
          CASE kol = m.n_kd
           summa = m.price
          CASE kol > m.n_kd and kol <= m.kulimit
           summa = round(m.kd_price * kol,2)
          CASE kol > m.n_kd and kol > m.kulimit
           summa = round(m.kd_price * m.kulimit,2)
         ENDCASE  
        ENDIF 

      CASE INT(Usl/1000) = 183
       summa = iif(Kol<=m.kulimit, ROUND(Kol*m.kd_price,2), ROUND(m.kulimit*m.kd_price,2))

      CASE INT(Usl/1000) = 200
       summa = m.price + m.doplata
      
      OTHERWISE 
      
       IF INLIST(m.sstip, '0', 'v', '�', 'A') OR (INLIST(m.sstip, 'T', '�') AND kol>=10) OR ;
       	(INLIST(m.sstip, 'R') AND kol>=14)
        summa = m.price
       ELSE 
        summa = iif(Kol<m.n_kd, round(m.kd_price * Kol,2), m.price)
       ENDIF 

     ENDCASE
     
    ELSE 
     summa = 0
    ENDIF && !EMPTY(m.sstip)
   ELSE 
    IF INT(Usl/1000) <> 297
     summa = Kol * m.price
    ELSE 
     summa = m.price
    ENDIF 
   ENDIF && !EMPTY(Tarif.n_kd)
 ELSE 
  summa = 0 
 ENDIF && SEEK(Usl, 'Tarif')
RETURN Summa
&& ������ �� 12 ������� 2018 ����

FUNCTION FSummOld(Usl, STip, Kol, IsVed, pdatt)
 IF PARAMETERS()<5
  m.pdatt=m.tdat1
 ENDIF 
 m.kulimit = IIF(m.tdat1<{01.01.2016}, 30, 9999)
 IF SEEK(Usl, 'Tarif')
   m.n_kd    = Tarif.n_kd
   m.price = IIF(IsVed==.F., Tarif.Tarif, Tarif.Tarif_V)
   m.doplata = Tarif.Doplata
   m.kd_price = IIF(IsVed==.F., Tarif.stkd, Tarif.stkdv)
   IF !EMPTY(Tarif.n_kd)
    IF !EMPTY(STip)
     DO CASE
      CASE INT(Usl/1000) = 83
       IF m.pdatt<{01.09.2017}
        DO CASE
         CASE kol < m.n_kd
          summa = round(m.kd_price * kol,2)
         CASE kol = m.n_kd
          summa = m.price
         CASE kol > m.n_kd and kol <= m.kulimit
          summa = round(m.kd_price * kol,2)
         CASE kol > m.n_kd and kol > m.kulimit
          summa = round(m.kd_price * m.kulimit,2)
        ENDCASE  
       ELSE && m.pdatt>={01.09.2017}
        IF Usl!=83050
         DO CASE
          CASE kol < m.n_kd
           summa = round(m.kd_price * kol,2)
          OTHERWISE 
           summa = m.price
         ENDCASE  
        ELSE
         DO CASE
          CASE kol < m.n_kd
           summa = round(m.kd_price * kol,2)
          CASE kol = m.n_kd
           summa = m.price
          CASE kol > m.n_kd and kol <= m.kulimit
           summa = round(m.kd_price * kol,2)
          CASE kol > m.n_kd and kol > m.kulimit
           summa = round(m.kd_price * m.kulimit,2)
         ENDCASE  
        ENDIF 
       ENDIF 

      CASE INT(Usl/1000) = 183
       summa = iif(Kol<=m.kulimit, ROUND(Kol*m.kd_price,2), ROUND(m.kulimit*m.kd_price,2))

      CASE INT(Usl/1000) = 200
       summa = m.price + m.doplata

      OTHERWISE 
       IF m.pdatt>{01.07.2014}
*       IF INLIST(STip,'�','�') �������� 03.12.2014 !
       IF INLIST(STip,'8','�','�')
        summa = m.price
       ELSE
        IF m.pdatt>={01.01.2015}
         IF STip='0'
          summa = m.price
         ELSE 
          summa = iif(Kol<m.n_kd, round(m.kd_price * Kol,2), m.price)
         ENDIF 
        ELSE 
         summa = iif(Kol<m.n_kd, round(m.kd_price * Kol,2), m.price)
        ENDIF 
       ENDIF 
       ELSE 
*       IF INLIST(STip,'�','�') �������� 03.12.2014 !
       IF INLIST(STip,'�','�')
        summa = m.price
       ELSE
        IF m.pdatt>={01.01.2015}
         IF STip='0'
          summa = m.price
         ELSE 
          summa = iif(Kol<m.n_kd, round(m.kd_price * Kol,2), m.price)
         ENDIF 
        ELSE 
         summa = iif(Kol<m.n_kd, round(m.kd_price * Kol,2), m.price)
        ENDIF 
       ENDIF 
       ENDIF 
     ENDCASE
     
    ELSE 
     summa = 0
*     summa = m.price
    ENDIF && !EMPTY(STip)
   ELSE 
    summa = Kol * m.price
   ENDIF && !EMPTY(Tarif.n_kd)
 ELSE 
  summa = 0 
 ENDIF && SEEK(Usl, 'Tarif')
RETURN Summa

FUNCTION FSummVeryOld(Usl, STip, Kol, IsVed, pdatt)
 IF PARAMETERS()<5
  m.pdatt=m.tdat1
 ENDIF 
 m.kulimit = IIF(m.tdat1<{01.01.2016}, 30, 9999)
 IF SEEK(Usl, 'Tarif')
   m.n_kd    = Tarif.n_kd
   m.price = IIF(IsVed==.F., Tarif.Tarif, Tarif.Tarif_V)
   m.kd_price = IIF(IsVed==.F., Tarif.stkd, Tarif.stkdv)
   IF !EMPTY(Tarif.n_kd)
    IF !EMPTY(STip)
     DO CASE
      CASE INT(Usl/1000) = 83
       DO CASE
        CASE kol < m.n_kd
         summa = round(m.kd_price * kol,2)
        CASE kol = m.n_kd
         summa = m.price
        CASE kol > m.n_kd and kol <= m.kulimit
         summa = round(m.kd_price * kol,2)
        CASE kol > m.n_kd and kol > m.kulimit
         summa = round(m.kd_price * m.kulimit,2)
       ENDCASE  

      CASE INT(Usl/1000) = 183
       summa = iif(Kol<=m.kulimit, ROUND(Kol*m.kd_price,2), ROUND(m.kulimit*m.kd_price,2))

      CASE INT(Usl/1000) = 200
       summa = m.price

      OTHERWISE 
       IF m.pdatt>{01.07.2014}
*       IF INLIST(STip,'�','�') �������� 03.12.2014 !
       IF INLIST(STip,'8','�','�')
        summa = m.price
       ELSE
        IF m.pdatt>={01.01.2015}
         IF STip='0'
          summa = m.price
         ELSE 
          summa = iif(Kol<m.n_kd, round(m.kd_price * Kol,2), m.price)
         ENDIF 
        ELSE 
         summa = iif(Kol<m.n_kd, round(m.kd_price * Kol,2), m.price)
        ENDIF 
       ENDIF 
       ELSE 
*       IF INLIST(STip,'�','�') �������� 03.12.2014 !
       IF INLIST(STip,'�','�')
        summa = m.price
       ELSE
        IF m.pdatt>={01.01.2015}
         IF STip='0'
          summa = m.price
         ELSE 
          summa = iif(Kol<m.n_kd, round(m.kd_price * Kol,2), m.price)
         ENDIF 
        ELSE 
         summa = iif(Kol<m.n_kd, round(m.kd_price * Kol,2), m.price)
        ENDIF 
       ENDIF 
       ENDIF 
     ENDCASE
     
    ELSE 
     summa = 0
*     summa = m.price
    ENDIF && !EMPTY(STip)
   ELSE 
    summa = Kol * m.price
   ENDIF && !EMPTY(Tarif.n_kd)
 ELSE 
  summa = 0 
 ENDIF && SEEK(Usl, 'Tarif')
RETURN Summa

FUNCTION  DToH
 PARAMETERS  iDig
 oDig = Replicate(' ',8)
 x = Mod(iDig,16)
 i = 8
 oDig = stuff(oDig,i,1,TransA(x))
 Do While iDig > 0
  oDig = stuff(oDig,i,1,TransA(x))
  i = i-1
  iDig = Int(iDig/16)
  x = Mod(iDig,16)
 EndD
 oDig = iif(iDig>0, stuff(oDig,i,1,TransA(iDig)), oDig)
RETURN allt(oDig)

FUNCTION  TransA
PARAMETERS  Pt1
Do Case
 Case Pt1=0
  Pt2 = '0'
 Case Pt1=1
  Pt2 = '1'
 Case Pt1=2
  Pt2 = '2'
 Case Pt1=3
  Pt2 = '3'
 Case Pt1=4
  Pt2 = '4'
 Case Pt1=5
  Pt2 = '5'
 Case Pt1=6
  Pt2 = '6'
 Case Pt1=7
  Pt2 = '7'
 Case Pt1=8
  Pt2 = '8'
 Case Pt1=9
  Pt2 = '9'
 Case Pt1=10
  Pt2 = 'A'
 Case Pt1=11
  Pt2 = 'B'
 Case Pt1=12
  Pt2 = 'C'
 Case Pt1=13
  Pt2 = 'D'
 Case Pt1=14
  Pt2 = 'E'
 Case Pt1=15
  Pt2 = 'F'
EndC
RETURN  Pt2

FUNCTION  HToD
 PARAMETERS  iDig
 iDig = Uppe(allt(strt(iDig,' ')))
 dLen = len(iDig)
 oDig = 0
 For i=1 to dLen-1
  oDig = (oDig+TransT(Subs(iDig,i,1)))*16
 EndF
 oDig = oDig+TransT(Subs(iDig,i,1))
RETURN oDig

FUNCTION TransT
PARAMETERS Pt1
Do Case
 Case Pt1='0'
  Pt2 = 0
 Case Pt1='1'
  Pt2 = 1
 Case Pt1='2'
  Pt2 = 2
 Case Pt1='3'
  Pt2 = 3
 Case Pt1='4'
  Pt2 = 4
 Case Pt1='5'
  Pt2 = 5
 Case Pt1='6'
  Pt2 = 6
 Case Pt1='7'
  Pt2 = 7
 Case Pt1='8'
  Pt2 = 8
 Case Pt1='9'
  Pt2 = 9
 Case Pt1='A'
  Pt2 = 10
 Case Pt1='B'
  Pt2 = 11
 Case Pt1='C'
  Pt2 = 12
 Case Pt1='D'
  Pt2 = 13
 Case Pt1='E'
  Pt2 = 14
 Case Pt1='F'
  Pt2 = 15
EndC
RETURN Pt2

FUNCTION TransPol
LPARAMETERS tPolis, nVozvrat
tPolis = ALLTRIM(tPolis)
nVozvrat = IIF(EMPTY(nVozvrat), 1, nVozvrat)
DO CASE  
	CASE LEFT(tPolis,2) = '77'
 		DO CASE 
 			CASE INLIST(Upper(EngToRus(SUBSTR(tPolis,5,1))), '�', '�', '�')
	    		Pps = Upper(EngToRus(Left(tPolis,5)))
	    		Ppn = Int(Val(Subs(tPolis,6)))
	    		IF BETWEEN(ppn,1,999999)
	    		 tPolis = Pps + ' ' +Padl(Ppn,6,'0')
	    		ELSE 
	  	 		 tPolis = IIF(nVozvrat = 2, '', tPolis)
	    		ENDIF 

	  		CASE  BETWEEN(VAL(SUBSTR(tPolis,5,2)),0,27) OR ;
	  		 INLIST(SUBSTR(tPolis,5,2),'45','50','51','52','73','77','99')
	    		Pps   = LEFT(tPolis,6)
	    		Ppn   = INT(VAL(SUBSTR(tPolis,7)))
	    		IF BETWEEN(ppn, 1, 9999999999)
	    		 tPolis = Pps + ' ' +Padl(Ppn,10,'0')
	    		ELSE
	  	 		 tPolis = IIF(nVozvrat = 2, '', tPolis)
	    		ENDIF 
   
	  		OTHERWISE 
	  	 		tPolis = IIF(nVozvrat = 2, '', tPolis)
	  	 	    
  		ENDCASE 
      
	 CASE InList(Upper(RusToEng(Subs(tPolis,1,2))), 'I1', 'M1', 'M4', 'M6', 'R2', 'S2', 'S5', 'V2', 'P2','R4')
	 	tPolis = STRTRAN(ALLTRIM(tPolis)," ","")
	 	Pps = Uppe(RusToEng(Left(tPolis,5)))
	 	Ppn = Int(Val(Subs(tPolis,6)))
	 	tPolis = Pps + ' ' +Padl(Ppn,5,'0')
 
	 OTHERWISE 
		tPolis = IIF(nVozvrat = 2, '', tPolis)

ENDCASE 

RETURN PADR(ALLTRIM(tPolis),17)


FUNCTION SV
	PARAMETERS ppp,x,y,z
RETURN 

FUNCTION NV
PARAMETERS ppp,x,y,z

DO CASE 

 CASE EMPTY(ppp) && ������ �����
  RETURN 0

 CASE SUBSTR(ppp,1,3)='46-' && ��������� �����
  IF !INLIST(SUBSTR(ppp,4,3),'01 ','02 ','03 ','04 ','05 ','06 ','07 ','08 ','09 ','10 ','11 ','12 ','13 ','14 ','15 ','16 ','17 ','18 ','19 ','20 ','21 ','22 ','23 ','24 ')
   RETURN 1
  ENDIF   

  IF !BETWEEN(VAL(SUBSTR(ppp,7)),1,999999)
   =MESSAGEBOX("���-�� �� �� � ��������� �������!",0+48,"��������!")
   RETURN 2
  ENDIF 


 CASE SUBSTR(ppp,1,2) = '77'  && ���������� �����
  DO CASE 
   CASE INLIST(SUBSTR(ppp,5,1),'�','�','�') AND ;
    INLIST(SUBSTR(ppp,3,2),'01','02','03','04','05','06','07','08','09','10') AND  ;
    BETWEEN(INT(VAL(SUBSTR(ppp,7))),1,999999)

   CASE ISDIGIT(SUBSTR(ppp,3,1)) AND ISDIGIT(SUBSTR(ppp,4,1)) AND ISDIGIT(SUBSTR(ppp,5,1)) AND  ;
    ISDIGIT(Subs(ppp,6,1)) AND BETWEEN(VAL(SUBSTR(ppp,3,2)),0,9) AND (BETWEEN(VAL(SUBSTR(ppp,5,2)),0,27) ;
    OR INLIST(SUBSTR(ppp,5,2),'45','50','51','52','73','77','99')) AND VAL(SUBSTR(ppp,8))>0

    IF !chk_nkms(PADL(ALLTRIM(SUBSTR(ppp,8)),10,'0')) 
     RETURN 3
    ENDIF 

   OTHERWISE 
    RETURN 4
   
  ENDCASE 

 CASE INLIST(SUBSTR(ppp,1,2), 'I1', 'M1', 'M4', 'M6', 'R2', 'S2','S5', 'P2' ) && ���� �����������
  DO CASE     
   CASE  UPPER(RusToEng(SUBSTR(ppp,1,2)))='I1' AND  BETWEEN(VAL(SUBSTR(ppp,3,3)),151,199)
   CASE  UPPER(RusToEng(SUBSTR(ppp,1,2)))='M1' AND  BETWEEN(VAL(SUBSTR(ppp,3,3)),200,399)
   CASE  UPPER(RusToEng(SUBSTR(ppp,1,2)))='M4' AND  BETWEEN(VAL(SUBSTR(ppp,3,3)),401,499)
   CASE  UPPER(RusToEng(SUBSTR(ppp,1,2)))='R4' AND  BETWEEN(VAL(SUBSTR(ppp,3,3)),051,075)
   CASE  UPPER(RusToEng(SUBSTR(ppp,1,2)))='P2' AND  BETWEEN(VAL(SUBSTR(ppp,3,3)),101,150)
   CASE  UPPER(RusToEng(SUBSTR(ppp,1,2)))='R2' AND  BETWEEN(VAL(SUBSTR(ppp,3,3)),501,699)
   CASE  UPPER(RusToEng(SUBSTR(ppp,1,2)))='S2' AND  BETWEEN(VAL(SUBSTR(ppp,3,3)),701,799)
   CASE  UPPER(RusToEng(SUBSTR(ppp,1,2)))='S5' AND  (BETWEEN(VAL(SUBSTR(ppp,3,3)),801,899) OR ;
		BETWEEN(VAL(SUBSTR(ppp,3,3)),951,980))
   OTHERWISE 
    RETURN 5
  ENDCASE 

  IF !BETWEEN(VAL(SUBSTR(ppp,7)), 1, 99999)
   RETURN 6
  ENDIF 
             
 OTHERWISE && ������ ����������

ENDCASE 

RETURN 10 && Ok!


FUNCTION NULLIF(LnSRC1)
	RETURN ICASE(VARTYPE(m.LnSRC1) ="X", 'NULL', EMPTY(m.LnSRC1), ;
		'NULL', VARTYPE(m.LnSRC1)=="C", ['] + ALLTRIM(m.LnSRC1) + ['], ;
		VARTYPE(m.LnSRC1)== "D", ['] + DTOC(m.LnSRC1) + ['], ;
		VARTYPE(m.LnSRC1)== "T", ['] + TTOC(m.LnSRC1) + ['], ;
		VARTYPE(m.LnSRC1)= "N" AND !EMPTY(m.LnSRC1), ALLTRIM(STR(m.LnSRC1)), 'NULL')
ENDFUNC

FUNCTION OnShutDown()
	IF VARTYPE(oApp) == "O"
		oApp.Exit_App()
	ELSE
		ON ERROR
		ON SHUTDOWN
		CLEAR EVENTS
	ENDIF
ENDFUNC

FUNCTION NEWOBJ
	LPARAMETERS tcClass, tuParm1, tuParm2, ;
		tuParm3, tuParm4
	LOCAL lcClass, ;
		lcLibrary, ;
		lnPos, ;
		loObject

	lcClass   = UPPER(tcClass)
	lcLibrary = ''
	IF ',' $ lcClass AND LEFT(lcClass, 1) <> ','
		lnPos     = AT(',', lcClass)
		lcLibrary = ALLTRIM(LEFT(lcClass, lnPos - 1))
		lcClass   = ALLTRIM(SUBSTR(lcClass, lnPos + 1))
		IF lcLibrary $ UPPER(SET('CLASSLIB'))
			lcLibrary = ''
		ELSE
			SET CLASSLIB TO (lcLibrary) ADDITIVE
		ENDIF
	ENDIF

	DO CASE
		CASE pcount() = 1
			loObject = CREATEOBJECT(lcClass)
		CASE pcount() = 2
			loObject = CREATEOBJECT(lcClass, @tuParm1)
		CASE pcount() = 3
			loObject = CREATEOBJECT(lcClass, @tuParm1, @tuParm2)
		CASE pcount() = 4
			loObject = CREATEOBJECT(lcClass, @tuParm1, @tuParm2, @tuParm3)
		CASE pcount() = 5
			loObject = CREATEOBJECT(lcClass, @tuParm1, @tuParm2, @tuParm3, ;
				@tuParm4)
	ENDCASE
	RETURN loObject
ENDFUNC

FUNCTION SetDSession
	*-- ����� ��������� ����� ���������� ���������
	SET DELETED ON 
	SET ANSI ON 
	SET CENTURY ON 
	SET CENTURY TO 19 ROLLOVER 10
	SET CONFIRM OFF 
	SET DATE TO GERMAN 
	SET EXACT OFF 
	SET EXCLUSIVE ON 
	SET HOURS TO 24
	SET LOCK OFF 
	SET MARK TO 
	SET MEMOWIDTH TO 120
	SET MULTILOCKS OFF
	SET NEAR ON 
	SET NULL OFF 
	SET POINT TO "."
	SET REPROCESS TO 10 SECONDS
	SET SAFETY OFF 
	SET SEPARATOR TO "`"
	SET SYSFORMATS OFF 
	SET TALK OFF 

	IF 3=2
	SET TALK OFF
	SET NOTIFY OFF
	SET COMPATIBLE OFF

	SET SYSFORMATS OFF
	SET CURRENCY TO '�'
	SET CURRENCY RIGHT
	SET CENTURY ON
	SET DATE TO GERMAN
	SET DECIMALS TO 2
	SET HOURS TO 24
	SET POINT TO "."
	SET SEPARATOR TO "`"
	SET FDOW  TO 2
	SET FWEEK TO 1

	=CURSORSETPROP("Buffering", 1, 0)
	SET MEMOWIDTH TO 120
	SET FDOW TO 2
	SET ANSI ON
	SET SAFETY OFF
	SET MEMOWIDTH TO 80
	SET MULTILOCKS OFF
	SET EXCLUSIVE ON
	SET BELL OFF
	* SET NEAR OFF
	SET EXACT OFF
	SET EXACT ON
	SET INTENSITY OFF
	SET CONFIRM ON
	SET LOCK OFF
	SET REPROCESS TO 10 SECONDS
	SET NULL ON
	SET NULLDISPLAY TO " "
	CLEAR MACROS
	#IF DEBUGMODE
		SET RESOURCE TO FOXUSER.DBF
		SET RESOURCE ON
		SET DEBUG ON
		SET ESCAPE ON
	#ELSE
		SET RESOURCE OFF
		SET DEBUG OFF
		SET ESCAPE OFF
	#ENDIF
	SET ASSERTS OFF
	SET CPCOMPILE TO 1251
	ENDIF 

	RETURN .T.
ENDFUNC

FUNCTION ModifyColor
	*	��������� �����
	LPARAMETERS lnColor, lnDelta, lnIndex
	LOCAL ARRAY S_Rgb[3]
	*	���� ��� ��������� 0 - 16`777`216
	lnColor	= MIN( 16777216, MAX( 0, iif( VARTYPE( lnColor ) == 'N', lnColor, 0 )))
	*	��� ���� �������� 0 - ��� �����, 1-�������, 2-�������, 3-�����
	lnIndex	= MIN( 3, MAX( 0, iif( VARTYPE( lnIndex ) == 'N', lnIndex, 0 )))
	*	�� ������� ���� �������� ����
	lnDelta	= IIF( VARTYPE( lnDelta ) == 'N', lnDelta , 0 )

	S_Rgb[3]	= INT( lnColor/ 65536 )	&& �����
	lnColor		= lnColor- ( S_Rgb[3] * 65536 )
	S_Rgb[2]	= INT( lnColor/ 256 )	&& �������
	lnColor		= lnColor- ( S_Rgb[2] * 256 )
	S_Rgb[1]	= lnColor				&& �������

	IF lnIndex = 0	&& ��� �����
		RETURN ( rgb(;
			MAX( 0, MIN( 255,  S_Rgb[ 1 ] + lnDelta )),;
			MAX( 0, MIN( 255,  S_Rgb[ 2 ] + lnDelta )),;
			MAX( 0, MIN( 255,  S_Rgb[ 3 ] + lnDelta ))))
	ELSE			&& ���� ����
		RETURN ( MAX( 0, MIN( 255,  S_Rgb[ lnIndex ] + lnDelta )))
	ENDIF
ENDFUNC


*-- ��������� �������� ���������� ��� �������� ���������� � �������������� �� � ������
FUNCTION Revers
	LPARAMETERS luStr, lnPrecision, lnScale
	DO CASE
		CASE VARTYPE(m.luStr) = "T"
			RETURN Revers(TTOC(m.luStr,1))
		CASE VARTYPE(m.luStr) == "C"
			m.luStr = CHRTRAN(NVL(m.luStr, ""), ["�], [�"])
			RETURN CHRTRAN(m.luStr, " !#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~�����������������������������������������������������������������", ;
				"�����������������������������������������������������������������~}|{zyxwvutsrqponmlkjihgfedcba`_^]\[ZYWVUTSRQPONMLKJIHGFEDCBA@?>=<;:9876543210/.-,+*)('&%$#! ")
		CASE VARTYPE(m.luStr) == "N"
			m.lnPrecision = IIF(VARTYPE(m.lnPrecision)=="N", m.lnPrecision, 18)
			m.lnScale = IIF(VARTYPE(m.lnScale)=="N", m.lnScale, 2)
			RETURN Revers( NumToStr( m.luStr, m.lnPrecision, m.lnScale ))
		CASE VARTYPE(m.luStr) = "D"
			m.luStr = NVL(m.luStr, {})
			RETURN  STR(99999999 - VAL(DTOS(m.luStr)))
		OTHERWISE
			RETURN ""
	ENDCASE
ENDFUNC

FUNCTION NumToStr
	LPARAMETERS luStr, lnPrecision, lnScale
	luStr = NVL(luStr, 0)
	lnPrecision = IIF(VARTYPE(lnPrecision)=="N", lnPrecision, 18)
	lnScale = IIF(VARTYPE(lnScale)=="N", lnScale, 2)
	RETURN IIF(luStr < 0, 'A', 'Z') + ;
		IIF(luStr < 0, STR(-99999999999999.99 - luStr,lnPrecision,lnScale), STR(luStr,lnPrecision,lnScale))
ENDFUNC

FUNCTION SaveToINIFile
	LPARAMETERS lcName, lcRasdel, lcValue, lcIniFile
	*-- ������ ������ ��������� � INI - ����
	LOCAL lcEntry, lPath
	m.lcIniFile = IIF(VARTYPE(m.lcIniFile) == "C", m.lcIniFile, INIFILE)
	m.lPath = SYS(5)+SYS(2003) + '\'
*	=WritePrivStr(m.lcRasdel, m.lcName, m.lcValue, m.lPath + m.lcIniFile)
ENDFUNC

FUNCTION ReadFromINIFile
	LPARAMETERS lcName, lcRasdel, lcIniFile
	*-- ������ ������ ��������� �� INI - �����
	LOCAL  lcBuffer, lcRetValue, lPath
	m.lcIniFile = IIF(VARTYPE(m.lcIniFile) == "C", m.lcIniFile, INIFILE)
	m.lPath = SYS(5)+SYS(2003) + '\'
	m.lcBuffer = SPACE(100) + CHR(0)
	*-- ������ ������� ���� �� INI �����
	m.lcRetValue = ""
	TRY
*		IF GetPrivStr(m.lcRasdel, m.lcName, "", @m.lcBuffer, LEN(m.lcBuffer), ;
				m.lPath + m.lcIniFile) > 0
			m.lcRetValue = ALLTRIM(LEFT(m.lcBuffer, AT(CHR(0), m.lcBuffer)-1))
*		ENDIF
	CATCH
		m.lcRetValue = ""
	ENDTRY
	RETURN (m.lcRetValue)
ENDFUNC

FUNCTION FormIsObject()
	LPARAMETERS NameForm
	*-- ���������� .T. ���� �������� ����� ������ (���- "O") � �����
	*-- ������� ����� "Form".
	IF PCOUNT() = 0 OR EMPTY(NameForm)
		RETURN (TYPE("_SCREEN.ActiveForm") == "O" AND ;
			UPPER(_SCREEN.ACTIVEFORM.BASECLASS) = "FORM")
	ELSE
		RETURN (TYPE("_SCREEN.ActiveForm") == "O" AND ;
			UPPER(_SCREEN.ACTIVEFORM.BASECLASS) = "FORM" ;
			AND UPPER(NameForm) == UPPER(_SCREEN.ACTIVEFORM.NAME))
	ENDIF
ENDFUNC

  
FUNCTION cKbLayOut
  LPARAMETERS lnNeedKeyboard  

  #DEFINE KEYBOARD_GERMAN_ST 	0x0407		&& �������� (��������)  
  #DEFINE KEYBOARD_ENGLISH_US 	0x0409		&& ���������� (����������� �����)  
  #DEFINE KEYBOARD_FRENCH_ST 	0x040c		&& ����������� (��������)  
  #DEFINE KEYBOARD_RUSSIAN 		0x0419		&& �������  
    
 * ������ ������� ��������� ����������  
  DECLARE INTEGER GetKeyboardLayout IN Win32API Integer  
  LOCAL lnCurrentKeyboard  
  lnCurrentKeyboard = GetKeyboardLayout(0)  
 * ��������� ������� ����� (������� 16 ��� �� 32)  
  lnCurrentKeyboard = BitRShift(m.lnCurrentKeyboard,16)  
    
 * ���� ������� ��������� ������� �� ������  
  IF m.lnCurrentKeyboard <> m.lnNeedKeyboard  
 	* ������� �������� ������ �������������� ������ ���������  
  	DECLARE INTEGER ActivateKeyboardLayout IN Win32API Integer, Integer  
  	LOCAL lnNewKeyboard  
  	lnNewKeyboard = ActivateKeyboardLayout(m.lnNeedKeyboard,0)  
    
  	IF m.lnNewKeyboard=0  
 		* ������ ��������� ���������� �� ������� � ������ �������������� ���������  
 		* ��������� ������ ���������  
  		DECLARE INTEGER LoadKeyboardLayout IN Win32API String, Integer  
 		* ��� ������ ��������� ���� ���������������� � ������ ���� "00000419"  
 		* �.�. ������ �� 8 ��������, ��� ������ 4 - ��� ����, � ��������� 4 ��� � 16-������ �������  
  		LOCAL lcNeedKeyboard  
  		lcNeedKeyboard = RIGHT(TRANSFORM(m.lnNeedKeyboard,"@0"),8)  
 		* ���������� ��������  
  		m.lnNewKeyboard = LoadKeyboardLayout(m.lcNeedKeyboard,1)  
 		* ��������� ������� ����� (������� 16 ��� �� 32)  
  		m.lnNewKeyboard = BitRShift(m.lnNewKeyboard,16)  
  		IF m.lnNewKeyboard <> m.lnNeedKeyboard  
 			* ��������� ������ ��������� �� �������  
  		ENDIF  
  	ENDIF  
  ENDIF
RETURN

FUNCTION whatKb
  #DEFINE KEYBOARD_GERMAN_ST 	0x0407		&& �������� (��������)  
  #DEFINE KEYBOARD_ENGLISH_US 	0x0409		&& ���������� (����������� �����)  
  #DEFINE KEYBOARD_FRENCH_ST 	0x040c		&& ����������� (��������)  
  #DEFINE KEYBOARD_RUSSIAN 		0x0419		&& �������  

 * ������ ������� ��������� ����������  
  DECLARE INTEGER GetKeyboardLayout IN Win32API Integer  
  LOCAL lnCurrentKeyboard  
  lnCurrentKeyboard = GetKeyboardLayout(0)  
 * ��������� ������� ����� (������� 16 ��� �� 32)  
  lnCurrentKeyboard = BitRShift(m.lnCurrentKeyboard,16)  
    
RETURN lnCurrentKeyboard


FUNCTION IsKms(tcPolis)
 IF LEN(ALLTRIM(tcPolis))!=17
  RETURN .f. 
 ENDIF 
 IF SUBSTR(tcPolis,7,1)!=' '
  RETURN .f.
 ENDIF 
 IF LEFT(tcPolis,2)!='77'
  RETURN .f.
 ENDIF 
 IF !ISDIGIT(SUBSTR(tcPolis,3,1)) OR !ISDIGIT(SUBSTR(tcPolis,4,1)) OR ;
    !ISDIGIT(SUBSTR(tcPolis,5,1)) OR !ISDIGIT(SUBSTR(tcPolis,6,1))
  RETURN .f.
 ENDIF 
 IF !BETWEEN(VAL(SUBSTR(tcPolis,3,2)),0,99)
  RETURN .f.
 ENDIF 
 IF !BETWEEN(VAL(SUBSTR(tcPolis,5,2)),0,27) AND !INLIST(VAL(SUBSTR(tcPolis,5,2)),45,50,51,52,73,77,99)
  RETURN .f.
 ENDIF 
 IF !ISDIGIT(SUBSTR(tcPolis,8,1)) OR !ISDIGIT(SUBSTR(tcPolis,9,1)) OR ;
    !ISDIGIT(SUBSTR(tcPolis,10,1)) OR !ISDIGIT(SUBSTR(tcPolis,11,1)) OR ;
    !ISDIGIT(SUBSTR(tcPolis,12,1)) OR !ISDIGIT(SUBSTR(tcPolis,13,1)) OR ;
    !ISDIGIT(SUBSTR(tcPolis,14,1)) OR !ISDIGIT(SUBSTR(tcPolis,15,1)) OR ;
    !ISDIGIT(SUBSTR(tcPolis,16,1)) OR !ISDIGIT(SUBSTR(tcPolis,17,1))
  RETURN .f.
 ENDIF 
 IF !chk_nkms(SUBSTR(tcPolis,8,10))  
  RETURN .f.
 ENDIF 
RETURN .t. 

*FUNCTION IsVs(tcPolis)
* IF LEN(ALLTRIM(tcPolis))!=15
*  RETURN .f. 
* ENDIF 
* IF SUBSTR(tcPolis,6,1)!=' '
*  RETURN .f.
* ENDIF 
* IF !INLIST(LEFT(tcPolis,2), 'I1','M1','M4','R2','S2','S5','R4','P2','I3','P3','S6','S7')
*  RETURN .f.
* ENDIF 
* IF !ISDIGIT(SUBSTR(tcPolis,3,1)) OR !ISDIGIT(SUBSTR(tcPolis,4,1)) OR !ISDIGIT(SUBSTR(tcPolis,5,1))
*  RETURN .f.
* ENDIF 
* IF !ISDIGIT(SUBSTR(tcPolis,7,1))  OR !ISDIGIT(SUBSTR(tcPolis,8,1))  OR ;
*    !ISDIGIT(SUBSTR(tcPolis,9,1))  OR !ISDIGIT(SUBSTR(tcPolis,10,1)) OR ;
*    !ISDIGIT(SUBSTR(tcPolis,11,1)) OR !ISDIGIT(SUBSTR(tcPolis,12,1)) OR ;
*    !ISDIGIT(SUBSTR(tcPolis,13,1)) OR !ISDIGIT(SUBSTR(tcPolis,14,1)) OR ;
*    !ISDIGIT(SUBSTR(tcPolis,15,1))
*  RETURN .f.
* ENDIF  
*RETURN .t. 

FUNCTION IsVs(tcPolis) && ������ �� 05.04.2020
 IF LEN(ALLTRIM(tcPolis))<>14
  RETURN .f. 
 ENDIF 
 IF !INLIST(LEFT(tcPolis,2), 'I3','M1','M4','R2','R4','R8','S5','S7')
  RETURN .f.
 ENDIF 
 IF !ISDIGIT(SUBSTR(tcPolis,3,1)) OR !ISDIGIT(SUBSTR(tcPolis,4,1)) OR !ISDIGIT(SUBSTR(tcPolis,5,1))
  RETURN .f.
 ENDIF 
 IF !ISDIGIT(SUBSTR(tcPolis,6,1))  OR !ISDIGIT(SUBSTR(tcPolis,7,1))  OR ;
    !ISDIGIT(SUBSTR(tcPolis,8,1))  OR !ISDIGIT(SUBSTR(tcPolis,9,1)) OR ;
    !ISDIGIT(SUBSTR(tcPolis,10,1)) OR !ISDIGIT(SUBSTR(tcPolis,11,1)) OR ;
    !ISDIGIT(SUBSTR(tcPolis,12,1)) OR !ISDIGIT(SUBSTR(tcPolis,13,1)) OR ;
    !ISDIGIT(SUBSTR(tcPolis,14,1))
  RETURN .f.
 ENDIF  
RETURN .t. 

FUNCTION IsVsN(tcPolis)
 IF LEN(ALLTRIM(tcPolis))!=9
  RETURN .f. 
 ENDIF 
 IF !ISDIGIT(SUBSTR(tcPolis,1,1))  OR !ISDIGIT(SUBSTR(tcPolis,2,1))  OR ;
    !ISDIGIT(SUBSTR(tcPolis,3,1))  OR !ISDIGIT(SUBSTR(tcPolis,4,1)) OR ;
    !ISDIGIT(SUBSTR(tcPolis,5,1)) OR !ISDIGIT(SUBSTR(tcPolis,6,1)) OR ;
    !ISDIGIT(SUBSTR(tcPolis,7,1)) OR !ISDIGIT(SUBSTR(tcPolis,8,1)) OR ;
    !ISDIGIT(SUBSTR(tcPolis,9,1))
  RETURN .f.
 ENDIF  
RETURN .t. 

FUNCTION IsENP(tcPolis)
 IF LEN(ALLTRIM(tcPolis))!=16
  RETURN .f. 
 ENDIF 
 IF !ISDIGIT(SUBSTR(tcPolis,1,1))  OR !ISDIGIT(SUBSTR(tcPolis,2,1))  OR ;
    !ISDIGIT(SUBSTR(tcPolis,3,1))  OR !ISDIGIT(SUBSTR(tcPolis,4,1))  OR ;
    !ISDIGIT(SUBSTR(tcPolis,5,1))  OR !ISDIGIT(SUBSTR(tcPolis,6,1))  OR ;
    !ISDIGIT(SUBSTR(tcPolis,7,1))  OR !ISDIGIT(SUBSTR(tcPolis,8,1))  OR ;
    !ISDIGIT(SUBSTR(tcPolis,9,1))  OR !ISDIGIT(SUBSTR(tcPolis,10,1)) OR ;
    !ISDIGIT(SUBSTR(tcPolis,11,1)) OR !ISDIGIT(SUBSTR(tcPolis,12,1)) OR ;
    !ISDIGIT(SUBSTR(tcPolis,13,1)) OR !ISDIGIT(SUBSTR(tcPolis,14,1)) OR ;
    !ISDIGIT(SUBSTR(tcPolis,15,1)) OR !ISDIGIT(SUBSTR(tcPolis,16,1))
  RETURN .f.
 ENDIF  
* IF chk_nenp(SUBSTR(tcPolis,7,10))
*  RETURN .f.
* ENDIF 
RETURN .t.

Func SndWait
 Para BFile, EFile, DFile, Regim
 Regim = iif(Empty(Regim),0,1)
 Push Key Clea
 Set Escape On
 On Escape Do StopScan
 StopIt = .t.

 Do While File('&PAisOms\&User\OutPut\&BFile') And StopIt
  Wait iif(Regim=0, "�������� �������� ��...", "�������� ��...") Wind Nowa  
 EndD

 On  Escape
 Set Escape Off
 Pop Key
 
  Do Case 
   Case !StopIt
    dele file &PAisOms\&User\OutPut\&BFile
    dele file &PAisOms\&User\OutPut\&DFile
    wait " �������� "+iif(Regim=0,"�������� ",'')+"��������! " wind nowa
    retu .f.
   Othe
    Do Case
     case file('&PAisOms\&User\OutPut\&EFile')
      dele file &PAisOms\&User\OutPut\&EFile
      dele file &PAisOms\&User\OutPut\&DFile
      wait " ���������� ��������� ��! " wind nowa
      retu .f.
     case !file('&PAisOms\&User\OutPut\&DFile')
      retu .t.
     Other
      wait "������������� ��������!" wind nowa
      retu .f.
    EndC
  EndC

PROCEDURE  StopScan
 Flag = MESSAGEBOX("���������� �������� ������?",4+48,"��������!")
 IF Flag == 6
  StopIt = .f.
 ELSE
  StopIt =  .t.
 ENDIF
RETURN
 

PROCEDURE StopScan2
Flag = MESSAGEBOX("���������� �������� ������?",4+48,"��������!")
IF Flag == 6
 StopIt = .f.
ELSE
 StopIt = .t.
ENDIF
RETURN 

Func ScanDir
 Para par0, par1, par2, par3
 
 par0 = allt(par0)
 par1 = allt(par1)
 par2 = allt(par2)
 par3 = allt(par3)

 * par0 - ���������� ������������ 
 * par1 - �������� 1-�� �����
 * par2 - Subject  1-�� �����
 * par3 - Resent-Message-Id 1-�� �����

 ikl = .f.
    
  _name = sys(2000, par0+'\'+par1)
  poi = fopen(par0+'\'+_name)
  if poi != -1
   x = ''
   =fseek(poi,0) && �������  � ������ �����
   do while !feof(poi)
    x=allt(fgets(poi))
    if iif(!empty(par2), x = 'Subject' and allt(subs(x, at(':',x)+1)) = par2, 1=1)
     =fseek(poi,0) && �������  � ������ �����
     if !empty(par3)
      y = ''
      do while !feof(poi)
       y=allt(fgets(poi))
       if y = 'Resent-Message-Id' and allt(subs(y, at(':',y)+1)) = par3
        wait "��������� ���� " + lower(_name) + " � ��������� �����������!" wind time 1
        =fclos(poi)
        retu _name
       else
       endi
      endd
     else
      =fseek(poi,0) && �������  � ������ �����
      y    = ''
      rmid = ''
      do while !feof(poi)
       y = allt(fgets(poi))
       if y = 'Resent-Message-Id'
        rmid = allt(subs(y, at(':',y)+1))
       else
       endi
      endd
      if empty(rmid)
       wait "��������� ���� " + lower(_name) wind time 1
       =fclos(poi)
       retu _name
      Else
       wait "��������� ���� " + lower(_name) wind time 1
       =fclos(poi)
       retu _name
      endi
     endi
    else
    endi
   endd
   =fclos(poi)
  else
  endi

  _name = sys(2000, par0+'\'+par1, 1)
  poi = fopen(par0+'\'+_name)
  do while  poi != -1
   x = ''
   =fseek(poi,0) && �������  � ������ �����
   do while !feof(poi)
    x=allt(fgets(poi))
    if iif(!empty(par2), x = 'Subject' and allt(subs(x, at(':',x)+1)) = par2, 1=1)
     =fseek(poi,0) && �������  � ������ �����
     if !empty(par3)
      y = ''
      do while !feof(poi)
       y=allt(fgets(poi))
       if y = 'Resent-Message-Id' and allt(subs(y, at(':',y)+1)) = par3
        wait "��������� ���� " + lower(_name) + "� ��������� �����������!" wind time 1
        =fclos(poi)
        retu _name
       else
       endi
      endd
     else
      =fseek(poi,0) && �������  � ������ �����
      y    = ''
      rmid = ''
      do while !feof(poi)
       y = allt(fgets(poi))
       if y = 'Resent-Message-Id'
        rmid = allt(subs(y, at(':',y)+1))
       else
       endi
      endd
      if empty(rmid)
       wait "��������� ���� " + lower(_name) wind time 1
       =fclos(poi)
       retu _name
      Else
       wait "��������� ���� " + lower(_name) wind time 1
       =fclos(poi)
       retu _name
      endi
     endi
    else
    endi
   endd
   =fclos(poi)
   _name = sys(2000, par0+'\'+par1, 1)
   poi = fopen(par0+'\'+_name)
  endd
  =fclos(poi)

Retu ''

FUNCTION TransKms
 LPARAMETERS tPolis, nVozvrat
 tPolis = ALLTRIM(tPolis)
 nVozvrat = IIF(EMPTY(nVozvrat), 1, nVozvrat)
 Pps   = LEFT(tPolis,6)
 Ppn   = INT(VAL(SUBSTR(tPolis,7)))
 IF BETWEEN(ppn, 1, 9999999999) AND chk_nkms(PADL(Ppn,10,'0')) 
  tPolis = Pps + ' ' +PADL(Ppn,10,'0')
 ELSE 
  WAIT WINDOW "������ � ������ ���!" NOWAIT 
 ENDIF 
 tPolis = IIF(nVozvrat = 2, '', tPolis)
RETURN PADR(ALLTRIM(tPolis),17)

FUNCTION TransENP
 LPARAMETERS tPolis, nVozvrat
 tPolis = ALLTRIM(tPolis)
 nVozvrat = IIF(EMPTY(nVozvrat), 1, nVozvrat)
 Pps   = LEFT(tPolis,6)
 Ppn   = INT(VAL(SUBSTR(tPolis,7)))
 IF BETWEEN(ppn, 1, 9999999999) AND chk_nenp(ALLTRIM(SUBSTR(tPolis,7)))
  tPolis = Pps + Padl(Ppn,10,'0')
 ELSE
  WAIT WINDOW "������ � ������ ���!" NOWAIT 
 ENDIF 
 tPolis = IIF(nVozvrat = 2, '', tPolis)
RETURN PADR(ALLTRIM(tPolis),17)

FUNCTION TransVS
 LPARAMETERS tPolis, nVozvrat
RETURN tPolis

FUNCTION chk_ogrn
PARAMETERS lvOGRN, tcOGRN

DO CASE 
 CASE VARTYPE(lvOGRN) = 'N'
  tcOGRN = PADL(lvOGRN,10)
 CASE VARTYPE(lvOGRN) = 'C'
  tcOGRN=PADL(ALLTRIM(lvOGRN),13)
 OTHERWISE 
  RETURN .f. 
ENDCASE

prt1 = INT(VAL(SUBSTR(tcOGRN,1,12)))
prt2 = INT(VAL(SUBSTR(tcOGRN,13,1)))
ttt = IIF(MOD(prt1,11)<10, MOD(prt1,11), 0)

IF ttt == prt2
 RETURN .t.
ELSE
 RETURN .f.
ENDIF 
 
RETURN .t.

DO CASE 
	CASE VARTYPE(lnENP) = 'N'
		lcENP = STRTRAN(STR(lnENP,16,0),' ','0')
	CASE VARTYPE(lnENP) = 'C'
		lcENP = PADL(ALLTRIM(lnENP),16,'0')
	OTHERWISE
		lcENP = STRTRAN(SPACE(16),' ','0')
ENDCASE

VLast = RIGHT(lcENP,1)
V2 = LEFT(lcENP,LEN(lcENP)-1)

VA1 = ''
VB1 = ''

lMod = MOD(LEN(V2),2)

IF lMod = 1
	lModa = 0
	lModb = -1
ELSE
	lModb = 0
	lModa = -1
ENDIF

FOR NPos = LEN(V2) + lModa TO 1 STEP -2
	VA1 = VA1 + SUBSTR(V2,NPos,1)
ENDFOR

FOR NPos = LEN(V2) + lModb TO 1 STEP -2
	VB1 = VB1 + SUBSTR(V2,NPos,1)
ENDFOR

VA = ALLTRIM(STR(INT(VAL(VA1)) * 2))
VBA = VB1 + VA

VC = 0
FOR NPos = 1 TO LEN(VBA)
	VC = VC + VAL(SUBSTR(VBA,NPos,1))
ENDFOR

VD = (CEILING(VC/10) * 10) - VC

IF NOT VD = VAL(VLast)
	RETURN (.T.)
ELSE
	RETURN (.F.)
ENDIF

FUNCTION chk_snils
 PARAMETERS lcSNILS && XXX-XXX-XXX YY ������ �����, ��� YY - ����������� �����
 tcSNILS = ALLTRIM(STRTRAN(SUBSTR(lcSNILS,1,11),'-',''))
 tnKC = INT(VAL(SUBSTR(lcSNILS,13,2)))
 
 IF EMPTY(lcSNILS)
  RETURN .t.
 ENDIF 

 IF VAL(tcSNILS) <= 1001998 && ��� ������� �������� �� �� �� �����������!
 ELSE
  tn_rslt = 0
  FOR npos=1 TO 9
   tn_rslt = tn_rslt + VAL(SUBSTR(tcSNILS,npos,1)) * (10-npos)
  ENDFOR 
  DO CASE 
   CASE tn_rslt < 100
    IF tn_rslt==tnKC
     RETURN .t.
    ELSE 
     RETURN .f. 
    ENDIF 
   CASE INLIST(tn_rslt, 100, 101)
    IF tnKC==0
     RETURN .t.
    ELSE 
     RETURN .f. 
    ENDIF 
   CASE tn_rslt > 101
    DO WHILE tn_rslt >= 102
     tn_rslt = tn_rslt - 101
    ENDDO 
    IF tn_rslt < 100
     IF tn_rslt==tnKC
      RETURN .t.
     ELSE 
      RETURN .f. 
     ENDIF 
    ELSE
     IF tnKC==0
      RETURN .t.
     ELSE 
      RETURN .f. 
     ENDIF 
    ENDIF 
   OTHERWISE 
    RETURN .f. 
  ENDCASE 
 ENDIF 

FUNCTION NameOfMonth
 PARAMETERS nmonth
 DO CASE 
  CASE nmonth == 1
   MonName = '������'
  CASE nmonth == 2
   MonName = '�������'
  CASE nmonth == 3
   MonName = '����'
  CASE nmonth == 4
   MonName = '������'
  CASE nmonth == 5
   MonName = '���'
  CASE nmonth == 6
   MonName = '����'
  CASE nmonth == 7
   MonName = '����'
  CASE nmonth == 8
   MonName = '������'
  CASE nmonth == 9
   MonName = '��������'
  CASE nmonth == 10
   MonName = '�������'
  CASE nmonth == 11
   MonName = '������'
  CASE nmonth == 12
   MonName = '�������'
  OTHERWISE 
   MonName = ''
 ENDCASE 
RETURN MonName

FUNCTION NameOfMonth2
 PARAMETERS nmonth
 DO CASE 
  CASE nmonth == 1
   MonName = '������'
  CASE nmonth == 2
   MonName = '�������'
  CASE nmonth == 3
   MonName = '�����'
  CASE nmonth == 4
   MonName = '������'
  CASE nmonth == 5
   MonName = '���'
  CASE nmonth == 6
   MonName = '����'
  CASE nmonth == 7
   MonName = '����'
  CASE nmonth == 8
   MonName = '�������'
  CASE nmonth == 9
   MonName = '��������'
  CASE nmonth == 10
   MonName = '�������'
  CASE nmonth == 11
   MonName = '������'
  CASE nmonth == 12
   MonName = '�������'
  OTHERWISE 
   MonName = ''
 ENDCASE 
RETURN MonName

FUNCTION erz_show(erz_status)
 DO CASE 
  CASE erz_status == 0
   erz_say = '�� �����.'
  CASE erz_status == 1
   erz_say = '���������'
  CASE erz_status == 2
   erz_say = ' ������� '
  OTHERWISE 
   erz_say = ''
 ENDCASE 
RETURN  erz_say

Func cpr
parameter _c
**  _c - ����� �� �����. �����������: �� 0 �� 999 ����������.
if _c<0
   return ''
endif
if _c=0
   return '���� ������'
endif
if _c>=1000000000000000
   return ''
endif
**  ����������� �������� � ����������:
**  m1(20,4), m2(5,3)
**  _p - �������� ����� �� ������.
set talk off
_c=1000*int(_c)+100*(round(_c,2)-int(_c))
_p=''
dime m1(20,4), m2(6,3)

m1(1,1)=''
m1(2,1)='����'
m1(3,1)='���'
m1(4,1)='���'
m1(5,1)='������'
m1(6,1)='����'
m1(7,1)='�����'
m1(8,1)='����'
m1(9,1)='������'
m1(10,1)='������'
m1(11,1)=''
m1(12,1)=''
m1(13,1)=''
m1(14,1)=''
m1(15,1)=''
m1(16,1)=''
m1(17,1)=''
m1(18,1)=''
m1(19,1)=''
m1(20,1)=''

m1(1,2)=''
m1(2,2)=''
m1(3,2)='��������'
m1(4,2)='��������'
m1(5,2)='�����'
m1(6,2)='���������'
m1(7,2)='����������'
m1(8,2)='���������'
m1(9,2)='�����������'
m1(10,2)='���������'
m1(11,2)='������'
m1(12,2)='�����������'
m1(13,2)='����������'
m1(14,2)='����������'
m1(15,2)='������������'
m1(16,2)='����������'
m1(17,2)='�����������'
m1(18,2)='����������'
m1(19,2)='������������'
m1(20,2)='������������'

m1(1,3)=''
m1(2,3)='���'
m1(3,3)='������'
m1(4,3)='������'
m1(5,3)='���������'
m1(6,3)='�������'
m1(7,3)='��������'
m1(8,3)='�������'
m1(9,3)='���������'
m1(10,3)='���������'
m1(11,3)=''
m1(12,3)=''
m1(13,3)=''
m1(14,3)=''
m1(15,3)=''
m1(16,3)=''
m1(17,3)=''
m1(18,3)=''
m1(19,3)=''
m1(20,3)=''

m1(1,4)=''
m1(2,4)='����'
m1(3,4)='���'
m1(4,4)='���'
m1(5,4)='������'
m1(6,4)='����'
m1(7,4)='�����'
m1(8,4)='����'
m1(9,4)='������'
m1(10,4)='������'
m1(11,4)=''
m1(12,4)=''
m1(13,4)=''
m1(14,4)=''
m1(15,4)=''
m1(16,4)=''
m1(17,4)=''
m1(18,4)=''
m1(19,4)=''
m1(20,4)=''

m2(1,1)='�������'
m2(2,1)='�����'
m2(3,1)='������'
m2(4,1)='�������'
m2(5,1)='��������'
m2(6,1)='��������'

m2(1,2)='�������'
m2(2,2)='�����'
m2(3,2)='������'
m2(4,2)='��������'
m2(5,2)='���������'
m2(6,2)='���������'

m2(1,3)='������'
m2(2,3)='������'
m2(3,3)='�����'
m2(4,3)='���������'
m2(5,3)='����������'
m2(6,3)='����������'

**  ��������!!!
for _i=1 to 6
   _tri=_c-1000*int(_c/1000)
   _c=int(_c/1000)
   _dva=_tri-100*int(_tri/100)
   if _dva>9.and._dva<20
      _sl2=m1(1+_dva,y(3*_i-1))+chr(32)+m2(_i,3)+chr(32)
      _sl1=m1(1+int(_tri/100),y(3*_i))+chr(32)
      _p=ltrim(_sl1)+ltrim(_sl2)+ltrim(_p)
   else
      _odin=_tri-10*int(_tri/10)
      _sl3=m1(1+_odin,y(3*_i-2))+chr(32)
      _sl2=m1(1+int(_dva/10),y(3*_i-1))+chr(32)
      _sl1=m1(1+int(_tri/100),y(3*_i))+chr(32)
      if len(alltrim(_sl3+_sl2+_sl1))>0
         _p=ltrim(_sl1)+ltrim(_sl2)+ltrim(_sl3);
            +m2(_i,iif(_odin=1,1,iif(_odin>1.and._odin<5,2,3)))+chr(32)+ltrim(_p)
      else
         if _i=2
            _p=m2(2,3)+chr(32)+ltrim(_p)
         endif
      endif
   endif
   if _c=0
      exit
   endif
endfor
*****wait wind _p
* set talk on
Return _p

FUNCTION MakeZip(tcPath)
 poi = FCREATE(tcPath)
 IF poi != -1
  =FWRITE(poi, CHR(80))
  =FWRITE(poi, CHR(75))
  =FWRITE(poi, CHR(5))
  =FWRITE(poi, CHR(6))
  FOR iii = 1 TO 18
   =FWRITE(poi, CHR(0))
  ENDFOR 
 ENDIF 
 =FCLOSE(poi)
RETURN 

FUNCTION IsDopUsl(para1)
 PRIVATE cod
 
 m.cod = para1

 m.lIsDopUsl = .F.

RETURN m.lIsDopUsl

FUNCTION FMp(para1, para2, para3, para4, para5, para6, para7, para8, para9, para10, para11, para12)
 *Ex: FMp(m.cod, m.lpuid, m.mcod, m.ds, m.prmcod, m.prmcods, m.otd, m.IsTpnR, m.tip_p, m.ord, m.lpu_ord, m.d_type)
 PRIVATE cod, lpuid, mcod, ds
 * Mp = '4' - ��������� �������
 * Mp = '8' - ��������� ������������
 * Mp = 'p' - ��������� �������
 * Mp = 's' - ��������� ������������
 * Mp = 'm' - ����
 m.Mp = ''

 m.cod     = para1
 m.lpuid   = para2
 m.mcod    = para3
 m.ds      = para4
 m.prmcod  = para5
 m.prmcods = para6
 m.otd     = SUBSTR(para7,2,2)
 m.IsTpnR  = para8 
 m.tip_p   = para9
 m.ord     = para10
 m.lpu_ord = para11
 m.d_type  = para12
 
 IF IsDental(m.cod, m.lpuid, m.mcod, m.ds)
  DO CASE 
   CASE EMPTY(m.prmcods)    && ���������������
    m.Mp = 's'
   CASE m.mcod  = m.prmcods && ���� ��������
    DO CASE 
     CASE m.IsTpnR = .T. OR INLIST(m.otd,'08')
      m.Mp = '8'
     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      m.Mp = '8'
     CASE m.otd='93' AND IsStac(m.mcod)
      m.Mp = '8'
     OTHERWISE 
       m.Mp = 's'
    ENDCASE 
   CASE m.mcod != m.prmcods && ����� ��������
     m.Mp = 's'
  ENDCASE 
 ELSE 
  DO CASE 
   CASE EMPTY(m.prmcod)    && ���������������
    DO CASE 
     CASE m.IsTpnR = .T. OR m.d_type='s' OR (INLIST(m.otd,'08'))
      m.Mp = '4'
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      m.Mp = 'm'
     CASE INLIST(m.otd,'01') AND IsStac(m.mcod)
      m.Mp = '4'
     CASE INLIST(m.otd,'70','73','90','93') AND IsStac(m.mcod)
      m.Mp = '4'
     CASE m.ord=7 AND m.lpu_ord=7665
      m.Mp = '4'
     ** ��������� 16.04.2019 �� ���������� ������
     CASE INLIST(INT(m.cod/1000),49,149) AND m.mcod!=m.prmcod AND m.tip_p=3 
      m.Mp = '4'
     CASE INLIST(INT(m.cod/1000),29,129) AND m.mcod!=m.prmcod AND m.tip_p=3 
      m.Mp = '4'
     ** ��������� 16.04.2019 �� ���������� ������
     OTHERWISE 
       m.Mp = 'p'
    ENDCASE 
   CASE m.mcod  = m.prmcod && ���� ��������
    DO CASE 
     CASE m.IsTpnR = .T. OR m.d_type='s' OR (INLIST(m.otd,'08'))
      m.Mp = '4'
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      m.Mp = 'm'
     CASE INLIST(m.otd,'70','73','93') AND IsStac(m.mcod)
      m.Mp = '4'
     OTHERWISE 
       m.Mp = 'p'
    ENDCASE 
   CASE m.mcod != m.prmcod && ����� ��������
    DO CASE 
     CASE m.IsTpnR = .T. OR m.d_type='s' OR (INLIST(m.otd,'08'))
      m.Mp = '4'
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      m.Mp = 'm'
     CASE INLIST(m.otd,'01') AND IsStac(m.mcod)
      m.Mp = '4'
     CASE INLIST(m.otd,'70','73','90','93') AND IsStac(m.mcod)
      m.Mp = '4'
     CASE m.ord=7 AND m.lpu_ord=7665
      m.Mp = '4'
     ** ��������� 16.04.2019 �� ���������� ������
     CASE INLIST(INT(m.cod/1000),49,149) AND m.mcod!=m.prmcod AND m.tip_p=3 
      m.Mp = '4'
     CASE INLIST(INT(m.cod/1000),29,129) AND m.mcod!=m.prmcod AND m.tip_p=3 
      m.Mp = '4'
     ** ��������� 16.04.2019 �� ���������� ������
     OTHERWISE 
      m.Mp = 'p'
    ENDCASE 
  ENDCASE 
 ENDIF 
 
RETURN m.Mp

FUNCTION IsDental(para1, para2, para3, para4)
 PRIVATE cod, lpuid, mcod, ds

 m.cod   = para1
 m.lpuid = para2
 m.mcod  = para3
 m.ds    = para4
 
 m.IsStomat   = IIF(SUBSTR(m.mcod,3,2)='07', .T., .F.)
 m.UslIskl    = IIF(FLOOR(m.cod/1000)=146, .T., .F.)
 m.IsIskl     = IIF(INLIST(m.lpuid, 1912, 1940, 2049), .T., .F.)

 m.IsStomatUsl  = IIF(INLIST(FLOOR(m.cod/1000), 9, 109), .T., .F.)
 m.IsStomatUsl2 = IIF(INLIST(m.cod,1101,1102,101171,101172), .T., .F.) && �������������

 m.lIsDental = IIF(((m.IsStomat AND !m.IsIskl) AND (m.IsStomatUsl OR m.IsStomatUsl2)) OR ;
  	((m.IsStomat AND m.IsIskl) AND (m.IsStomatUsl OR m.IsStomatUsl2 OR m.UslIskl)) OR ;
  	(!m.IsStomat AND (m.IsStomatUsl OR (m.IsStomatUsl2 AND LEFT(m.ds,2)='K0'))), .T., .F.)
 
RETURN m.lIsDental

FUNCTION IsPlkOtd(cotd)
 m.lcotd = INT(VAL(SUBSTR(cotd,2,2)))
 m.IsPlkOtd = IIF(INLIST(m.lcotd,0,1,8,85,90,91,92,93), .T., .F.)
RETURN m.IsPlkOtd

FUNCTION IsDstOtd(cotd)
 m.lcotd = INT(VAL(SUBSTR(cotd,2,2)))
 m.IsPlkOtd = IIF(INLIST(m.lcotd,80,81), .T., .F.)
RETURN m.IsPlkOtd

FUNCTION IsHOOtd(cotd) && �������� �� ��������� �������������
 m.lcotd = INT(VAL(SUBSTR(cotd,2,2)))
 m.lIsHO = IIF(INLIST(m.lcotd,19,20,21,23,25,26,27,29,30,32,34,35,40,55,57,58,65,68), .T., .F.)
RETURN m.lIsHO

FUNCTION IsGsp(nCod)
 m.lIsGsp = IIF(IsMes(nCod) OR IsVmp(nCod) OR IsKdS(nCod), .T., .F.)
RETURN m.lIsGsp

FUNCTION IsDst(nCod)
 m.lIsDst = IIF(IsKdP(nCod) OR IsEKO(nCod), .T., .F.)
RETURN m.lIsDst

FUNCTION IsPlk(nCod)
 m.lIsPlk = IIF(IsUsl(nCod), .T., .F.)
RETURN m.lIsPlk

FUNCTION IsMes(nCod)
* RETURN IIF(BETWEEN(nCod,61000,96999) OR BETWEEN(nCod,161000,196999), .t., .f.)
*RETURN IIF(BETWEEN(nCod,61000,94999) OR BETWEEN(nCod,161000,194999) OR BETWEEN(nCod,97107,97999), .t., .f.) 
RETURN IIF(BETWEEN(nCod,61000,94999) OR BETWEEN(nCod,161000,194999), .t., .f.) 

FUNCTION IsVMP(nCod)
*RETURN IIF(nCod>=200000, .t., .f.)
RETURN IIF(INLIST(FLOOR(nCod/1000), 200, 297, 300, 397), .t., .f.)

FUNCTION IsUsl(nCod)
RETURN IIF(BETWEEN(nCod,1,60999) OR BETWEEN(nCod,101000,160999), .t., .f.)

FUNCTION IsKD(nCod)
RETURN IIF(BETWEEN(nCod,97000,99999) OR BETWEEN(nCod,197000,199999), .t., .f.)

FUNCTION IsKDP(nCod)
*RETURN IIF(INLIST(FLOOR(nCod/1000),97,197) AND nCod!=97041 AND !BETWEEN(nCod,97107,97999),.t.,.f.)
RETURN IIF(INLIST(FLOOR(nCod/1000),97,197) AND nCod!=97041 ,.t.,.f.)

FUNCTION IsEKO(nCod)
RETURN IIF(nCod=97041,.t.,.f.)

FUNCTION IsKDS(nCod)
RETURN IIF(INLIST(FLOOR(nCod/1000),99,199),.t.,.f.)

FUNCTION IsKompl(nCod)
RETURN IIF(BETWEEN(nCod,1900,1905) OR BETWEEN(nCod,101927,101932) OR INLIST(nCod,15001,115001), .t., .f.)

FUNCTION IsPat(nCod)
RETURN IIF(INLIST(FLOOR(nCod/1000),59,159),.t.,.f.)

FUNCTION IsSimult(nCod)
RETURN IIF(INLIST(FLOOR(nCod/1000),51,151,52,152,53,153,54,154,55,155),.t.,.f.)

FUNCTION Is02(nCod)
RETURN IIF(INLIST(FLOOR(nCod/1000),96,196) OR INLIST(nCod,56031,156002),.t.,.f.)

FUNCTION Is70(cOtd)
RETURN IIF(SUBSTR(cOtd,2,2)='70', .t., .f.)

FUNCTION Is73(cOtd)
RETURN IIF(SUBSTR(cOtd,2,2)='73', .t., .f.)

FUNCTION IsOKRecID(pRecId)
 len_recid=ALLTRIM(pRecid)
 tresult = .t.
 FOR pos=1 TO len_recid
  symbol=SUBSTR(pRecId,pos,1)
  IF !(BETWEEN(x,48,57) OR BETWEEN(x,65,90) OR BETWEEN(x,97,122))
   tresult = .F.
   LOOP 
  ENDIF 
 ENDFOR 
RETURN tresult

FUNCTION PDFCREATEDID
 m.PREVARRLEN = ALEN(ARRDATA)
 DIMENSION ARRDATA(m.PREVARRLEN+12)
 ARRDATA(m.PREVARRLEN + 1) = "1" + PDFOBJECT_BEGIN
 ARRDATA(m.PREVARRLEN + 2) = "<<"
 ARRDATA(m.PREVARRLEN + 3) = "/Title    (Ambulance account)"
 ARRDATA(m.PREVARRLEN + 4) = "/Author   <feff041C0438044504300438043B00200420044F0431043E0432>" && "������ �����"
 ARRDATA(m.PREVARRLEN + 5) = "/Subject  (Universal Healthcare Program)"
 ARRDATA(m.PREVARRLEN + 6) = "/Keywords (Universal Healthcare Program)"
 ARRDATA(m.PREVARRLEN + 7) = "/CreationDate (D:19410622040000)"
 ARRDATA(m.PREVARRLEN + 8) = "/ModDate      (D:19450509120000)"
 ARRDATA(m.PREVARRLEN + 9) = "/Creator  <feff00220421043804400435043D0430002B0022>"
 ARRDATA(m.PREVARRLEN + 10) = "/Producer (Low-level PDF creating software)"
 ARRDATA(m.PREVARRLEN + 11) = ">>"
 ARRDATA(m.PREVARRLEN + 12) = PDFOBJECT_END
RETURN(.T.)

FUNCTION PDFADDFONT()
 m.PREVARRLEN = ALEN(ARRDATA)
 DIMENSION ARRDATA(m.PREVARRLEN+63)

 ARRDATA(m.PREVARRLEN + 1) = "4" + PDFOBJECT_BEGIN
 ARRDATA(m.PREVARRLEN + 2) = "<<"
 ARRDATA(m.PREVARRLEN + 3) = "/Type /Font"
 ARRDATA(m.PREVARRLEN + 4) = "/Subtype /TrueType"
 ARRDATA(m.PREVARRLEN + 5) = "/Name /F1"
 ARRDATA(m.PREVARRLEN + 6) = "/BaseFont /ArialMT"
 ARRDATA(m.PREVARRLEN + 7) = "/FirstChar 32"
 ARRDATA(m.PREVARRLEN + 8) = "/LastChar 255"
 ARRDATA(m.PREVARRLEN + 9)=  "/Widths 5 0 R"
 ARRDATA(m.PREVARRLEN + 10) = "/FontDescriptor 6 0 R"
 ARRDATA(m.PREVARRLEN + 11) = "/Encoding  7 0 R" 
 ARRDATA(m.PREVARRLEN + 12) = ">>"
 ARRDATA(m.PREVARRLEN + 13) = PDFOBJECT_END

 ARRDATA(m.PREVARRLEN + 14) = "5" + PDFOBJECT_BEGIN
 ARRDATA(m.PREVARRLEN + 15) = "[ 277 277 354 556 556 889 666 190 333 333 389 583 277 333 277 "
 ARRDATA(m.PREVARRLEN + 16) = "277 556 556 556 556 556 556 556 556 556 556 277 277 583 583 583 556 "
 ARRDATA(m.PREVARRLEN + 17) = "1015 666 666 722 722 666 610 777 722 277 500 666 556 833 722 777 666 "
 ARRDATA(m.PREVARRLEN + 18) = "777 722 666 610 722 666 943 666 666 610 277 277 277 469 556 333 556 "
 ARRDATA(m.PREVARRLEN + 19) = "556 500 556 556 277 556 556 222 222 500 222 833 556 556 556 556 333 "
 ARRDATA(m.PREVARRLEN + 20) = "500 277 556 500 722 500 500 500 333 259 333 583 0 864 541 222 364 333 "
 ARRDATA(m.PREVARRLEN + 21) = "1000 556 556 556 1000 1057 333 1010 582 854 718 556 222 222 333 333 "
 ARRDATA(m.PREVARRLEN + 22) = "350 556 1000 0 1000 906 333 812 437 556 552 277 635 500 500 556 488 "
 ARRDATA(m.PREVARRLEN + 23) = "259 556 667 736 718 556 583 333 736 277 399 548 277 222 411 576 537 "
 ARRDATA(m.PREVARRLEN + 24) = "277 556 1072 510 556 222 666 500 277 666 656 666 541 677 666 923 604 "
 ARRDATA(m.PREVARRLEN + 25) = "718 718 582 656 833 722 777 718 666 722 610 635 760 666 739 666 916 "
 ARRDATA(m.PREVARRLEN + 26) = "937 791 885 656 718 1010 722 556 572 531 364 583 556 668 458 558 558 "
 ARRDATA(m.PREVARRLEN + 27) = "437 583 687 552 556 541 556 500 458 500 822 500 572 520 802 822 625 "
 ARRDATA(m.PREVARRLEN + 28) = "718 520 510 750 541 ]"
 ARRDATA(m.PREVARRLEN + 29) = PDFOBJECT_END

 ARRDATA(m.PREVARRLEN + 30) = "6" + PDFOBJECT_BEGIN
 ARRDATA(m.PREVARRLEN + 31) = "<<"
 ARRDATA(m.PREVARRLEN + 32) = "/Type /FontDescriptor"
 ARRDATA(m.PREVARRLEN + 33) = "/FontName /ArialMT"
 ARRDATA(m.PREVARRLEN + 34) = "/Flags 32"
 ARRDATA(m.PREVARRLEN + 35) = "/FontBBox [-222 -324 1071 1037]"
 ARRDATA(m.PREVARRLEN + 36) = "/ItalicAngle 0"
 ARRDATA(m.PREVARRLEN + 37) = "/Ascent 728"
 ARRDATA(m.PREVARRLEN + 38) = "/Descent -210"
 ARRDATA(m.PREVARRLEN + 39) = "/CapHeight 699" 
 ARRDATA(m.PREVARRLEN + 40) = "/StemV 80"
 ARRDATA(m.PREVARRLEN + 41) = ">>"
 ARRDATA(m.PREVARRLEN + 42) = PDFOBJECT_END

 ARRDATA(m.PREVARRLEN + 43) = "7" + PDFOBJECT_BEGIN
 ARRDATA(m.PREVARRLEN + 44) = "<</Type /Encoding /Differences"
 ARRDATA(m.PREVARRLEN + 45) = "["
 ARRDATA(m.PREVARRLEN + 46) = "32 /space /exclam /quotedbl /numbersign /dollar /percent /ampersand /quotesingle /parenleft /parenright /asterisk /plus /comma"
 ARRDATA(m.PREVARRLEN + 47) = "/hyphen /period /slash /zero /one /two /three /four /five /six /seven /eight /nine /colon /semicolon /less /equal /greater /question /at /A"
 ARRDATA(m.PREVARRLEN + 48) = "/B /C /D /E /F /G /H /I /J /K /L /M /N /O /P /Q /R /S /T /U /V /W /X /Y /Z /bracketleft /backslash /bracketright /asciicircum /underscore"
 ARRDATA(m.PREVARRLEN + 49) = "/grave /a /b /c /d /e /f /g /h /i /j /k /l /m /n /o /p /q /r /s /t /u /v /w /x /y /z /braceleft /bar /braceright /asciitilde /.notdef"
 ARRDATA(m.PREVARRLEN + 50) = "/afii10051 /afii10052 /quotesinglbase /afii10100 /quotedblbase /ellipsis /dagger /daggerdbl /Euro /perthousand /afii10058"
 ARRDATA(m.PREVARRLEN + 51) = "/guilsinglleft /afii10059 /afii10061 /afii10060 /afii10145 /afii10099 /quoteleft /quoteright /quotedblleft /quotedblright /bullet /endash"
 ARRDATA(m.PREVARRLEN + 52) = "/emdash /.notdef /trademark /afii10106 /guilsinglright /afii10107 /afii10109 /afii10108 /afii10193 /space /afii10062 /afii10110"
 ARRDATA(m.PREVARRLEN + 53) = "/afii10057 /currency /afii10050 /brokenbar /section /afii10023 /copyright /afii10053 /guillemotleft /logicalnot /hyphen /registered"
 ARRDATA(m.PREVARRLEN + 54) = "/afii10056 /degree /plusminus /afii10055 /afii10103 /afii10098 /mu /paragraph /periodcentered /afii10071 /afii61352 /afii10101"
 ARRDATA(m.PREVARRLEN + 55) = "/guillemotright /afii10105 /afii10054 /afii10102 /afii10104 /afii10017 /afii10018 /afii10019 /afii10020 /afii10021 /afii10022 /afii10024"
 ARRDATA(m.PREVARRLEN + 56) = "/afii10025 /afii10026 /afii10027 /afii10028 /afii10029 /afii10030 /afii10031 /afii10032 /afii10033 /afii10034 /afii10035 /afii10036"
 ARRDATA(m.PREVARRLEN + 57) = "/afii10037 /afii10038 /afii10039 /afii10040 /afii10041 /afii10042 /afii10043 /afii10044 /afii10045 /afii10046 /afii10047 /afii10048"
 ARRDATA(m.PREVARRLEN + 58) = "/afii10049 /afii10065 /afii10066 /afii10067 /afii10068 /afii10069 /afii10070 /afii10072 /afii10073 /afii10074 /afii10075 /afii10076"
 ARRDATA(m.PREVARRLEN + 59) = "/afii10077 /afii10078 /afii10079 /afii10080 /afii10081 /afii10082 /afii10083 /afii10084 /afii10085 /afii10086 /afii10087 /afii10088"
 ARRDATA(m.PREVARRLEN + 60) = "/afii10089 /afii10090 /afii10091 /afii10092 /afii10093 /afii10094 /afii10095 /afii10096 /afii10097"
 ARRDATA(m.PREVARRLEN + 61) = "]"
 ARRDATA(m.PREVARRLEN + 62) = ">>"
 ARRDATA(m.PREVARRLEN + 63) = PDFOBJECT_END
	
RETURN(.T.)

FUNCTION PDFINITIALISE
 m.PREVARRLEN = ALEN(ARRDATA)
 DIMENSION ARRDATA(m.PREVARRLEN+15)

 ARRDATA(m.PREVARRLEN + 1) = "8" + PDFOBJECT_BEGIN
 ARRDATA(m.PREVARRLEN + 2) = "<<"
 ARRDATA(m.PREVARRLEN + 3) = "/Font << /F1 4 0 R   >>  /ProcSet [ /PDF  /Text ]"
 ARRDATA(m.PREVARRLEN + 4) = ">>"
 ARRDATA(m.PREVARRLEN + 5) = PDFOBJECT_END

 ARRDATA(m.PREVARRLEN + 6) = "9" + PDFOBJECT_BEGIN
 ARRDATA(m.PREVARRLEN + 7) = "<<"
 ARRDATA(m.PREVARRLEN + 8) = "/HideToolbar false"
 ARRDATA(m.PREVARRLEN + 9) = "/HideMenubar false"
 ARRDATA(m.PREVARRLEN + 10) = "/HideWindowUI false"
 ARRDATA(m.PREVARRLEN + 11) = "/FitWindow true"
 ARRDATA(m.PREVARRLEN + 12) = "/CenterWindow false"
 ARRDATA(m.PREVARRLEN + 13) = "/DisplayDocTitle false"
 ARRDATA(m.PREVARRLEN + 14) = ">>"
 ARRDATA(m.PREVARRLEN + 15) = PDFOBJECT_END
RETURN(.T.)

FUNCTION PDFBEGINPAGE
 m.OBJECTCOUNT = m.OBJECTCOUNT + 1

 m.STRPAGES = m.STRPAGES + " " + ALLTRIM(STR(m.OBJECTCOUNT)) + " 0 R"

 m.PREVARRLEN = ALEN(ARRDATA)
 DIMENSION ARRDATA(m.PREVARRLEN + 12)

 ARRDATA(m.PREVARRLEN + 1) = ALLTRIM(STR(m.OBJECTCOUNT)) + PDFOBJECT_BEGIN
 ARRDATA(m.PREVARRLEN + 2) = "<<"
 ARRDATA(m.PREVARRLEN + 3) = "/Type /Page"
 ARRDATA(m.PREVARRLEN + 4) = "/Parent 3 0 R" 

 ARRDATA(m.PREVARRLEN + 5) = "/Contents " + ALLTRIM(STR(m.OBJECTCOUNT+1)) + " 0 R"
 ARRDATA(m.PREVARRLEN + 6) = ">>"
 ARRDATA(m.PREVARRLEN + 7) = PDFOBJECT_END

 DIMENSION ARRXREF(ALEN(ARRXREF) + 1)
 ARRXREF(ALEN(ARRXREF)) = PDFXREFMARKER

 m.OBJECTCOUNT = m.OBJECTCOUNT + 1
 ARRDATA(m.PREVARRLEN + 8) = ALLTRIM(STR(m.OBJECTCOUNT)) + PDFOBJECT_BEGIN
 ARRDATA(m.PREVARRLEN + 9) = "<<"
 ARRDATA(m.PREVARRLEN + 10) = "/Length " + ALLTRIM(STR(m.OBJECTCOUNT+1)) + " 0 R"
 ARRDATA(m.PREVARRLEN + 11) = ">>"

 ARRDATA(m.PREVARRLEN + 12) = "stream"
 
RETURN 

FUNCTION BT
 m.PREVARRLEN = ALEN(ARRDATA)
 DIMENSION ARRDATA(m.PREVARRLEN+1)

 ARRDATA(m.PREVARRLEN + 1) = "BT"
 m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(m.PREVARRLEN + 1))+2
RETURN 

FUNCTION ET
 m.PREVARRLEN = ALEN(ARRDATA)
 DIMENSION ARRDATA(m.PREVARRLEN+1)

 ARRDATA(m.PREVARRLEN+ 1) = "ET"  && end text
 m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(m.PREVARRLEN+ 1))+2
RETURN 

FUNCTION PDFSETSIZEOFFONT(m.SIZEFONT)
 PRIVATE m.SIZEFONT
 m.PREVARRLEN = ALEN(ARRDATA)
 DIMENSION ARRDATA(m.PREVARRLEN+1)

 ARRDATA(m.PREVARRLEN + 1) = "/F1 "+ALLTRIM(STR(m.SIZEFONT))+" Tf" && ������������� ������ ������
 m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(m.PREVARRLEN+ 1))+2
RETURN 

FUNCTION PDFSETTM(A,B,C,D,E,F)
 && [a b c d e f Tm] a-Th-horizontal scaling, default value=1 - 100%
 && b - ������ ������ ������ ������� �������, �� ��������� - 0
 && c - ������ ������ ������, �� ��������� - 0 
 && d - vertiacl scaling, default value=1 - 100%
 && e- 
 && f -
 &&
 PRIVATE m.A,m.B,m.C,m.D,m.E,m.F
 m.A = ALLTRIM(STR(m.A))
 m.B = ALLTRIM(STR(m.B))
 m.C = ALLTRIM(STR(m.C))
 m.D = ALLTRIM(STR(m.D))
 m.E = ALLTRIM(STR(m.E))
 m.F = ALLTRIM(STR(m.F))
 m.PREVARRLEN = ALEN(ARRDATA)
 DIMENSION ARRDATA(m.PREVARRLEN+1)

 ARRDATA(m.PREVARRLEN + 1) = m.A+" "+m.B+" "+m.C+" "+m.D+" "+m.E+" "+m.F+" Tm"
 m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(m.PREVARRLEN+ 1))+2
RETURN 

FUNCTION PDFSETTEXTLEADING(m.TEXTLEADING)
 PRIVATE m.TEXTLEADING
 m.PREVARRLEN = ALEN(ARRDATA)
 DIMENSION ARRDATA(m.PREVARRLEN+1)

 ARRDATA(m.PREVARRLEN + 1) = ALLTRIM(STR(m.TEXTLEADING))+" TL" && ����������� ����������
 m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(m.PREVARRLEN+ 1))+2
RETURN 

FUNCTION PDFSETCHARSPACING(m.CHARSPACING)
 PRIVATE m.CHARSPACING
 m.PREVARRLEN = ALEN(ARRDATA)
 DIMENSION ARRDATA(m.PREVARRLEN+1)

 ARRDATA(m.PREVARRLEN + 1) = ALLTRIM(STR(m.CHARSPACING))+" Tc" && ������������ ����������
 m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(m.PREVARRLEN+ 1))+2
RETURN 

FUNCTION PDFSETWORDSPACING(m.CHARSPACING)
 PRIVATE m.CHARSPACING
 m.PREVARRLEN = ALEN(ARRDATA)
 DIMENSION ARRDATA(m.PREVARRLEN+1)

 ARRDATA(m.PREVARRLEN + 1) = ALLTRIM(STR(m.CHARSPACING))+" Tw" && ���������� ����������
 m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(m.PREVARRLEN+ 1))+2
RETURN 

FUNCTION PDFSETINITTEXTPOSITION(m.XPOSITION, m.YPOSITION)
 PRIVATE m.XPOSITION, m.YPOSITION
 m.PREVARRLEN = ALEN(ARRDATA)
 DIMENSION ARRDATA(m.PREVARRLEN+1)

 ARRDATA(m.PREVARRLEN + 1) = PADL(m.XPOSITION,3,'0')+" "+PADL(m.YPOSITION,3,'0')+" Td" && �������� ��������� ������
 m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(m.PREVARRLEN+ 1))+2
RETURN 

FUNCTION PDFTYPETEXT(m.LINEOFTEXT, m.ISNEWLINE, m.XPOS, m.YPOS)
 PRIVATE m.LINEOFTEXT, m.ISNEWLINE, m.XPOS, m.YPOS
 DO CASE 
  CASE PARAMETERS()=0
   m.LINEOFTEXT = ''
   m.ISNEWLINE = .T.
  CASE PARAMETERS()=1
   m.ISNEWLINE = .T.
  OTHERWISE 
 ENDCASE 
 
 IF m.ISNEWLINE = .F.
  BT()
  PDFSETINITTEXTPOSITION(m.XPOS, m.YPOS)
 ENDIF 

 DIMENSION ARRDATA(ALEN(ARRDATA)+1)

 m.LINEOFTEXT = STRTRAN(m.LINEOFTEXT,"\","\\")
 m.LINEOFTEXT = STRTRAN(m.LINEOFTEXT,"(","\(")
 m.LINEOFTEXT = STRTRAN(m.LINEOFTEXT,")","\)")
 
 ARRDATA(ALEN(ARRDATA)) = IIF(m.ISNEWLINE=.T., "T* ", "")+"("+m.LINEOFTEXT+") Tj"
* ARRDATA(ALEN(ARRDATA)) = "("+m.LINEOFTEXT+") '"
* ARRDATA(ALEN(ARRDATA)) = "T* ["+m.LINEOFTEXT+"] TJ"

 m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(ALEN(ARRDATA)))+2

 IF m.ISNEWLINE = .F.
  ET()
 ENDIF 

RETURN 

FUNCTION PDFTYPETEXT3(m.LINEOFTEXT)
 PRIVATE m.LINEOFTEXT
 DIMENSION ARRDATA(ALEN(ARRDATA)+1)

 m.LINEOFTEXT = STRTRAN(m.LINEOFTEXT,"\","\\")
 m.LINEOFTEXT = STRTRAN(m.LINEOFTEXT,"(","\(")
 m.LINEOFTEXT = STRTRAN(m.LINEOFTEXT,")","\)")
 
 ARRDATA(ALEN(ARRDATA)) = "("+m.LINEOFTEXT+") Tj"

 m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(ALEN(ARRDATA)))+2

RETURN 

FUNCTION PDFTYPETEXT2(m.LINEOFTEXT)
 PRIVATE m.LINEOFTEXT
 DIMENSION ARRDATA(ALEN(ARRDATA)+1)

* m.LINEOFTEXT = STRTRAN(m.LINEOFTEXT,"\","\\")
* m.LINEOFTEXT = STRTRAN(m.LINEOFTEXT,"(","\(")
* m.LINEOFTEXT = STRTRAN(m.LINEOFTEXT,")","\)")
 
 ARRDATA(ALEN(ARRDATA)) = "T* ["+m.LINEOFTEXT+"] TJ"

 m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(ALEN(ARRDATA)))+2

RETURN 


FUNCTION PDFTYPETEXT4(m.LINEOFTEXT)
 PRIVATE m.LINEOFTEXT
 DIMENSION ARRDATA(ALEN(ARRDATA)+1)

 m.LINEOFTEXT = STRTRAN(m.LINEOFTEXT,"\","\\")
 m.LINEOFTEXT = STRTRAN(m.LINEOFTEXT,"(","\(")
 m.LINEOFTEXT = STRTRAN(m.LINEOFTEXT,")","\)")
 
 ARRDATA(ALEN(ARRDATA)) = "1 0 0 1 050 050 Tm ("+m.LINEOFTEXT+") Tj"

 m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(ALEN(ARRDATA)))+2

RETURN 

*FUNCTION PDFADDTABLE(A1,A2,A3,A6) && 100 120 m 500 120 l 500 40 l 100 40 l 100 120 l S
* PRIVATE A1,A2,A3,A4,A5,A6,A7,A8
* A1 = PADL(A1,3,'0')
* A6 = PADL(A2-A6,3,'0')
* A2 = PADL(A2,3,'0')
* A3 = PADL(A3,3,'0')
* A4 = A2
* A5 = A3
* A7 = A1
* A8 = A6
 
* m.PREVARRLEN = ALEN(ARRDATA)
* DIMENSION ARRDATA(m.PREVARRLEN+1)

* ARRDATA(m.PREVARRLEN+ 1) = "1 w "+A1+" "+A2+" m "+A3+" "+A4+" l "+A5+" "+A6+" l "+A7+" "+A8+" l "+A1+" "+A2+" l S"
* m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(m.PREVARRLEN+ 1))+2
*RETURN 

FUNCTION PDFADDTABLE(X,Y,WIDTH,HEIGHT) && X,Y, width, heigt re
 PRIVATE A1,A2,A3,A4
 A1 = PADL(X,3,'0')
 A2 = PADL(Y,3,'0')
 A3 = PADL(WIDTH,3,'0')
 A4 = PADL(HEIGHT,3,'0')
 
 m.PREVARRLEN = ALEN(ARRDATA)
 DIMENSION ARRDATA(m.PREVARRLEN+1)

 ARRDATA(m.PREVARRLEN+ 1) = A1+" "+A2+" "+A3+" "+A4+" re S"
 m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(m.PREVARRLEN+ 1))+2
RETURN 

FUNCTION PDFDRAWLINE(A1,A2,A3,A4)
 PRIVATE A1,A2,A3,A4
 A1 = PADL(A1,3,'0')
 A2 = PADL(A2,3,'0')
 A3 = PADL(A3,3,'0')
 A4 = PADL(A4,3,'0')
 
 m.PREVARRLEN = ALEN(ARRDATA)
 DIMENSION ARRDATA(m.PREVARRLEN+1)

 ARRDATA(m.PREVARRLEN+ 1) = "1 w "+A1+" "+A2+" m "+A3+" "+A4+" l S"
 m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(m.PREVARRLEN+ 1))+2
RETURN 

FUNCTION PDFENDPAGE
 m.PREVARRLEN = ALEN(ARRDATA)
 DIMENSION ARRDATA(m.PREVARRLEN+5)

 ARRDATA(m.PREVARRLEN+ 1) = "endstream"
 ARRDATA(m.PREVARRLEN+ 2) = PDFOBJECT_END

 DIMENSION ARRXREF(ALEN(ARRXREF)+1)
 ARRXREF(ALEN(ARRXREF)) = PDFXREFMARKER

 m.OBJECTCOUNT = m.OBJECTCOUNT + 1
 ARRDATA(m.PREVARRLEN+ 3) = ALLTRIM(STR(m.OBJECTCOUNT)) + PDFOBJECT_BEGIN
 ARRDATA(m.PREVARRLEN+ 4) = ALLTRIM(STR(m.STREAMLENGTH))

 ARRDATA(m.PREVARRLEN+ 5) = PDFOBJECT_END

 DIMENSION ARRXREF(ALEN(ARRXREF)+1)
 ARRXREF(ALEN(ARRXREF)) = PDFXREFMARKER
RETURN 

FUNCTION PDFWRITE
	PARAMETERS m.OUTFILENAME
	PRIVATE I, m.OUTFILENAME,m.STRING,m.TEMPBIT,X,XREFINDEX,m.OFFSETPOSN
	DECLARE ARRXREFS(m.OBJECTCOUNT-1) && SET UP A NEW XREFS TABLE TO HOLD POSITIONS OF ALL BLOCKS
	FOR X = 1 TO ALEN(ARRXREFS)
		ARRXREFS(X)=0
	NEXT
	m.OFFSETPOSN = 0
	m.STRING = ""
	XREFINDEX = 0
	FOR I = 1 TO ALEN(ARRDATA)
		m.TEMPBIT = ARRDATA(I)
		IF TYPE("m.tempbit") = "C" .AND. !EMPTY(m.TEMPBIT)
			DO CASE
			CASE RIGHT(UPPER(m.TEMPBIT),LEN(PDFOBJECT_BEGIN)) = UPPER(PDFOBJECT_BEGIN)
			    * ����� ������ ����� - ���������� ��� �������������� � �����, �������
			    * ����� ����� ������ �� ������ ����� �� ����� �����
				X =  VAL(LEFT(m.TEMPBIT,LEN(m.TEMPBIT)-LEN(PDFOBJECT_BEGIN)))
				ARRXREFS(X) = LEN(m.STRING)
			CASE UPPER(m.TEMPBIT) = PDFXREFMARKER
				XREFINDEX = XREFINDEX +1
				m.TEMPBIT = RIGHT("0000000000"+ALLTRIM(STR(ARRXREFS(XREFINDEX))),10) + XREF_END_CHAR
			CASE UPPER(m.TEMPBIT) == "XREF"
				** ������ xref-�����
				m.OFFSETPOSN = LEN(m.STRING)
			CASE UPPER(m.TEMPBIT) == "M.OFFSETPOSN"
				** ����������� �������� ����� ������ "M.OFFSETPOSN"
				m.TEMPBIT = ALLTRIM(STR(m.OFFSETPOSN))
			ENDCASE
			m.STRING = m.STRING + m.TEMPBIT + CHR(13)+CHR(10)
		ENDIF
	NEXT
	STRTOFILE(m.STRING,m.OUTFILENAME)
RETURN

FUNCTION PDFADDCATALOGDETAILS(m.XSIZE, m.YSIZE)
    PRIVATE m.XSIZE, m.YSIZE && 595 842
	PRIVATE m.STARTSIZE
	m.STARTSIZE = ALEN(ARRDATA)
	** ADD 17 LINES TO THE DATA ARRAY
	DIMENSION ARRDATA(ALEN(ARRDATA)+18)
	ARRDATA(ALEN(ARRDATA)) = ""
	ARRDATA(m.STARTSIZE + 1) = "2" + PDFOBJECT_BEGIN
	ARRDATA(m.STARTSIZE + 2) = "<<"
	ARRDATA(m.STARTSIZE + 3) = "/Type /Catalog"
	ARRDATA(m.STARTSIZE + 4) = "/Pages 3 0 R"   
	ARRDATA(m.STARTSIZE + 5) = "/PageLayout /SinglePage"
	ARRDATA(m.STARTSIZE + 6) = "/ViewerPreferences 9 0 R"
	ARRDATA(m.STARTSIZE + 7) = "/PageMode /UseNone"
	ARRDATA(m.STARTSIZE + 8) = ">>"
	ARRDATA(m.STARTSIZE + 9) = PDFOBJECT_END

	ARRDATA(m.STARTSIZE + 10) = "3" + PDFOBJECT_BEGIN
	ARRDATA(m.STARTSIZE + 11) = "<<"
	ARRDATA(m.STARTSIZE + 12) = "/Type /Pages"
	ARRDATA(m.STARTSIZE + 13) = "/Count " + ALLTRIM(STR(m.NOPAGES))
	ARRDATA(m.STARTSIZE + 14) = "/Kids [" + m.STRPAGES + " ]"
    ARRDATA(m.STARTSIZE + 15) = "/MediaBox [ 0 0 "+PADL(m.XSIZE,3,'0')+" "+PADL(m.YSIZE,3,'0')+" ]" && 8,28 in x 11,69 in - ����������� A4-��������
    ARRDATA(m.STARTSIZE + 16) = "/Resources 8 0 R"

	ARRDATA(m.STARTSIZE + 17) = ">>"
	ARRDATA(m.STARTSIZE + 18) = PDFOBJECT_END
RETURN


FUNCTION PDFFOOTER
	PRIVATE I
	DIMENSION ARRXREF(ALEN(ARRXREF)+9)
	ARRXREF(ALEN(ARRXREF)-8) = "trailer"
	ARRXREF(ALEN(ARRXREF)-7) = "<<"
	ARRXREF(ALEN(ARRXREF)-6) = "/Size " + ALLTRIM(STR(m.OBJECTCOUNT))
	ARRXREF(ALEN(ARRXREF)-5) = "/Root 2 0 R"
	ARRXREF(ALEN(ARRXREF)-4) = "/Info 1 0 R"
	ARRXREF(ALEN(ARRXREF)-3) = ">>"
	ARRXREF(ALEN(ARRXREF)-2) = "startxref"
	ARRXREF(ALEN(ARRXREF)-1) = "M.OFFSETPOSN" && REPLACE THIS VALUE AS WRITTEN OUT TO TEXT FILE
	ARRXREF(ALEN(ARRXREF)) = "%%EOF"
	FOR I = 1 TO ALEN(ARRXREF)
		IF !EMPTY(ARRXREF(I))
			DIMENSION ARRDATA(ALEN(ARRDATA)+1)
			ARRDATA(ALEN(ARRDATA)) = ARRXREF(I)
		ENDIF
	NEXT
RETURN(.T.)

FUNCTION CokrName(para1)
 PRIVATE m.cokr
 m.cokr = m.para1
 DO CASE 
  CASE m.cokr = 1
   m.cokrname = '���'
  CASE m.cokr = 2
   m.cokrname = '���'
  CASE m.cokr = 3
   m.cokrname = '����'
  CASE m.cokr = 4
   m.cokrname = '���'
  CASE m.cokr = 5
   m.cokrname = '����'
  CASE m.cokr = 6
   m.cokrname = '���'
  CASE m.cokr = 7
   m.cokrname = '����'
  CASE m.cokr = 8
   m.cokrname = '���'
  CASE m.cokr = 9
   m.cokrname = '����'
  CASE m.cokr = 10
   m.cokrname = '�����'
  CASE m.cokr = 11
   m.cokrname = '��� �.'
  CASE m.cokr = 12
   m.cokrname = '����'
  CASE m.cokr = 13
   m.cokrname = '����'
  OTHERWISE 
   m.cokrname = '???'
 ENDCASE 
RETURN m.cokrname

FUNCTION IsOpenWordDoc(parr)

 IF PARAMETERS()<1
  RETURN .f.
 ENDIF 

 LOCAL m.docname
 m.docname = STRTRAN(LOWER(ALLTRIM(parr)), '.doc', '')
 IF EMPTY(m.docname)
  RETURN .f.
 ENDIF 

 m.lIsWordActive = .t.
 TRY 
  oWord = GETOBJECT(,"Word.Application")
 CATCH 
  m.lIsWordActive = .f.
 ENDTRY 
 IF m.lIsWordActive = .f.
  RETURN .f. 
 ENDIF 

 oDocs = oWord.Documents
 nDocs = oDocs.count
 
 IF nDocs<=0
  RETURN .f. 
 ENDIF 
 
 m.lResult = .f.
 FOR EACH oDoc IN oDocs
  m.exname = STRTRAN(LOWER(oDoc.name), '.doc', '')
  IF m.exname=m.docname
   m.lResult = .t.
   EXIT 
  ENDIF 
 ENDFOR 
 
RETURN m.lResult

FUNCTION IsOpenExcelDoc(parr)

 IF PARAMETERS()<1
  RETURN .f.
 ENDIF 

 LOCAL m.docname
 m.docname = STRTRAN(LOWER(ALLTRIM(parr)), '.xls', '')
 IF EMPTY(m.docname)
  RETURN .f.
 ENDIF 

 m.lIsWordActive = .t.
 TRY 
  oWord = GETOBJECT(,"Excel.Application")
 CATCH 
  m.lIsWordActive = .f.
 ENDTRY 
 IF m.lIsWordActive = .f.
  RETURN .f. 
 ENDIF 

 oDocs = oWord.WorkBooks
 nDocs = oDocs.count
 
 IF nDocs<=0
  RETURN .f. 
 ENDIF 
 
 m.lResult = .f.
 FOR EACH oDoc IN oDocs
  m.exname = STRTRAN(LOWER(oDoc.name), '.xls', '')
  IF m.exname=m.docname
   m.lResult = .t.
   EXIT 
  ENDIF 
 ENDFOR 
 
RETURN m.lResult

FUNCTION CloseExcelDoc(parr)

 IF PARAMETERS()<1
  RETURN .f.
 ENDIF 

 LOCAL m.docname
 m.docname = STRTRAN(LOWER(ALLTRIM(parr)), '.xls', '')
 IF EMPTY(m.docname)
  RETURN .f.
 ENDIF 

 m.lIsWordActive = .t.
 TRY 
  oWord = GETOBJECT(,"Excel.Application")
 CATCH 
  m.lIsWordActive = .f.
 ENDTRY 
 IF m.lIsWordActive = .f.
  RETURN .f. 
 ENDIF 

 oDocs = oWord.WorkBooks
 nDocs = oDocs.count
 
 IF nDocs<=0
  RETURN .f. 
 ENDIF 
 
 m.lResult = .f.
 FOR EACH oDoc IN oDocs
  m.exname = STRTRAN(LOWER(oDoc.name), '.xls', '')
  IF m.exname=m.docname
   oDoc.Close(.f.)
   m.lResult = .t.
   EXIT 
  ENDIF 
 ENDFOR 
 
RETURN m.lResult

FUNCTION IsObr(m.codusl)
 LOCAL m.lIsObr, m.cod
 m.IsObr = .F.
 m.cod   = m.codusl
 IF (m.codusl<=2000 AND SUBSTR(PADL(m.codusl,6,'0'),6,1)='1') OR ;
  (BETWEEN(m.codusl,101001,102000) AND SUBSTR(PADL(m.codusl,6,'0'),6,1)='1')
  m.IsObr = .T.
 ENDIF 
RETURN m.IsObr

FUNCTION IsRowValid
RETURN .t.
FUNCTION IsAisValid
RETURN .t.
FUNCTION IsPplValid
RETURN .t.
FUNCTION IsTlnValid
RETURN .t.
FUNCTION uuser
RETURN .t.
FUNCTION comidx
RETURN .t.

* �������� ������ �� ��������� ��������
*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
DEFINE CLASS _scanner as custom
	*!!!* ������� ��������� ��� ������ ����������� �������.
	* ��������� ����������:
	on=.T.			&& ������ ������������
	commport= m.ComPort		&& COM-���� ���� �������� ������ ����� �����
*	commport= 4		&& COM-���� ���� �������� ������ ����� �����
	*         ^    @ ������ ������� � com port
	*^ ��������� ����������

	scancode=PADL('',17,'0')	&& ��������� �����
	onRun='this.parent.parent.onComm()'	&& ������ �� ��������� ���������
*	onRun='this.parent.onComm()'	&& ������ �� ��������� ���������
	*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE init
	ENDPROC
	*\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE activate		&& ���������
		IF !this.on
			RETURN .f.
		endIf
		if !(VARTYPE(this.scanner)='O')
			set Classlib To 'SCANCOM' additive	&& ��������� ���������� ������� �� ����� ����
			If !('SCANCOM.VCX'$Set("Classlib" ))
				wait WINDOW '!Set("Classlib" )'
				RETURN .f.
			ENDIF
			this.addObject('scanner','scancom.scanner_')
		endIf
		WITH This.scanner.olecomm
			* ��������� ���������� �� ����� ���������
			.commport=this.commport		&& COM-���� ���� �������� ������ ����� �����
			.EOFEnable=.T.
			.RTSEnable=.T.
			.RThreshold=1
			.Settings="9600,n,8,1"
			.SThreshold=0
			try
				.portOpen=.T.
			CATCH
			ENDTRY
			* ��������� ���������� �� ����� ���������
		endWith
	ENDPROC
	*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	*\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE off		&& ���������
		if (VARTYPE(this.scanner)='O')
			if	This.scanner.olecomm.portOpen=.T.
				This.scanner.olecomm.portOpen=.f.
			endIf
		endIf
	ENDPROC
	*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE destroy
		this.off()		&& ���������
	endProc
	*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE onComm
		WAIT WINDOW '����������� :'+this.scancode nowait
	endProc

ENDDEFINE

*SCANCOM.VCX

FUNCTION SecToHrs(para1) && �������� �������
 m.secs = ROUND(para1,0)
 m.rslt = ""
 
 m.days  = FLOOR(m.secs/(60*60*24))
 m.hours = FLOOR((m.secs - (m.days*(60*60*24)))/(60*60))
 m.mins  = FLOOR((m.secs - (m.days*(60*60*24)) - (m.hours*(60*60)))/60)
 m.ssecs = m.secs - m.mins*60 - m.hours*60*60 - m.days*60*60*24
 
 m.rslt = PADL(m.days,2,'0')+':'+PADL(m.hours,2,'0')+':'+PADL(m.mins,2,'0')+':'+PADL(m.ssecs,2,'0')

RETURN rslt && ���������� ��������� ���������� dd:hh:mm:ss

