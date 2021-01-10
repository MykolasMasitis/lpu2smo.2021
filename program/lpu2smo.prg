# DEFINE DEBUGMODE .F.
# DEFINE EXPMODE 0
# DEFINE SELMODE 1
*On Error Do ErrorHandler With Error( ), ;
							Message( ), ;
							Message(1), ;
							Program( ), ;
							Lineno(1)

ON SHUTDOWN DO ExitProg

RELEASE ALL EXTENDED
CLEAR ALL
SET TALK OFF
SET HOURS TO 24
SET DATE TO GERMAN
SET CENTURY ON 
SET CONSOLE OFF
SET RESOURCE OFF 
SET safety OFF 
SET REPROCESS TO 3 SECONDS 
SET DELETED ON 
SET ESCAPE OFF  

* Для функций ToBase64 FromBase64 в UploadMail.prg; спасла при работе с ВТБ!!!
DECLARE INTEGER CryptBinaryToString IN Crypt32;
	STRING @pbBinary, LONG cbBinary, LONG dwFlags,;
	STRING @pszString, LONG @pcchString

DECLARE INTEGER CryptStringToBinary IN crypt32;
	STRING @pszString, LONG cchString, LONG dwFlags,;
	STRING @pbBinary, LONG @pcbBinary,;
	LONG pdwSkip, LONG pdwFlags
* Для функций ToBase64 FromBase64 в UploadMail.prg; спасла при работе с ВТБ!!!

DECLARE INTEGER GetPrivateProfileString IN Win32API  AS GetPrivStr ;
	STRING cSection, STRING cKey, STRING cDefault, STRING @cBuffer, ;
	INTEGER nBufferSize, STRING cINIFile
DECLARE INTEGER WritePrivateProfileString IN Win32API AS WritePrivStr ;
	STRING cSection, STRING cKey, STRING cValue, STRING cINIFile
	
DECLARE INTEGER GetSysColor IN User32.DLL INTEGER
DECLARE INTEGER ShellExecute IN shell32.dll ;
	INTEGER hndWin, STRING cAction, STRING cFileName, ;
	STRING cParams, STRING cDir, INTEGER nShowWin

DECLARE ScreenSize In Tools32 ;
	Integer @nW, ;  && Ширина
	Integer @nH     && Высота

PUBLIC fso AS SCRIPTING.FileSystemObject, wshell AS Shell.Application

fso      = CREATEOBJECT('Scripting.FileSystemObject')
WShell   = CREATEOBJECT('Shell.Application')
WSHShell = CREATEOBJECT('Wscript.Shell')

SET PROCEDURE TO Utils
SET PROCEDURE TO Soap ADDITIVE 
SET PROCEDURE TO gpImage2.prg ADDITIVE 
SET PROCEDURE TO FoxBarCode.prg ADDITIVE 
SET PROCEDURE TO FoxBarCodeQR.prg ADDITIVE 
SET PROCEDURE TO lpu2sql.prg ADDITIVE 
SET PROCEDURE TO ContentHandler.prg ADDITIVE  && класс для сверки регистра

PUBLIC nWidth, nHeight, nDifSize, IsNotePad, m.ffoms
m.nWidth    = 0
m.nHeight   = 0
m.nDifSize  = 800-768

m.IsNotePad = .F.

=ScreenSize(@nWidth, @nHeight)

IF m.nHeight=768
_screen.WindowState = 2
ENDIF 

m.IsNotePad = IIF(m.nHeight<768, .T., .F.)
*m.IsNotePad = .T.

WITH _SCREEN
* .Width      = 1024
 .Width      = IIF(m.IsNotePad=.f., 1280, 1032)
 *.Height     = IIF(m.IsNotePad=.f., (800-m.nDifSize)-100, (600-m.nDifSize)-100)
 .Height     = IIF(m.IsNotePad=.f., (900-m.nDifSize)-100, (600-m.nDifSize)-100)
 .BackColor  = RGB(192,192,192)
 .AutoCenter = .t.
 .Picture    = 'lpu2smo.jpg'
 .Visible    = .t.
 .Icon = 'cross.ico'
ENDWITH 

lcPathSystem = sys(5) + sys(2003)
Set DEFAULT TO (lcPathSystem)
lcPathMain = sys(5) + sys(2003)
lcPathSystem = lcPathMain+;
	','+(lcPathMain+'\BITMAPS')+;
	','+(lcPathMain+'\FORMS')+;
	','+(lcPathMain+'\GRAPHICS')+;
	','+(lcPathMain+'\INCLUDE')+;
	','+(lcPathMain+'\LIBS')+;
	','+(lcPathMain+'\FOXCHARTS')+;
	','+(lcPathMain+'\DESKTOPALERTS')+;
	','+(lcPathMain+'\PARALLEL')+;
	','+(lcPathMain+'\MENU')+;
	','+(lcPathMain+'\PROGRAM')

SET PATH TO (lcPathSystem)

*LOCAL oMsg
*oMsg = NEWOBJECT("msgbox", "base")
*WITH oMsg
*	.SetCaption("Please wait...")
*	.Visible=.T.
*	.SetMessage(PRODUCT_NAME + CRLF + CRLF +;
*		COPYRIGHT_TEXT + CRLF +;
*		CONTACT_PERSON + CRLF +;
*		CONTACT_PHONE + CRLF +;
*		CONTACT_ADDRESS;
*		)
*ENDWITH

*RELEASE oMsg

PUBLIC paisoms, parc, pbase, pbin, pcommon, pout, ptempl, pmee, pexpimp, ;
 tyear, tmonth, tdat1, tdat2, curlpu, qcod, qmail, qobjid, UserERZ, qname, oMenu, gcPeriod, gcUser, gcFormat,;
 usrfam, usrim, usrot, usrfio, usrmail, supervisor, m.ynorm, IsUsrDir, IsServer, ComPort, plocal

PUBLIC m.y_app, m.y_st, m.y_dst, m.y_smp, m.y_vmp, m.y_prd && Нормативы для штрафов с 01.06.2019!

PUBLIC m.orgId, m.orgCode, m.soapSystem, m.soapUser, m.soapPass, m.pSoap
PUBLIC m.pumpUser, m.pumpPass, m.erzlUser, m.erzlPass, t_robot, t_start, ISPARALLEL, ISTESTMODE
m.soapSystem = 'lpu2smo'

PUBLIC m.LibVersion, m.IntDate
m.IntDate={04.09.2016}
m.LibVersion=2458119

PUBLIC m.SuperExp
m.SuperExp = 'EXP009'

PUBLIC m.IsTimerOn
m.IsTimerOn = .F.

* Переменная, переключающая режим начального определения прикрепления/страховой принадлежности/вектора сверки
* 0 - оставить так, как было представлено МО
* 1 - сверить с номерником на этапе формирования people (модуль MakePeople в CheckMail OneFlkSoap)
* 2 - сверка ЕРЗЛ
PUBLIC m.SaveInitPr
m.SaveInitPr = 1 && Сверка с номерником включена

m.ynorm = 0

PUBLIC ARRAY mes_text[12], mes_main[12]
mes_text[1]="января"
mes_text[2]="февраля"
mes_text[3]="марта"
mes_text[4]="апреля"
mes_text[5]="мая"
mes_text[6]="июня"
mes_text[7]="июля"
mes_text[8]="августа"
mes_text[9]="сентября"
mes_text[10]="октября"
mes_text[11]="ноября"
mes_text[12]="декабря"

mes_main[1]="Январь"
mes_main[2]="Февраль"
mes_main[3]="Март"
mes_main[4]="Апрель"
mes_main[5]="Май"
mes_main[6]="Июнь"
mes_main[7]="Июль"
mes_main[8]="Август"
mes_main[9]="Сентябрь"
mes_main[10]="Октябрь"
mes_main[11]="Ноябрь"
mes_main[12]="Декабрь"

numlib = adir(alib,lcPathMain+'\LIBS\*.vcx')
for i = 1 to numlib
	lcSetLibrary = lcPathMain+'\LIBS\' + alib(i,1)
	set classlib to (lcSetLibrary) additive
ENDFOR

SET CLASSLIB TO Parallel\ParallelFox ADDITIVE
SET CLASSLIB TO Parallel\WorkerMgr ADDITIVE

SET LIBRARY TO &lcPathMain\vfpzip.fll
*SET LIBRARY TO &lcPathMain\vfpexmapi.fll ADDITIVE

loFbc = CREATEOBJECT("FoxBarcodeQR")

** Проверка и создание фйла lpu2smo.cfg, soap.cfg
IF CfgBase() = -1
 =ExitProg()
ENDIF 
** Проверка и создание фйла lpu2smo.cfg, soap.cfg

DO CASE 
 CASE m.qcod = 'S6'
  m.ffoms = 77011
 CASE m.qcod = 'P2'
  m.ffoms = 77002
 CASE m.qcod = 'R4'
  m.ffoms = 77008
 CASE m.qcod = 'I3'
  m.ffoms = 77013
 CASE m.qcod = 'R2'
  m.ffoms = 77000
 CASE m.qcod = 'S7'
  m.ffoms = 77012
 CASE m.qcod = 'R8'
  m.ffoms = 77014
 CASE m.qcod = 'M4'
  m.ffoms = 77005
 CASE m.qcod = 'M1'
  m.ffoms = 77004
 OTHERWISE 
 
ENDCASE 

DO CASE 
 CASE m.ffoms = 77011
  m.LibVersion=2457789
 CASE m.ffoms = 77002
  m.LibVersion=2458028
  IF INT(VAL(SYS(1)))>m.LibVersion
   *=UpdateLibs() 
  ENDIF 
 CASE m.ffoms = 77008
  m.LibVersion=2458028
  IF INT(VAL(SYS(1)))>m.LibVersion
   *=UpdateLibs() 
  ENDIF 
 CASE m.ffoms = 77013
  m.LibVersion=2457836
 CASE m.ffoms = 77012
  m.LibVersion=2457836
 OTHERWISE 
 
ENDCASE 

IF !fso.FolderExists(pcommon)
 MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ИЛИ НЕДОСТУПНА'+CHR(13)+CHR(10)+'ДИРЕКТОРИЯ '+pcommon,0+16,'')
 =ExitProg()
ENDIF  

m.tdat1 = CTOD('01.'+PADL(tMonth,2,'0')+'.'+PADL(tYear,4,'0'))
m.tdat2 = GOMONTH(CTOD('01.'+PADL(tMonth,2,'0')+'.'+PADL(tYear,4,'0')),1)-1
m.gcPeriod = STR(tYear,4)+PADL(tMonth,2,'0')

DO CASE 
 CASE m.qcod == 'P2'
  m.qname = 'АО "МСК "УРАЛСИБ"'
  m.qmail = 'skpomed.msk.oms'
  m.qobjid = 3386
 CASE m.qcod == 'P3'
  m.qname = 'ООО МСО "ПАНАЦЕЯ" МОСКОВСКИЙ ФИЛИАЛ'
  m.qmail = 'panacea.msk.oms'
  m.qobjid = 5387
 CASE m.qcod == 'I3'
  m.qname = 'ООО СК "ИНГОССТРАХ-М"'
  m.qmail = 'ingos.msk.oms'
  m.qobjid = 5398
 CASE m.qcod == 'I1'
  m.qname = 'ООО МСК "ИКАР"'
  m.qmail = 'ikar.msk.oms'
  m.qobjid = 110
 CASE m.qcod == 'R4'
  m.qname = 'ООО "СТРАХОВАЯ МЕДИЦИНСКАЯ КОМПАНИЯ РЕСО-МЕД" МОСКОВСКИЙ ФИЛИАЛ'
  m.qmail = 'reso.msk.oms'
  m.qobjid = 3415
 CASE m.qcod == 'S7'
  m.qname = 'АО СК "СОГАЗ-Мед"'
  m.qmail = 'sogaz.msk.oms'
  m.qobjid = 5400
 CASE m.qcod == 'S2'
  m.qname = 'АО ВТБ Медицинское страхование'
  m.qmail = 'sovita.msk.oms'
  m.qobjid = 33
 CASE m.qcod == 'R2'
  m.qname = 'ООО ВТБ Медицинское страхование'
  m.qmail = 'sovita.msk.oms'
  *m.qobjid = 3522
  m.qobjid = 111
 CASE m.qcod == 'S6'
  m.qname = 'ЗАО СК "СОГЛАСИЕ-М"'
  m.qmail = 'soglasie.msk.oms'
  m.qobjid = 5404
 CASE m.qcod == 'R8'
  m.qname = 'РГС'
  m.qmail = 'rgs.msk.oms'
  m.qobjid = 6469
 CASE m.qcod == 'M4'
  m.qname = 'ООО МСК "МЕДСТРАХ"'
  m.qmail = 'medstrah.msk.oms'
  m.qobjid = 124
 CASE m.qcod == 'M1'
  m.qname = 'АО "МАКС-М"'
  m.qmail = 'maksm.msk.oms'
  m.qobjid = 124
 OTHERWISE 
  m.qname = 'ОАО "МСК "УРАЛСИБ"'
  m.qmail = 'skpomed.msk.oms'
  m.qobjid = 3386
ENDCASE 

public goApp
goApp = NEWOBJECT('_goapp','main')

ADDPROPERTY(goApp, "regim", 0)
ADDPROPERTY(goApp, "mcod", "")
ADDPROPERTY(goApp, "filial", "")
ADDPROPERTY(goApp, "vfilter", "")
ADDPROPERTY(goApp, "keypressed", "")
ADDPROPERTY(goApp, "tipacc", 0)
ADDPROPERTY(goApp, "flcod", "")
ADDPROPERTY(goApp, "Aisoms", "")
ADDPROPERTY(goApp, "People", "")
ADDPROPERTY(goApp, "Talon", "")
ADDPROPERTY(goApp, "eError", "")
ADDPROPERTY(goApp, "mError", "")
ADDPROPERTY(goApp, "lcDir", "")
ADDPROPERTY(goApp, "pPath", "")
*ADDPROPERTY(goApp, "etap", " ")
ADDPROPERTY(goApp, "etap", "2")
ADDPROPERTY(goApp, "docexp", "")
ADDPROPERTY(goApp, "nrecid", 0)
ADDPROPERTY(goApp, "supexp", SPACE(7))
ADDPROPERTY(goApp, "smoexp", SPACE(7))
ADDPROPERTY(goApp, "profil", SPACE(3))
ADDPROPERTY(goApp, "tiplpu", 0)
ADDPROPERTY(goApp, "tipacc", 0)
ADDPROPERTY(goApp, "exporsel", EXPMODE)
ADDPROPERTY(goApp, "d_exp", DATE())
*ADDPROPERTY(goApp, "d_acts", DATE())
ADDPROPERTY(goApp, "d_acts", GoNWrkDays(GOMONTH(m.tdat1,1), 6))
ADDPROPERTY(goApp, "mp", " ")
ADDPROPERTY(goApp, "reason", "0")
ADDPROPERTY(goApp, "theme", " ")
ADDPROPERTY(goApp, "vp", 0)
ADDPROPERTY(goApp, "callForm", "")
ADDPROPERTY(goApp, "recid_lpu", "")
ADDPROPERTY(goApp, "recid_sl", "")
ADDPROPERTY(goApp, "recid_usl", "")

PUBLIC ReindTimer
ReindTimer=CREATEOBJECT("Timer")

*goApp.Show()
goApp.Begin_process()

** Создание директорий и НСИ!
=chkbase()
** Создание директорий и НСИ!

IF !fso.FileExists(pCommon+'\Users.cdx')
 IF OpenFile(pcommon+'\Users', 'users', 'excl') <= 0
  SELECT Users 
  INDEX ON name TAG name 
  USE 
 ENDIF 
ENDIF 

IF !fso.FileExists(pCommon+'\pnyear.dbf')
 MESSAGEBOX('ФАЙЛ '+pCommon+'\pnyear.dbf'+' НЕ НАЙДЕН!',0+64,'')
 =ExitProg()
ELSE 
 IF OpenFile(pCommon+'\pnyear', 'pnyear', 'shar', 'period')>0
  IF USED('pnyear')
   USE IN pnyear
  ENDIF 
 ELSE 
  SELECT pnyear
  IF SEEK(m.gcperiod, 'pnyear')
   m.ynorm = pnyear.pnorm
   
  * m.y_app = pnyear.app
  * m.y_st  = pnyear.st
  * m.y_dst = pnyear.dst
  * m.y_smp = pnyear.smp
  * m.y_vmp = pnyear.vmp
  * m.y_prd = pnyear.prd
  ENDIF 
  IF USED('pnyear')
   USE IN pnyear
  ENDIF 
 ENDIF 
ENDIF 

=OpenFile(pCommon+'\Users', 'Users', 'shar', 'name')
SELECT Users
IF !SEEK(m.gcUser, 'Users')
 USE 
 MESSAGEBOX('ИМЯ '+ALLTRIM(m.gcUser)+' ОТСУТСТВУЕТ В СПРАВОЧНИКЕ!', 0+16, '')
 =ExitProg()
ELSE 
 IF !RLOCK()
  USE 
  MESSAGEBOX('ПОЛЬЗОВАТЕЛЬ '+ALLTRIM(m.gcUser)+' УЖЕ РАБОТАЕТ В СИСТЕМЕ!', 0+16, '')
  =ExitProg()
 ELSE 
  m.usrfam  = ALLTRIM(Fam)
  m.usrim   = ALLTRIM(im)
  m.usrot   = ALLTRIM(ot)
  m.usrfio  = ALLTRIM(fio)
  m.usrmail = ALLTRIM(usrmail)
*  m.supervisor = super
 ENDIF 
ENDIF 

goApp.smoexp = m.gcUser
** Проверяем зарегистрированы ли библиотеки
IF !IsRegistered("TTF161.TTF1.6")
 MESSAGEBOX('Зарегистрируйте ttf16.ocx!'+CHR(13)+CHR(10)+;
 	'(regsvr32 ttf16.ocx)', 0+64, 'ttf16.ocx')
ENDIF 
IF !IsRegistered("vfpalert.AlertManager")
 MESSAGEBOX('Зарегистрируйте vfpAlert.exe!'+CHR(13)+CHR(10)+;
 	'(vfpAlert.exe /regserver)', 0+64, 'vfpAlert.exe')
ENDIF 
IF !IsRegistered("ParallelFox.Application")
 MESSAGEBOX('Зарегистрируйте ParallelFox.exe!'+CHR(13)+CHR(10)+;
 	'(ParallelFox.exe /regserver)',0+64,'ParallelFox.exe')
ENDIF 
IF !IsRegistered("Excel.Application")
 MESSAGEBOX('На компьютере не установлен Excel'+CHR(13)+CHR(10)+;
 	'Формирование документов будет невозмжно',0+64,'Excel')
ENDIF 
IF !IsRegistered("MSComCtl2.MonthView.2")
 MESSAGEBOX('Зарегистрируйте MSCOMCT2.OCX!'+CHR(13)+CHR(10)+;
 	'(regsvr32 MSCOMCT2.OCX)'+CHR(13)+CHR(10)+;
 	'(В 64-битных версиях путь c:\windows\sysWOW64)',0+64,'MSCOMCT2.OCX')
ENDIF 
IF !IsRegistered("MSCOMMLib.MSComm.1")
 MESSAGEBOX('Зарегистрируйте MSCOMM32.OCX!'+CHR(13)+CHR(10)+;
 	'(regsvr32 MSCOMM32.OCX)'+CHR(13)+CHR(10)+;
 	'(В 64-битных версиях путь c:\windows\sysWOW64)',0+64,'MSCOMM32.OCX')
ENDIF 
IF !IsRegistered("Msxml2.MXXMLWriter.6.0")
 MESSAGEBOX('Отсутсвует библиотека MSXML6!'+CHR(13)+CHR(10)+;
 	'(В 32-битных версиях путь c:\windows\system32)'+CHR(13)+CHR(10)+;
 	'(В 64-битных версиях путь c:\windows\sysWOW64)',0+64,'MsXML6')
ENDIF 

IF !IsRegistered("MsXml2.XMLHTTP.3.0")
 MESSAGEBOX('Отсутсвует библиотека MSXML3!'+CHR(13)+CHR(10)+;
 	'(В 32-битных версиях путь c:\windows\system32)'+CHR(13)+CHR(10)+;
 	'(В 64-битных версиях путь c:\windows\sysWOW64)',0+64,'MsXML3')
ENDIF 

IF !IsRegistered("MsXml2.XMLHTTP.6.0")
 MESSAGEBOX('Отсутсвует библиотека MSXML6!'+CHR(13)+CHR(10)+;
 	'(В 32-битных версиях путь c:\windows\system32)'+CHR(13)+CHR(10)+;
 	'(В 64-битных версиях путь c:\windows\sysWOW64)',0+64,'MsXML6')
ENDIF 
** Проверяем зарегистрированы ли библиотеки

SET SYSMENU TO
SET SYSMENU ON
SET STATUS BAR ON 
WITH _SCREEN
 .Caption = m.qname+', ПОЛЬЗОВАТЕЛЬ: '+ALLTRIM(m.gcUser)+', ПЕРИОД: '+NameOfMonth(tMonth)+' '+STR(tYear,4)+' ГОДА'+;
  ' (с '+DTOC(tdat1)+' по '+DTOC(tdat2)+')'
ENDWITH 

PUBLIC IsAdmin
m.IsAdmin=.f.
IF fso.FileExists(pbin+'\admin')
 ffile = fso.GetFile(pbin+'\admin')
  IF ffile.size == 4
   fhandl = ffile.OpenAsTextStream
   lcHead = fhandl.Read(4)
   fhandl.Close
   IF lcHead == 'ruby'
    m.IsAdmin = .t.
   ENDIF 
  ENDIF 
ENDIF 

DO Base64_ini IN base64

DO System.APP

IF M.ISTESTMODE
MESSAGEBOX('ВКЛЮЧЕН ТЕСТОВЫЙ РЕЖИМ'+CHR(13)+CHR(10)+'ОТВЕТЫ В МО НЕ ОТПРАВЛЯЮТСЯ!',0+64,'test-mode')
ENDIF 

*DO demomenu1
DO m_menu
READ EVENTS

=ExitProg()

FUNCTION ExitProg
 IF USED('Users')
  USE IN Users
 ENDIF 
 RELEASE m.oError
 _SCREEN.Caption = ""
 RELEASE m.goApp
 ON SHUTDOWN
 QUIT
RETURN 
