FUNCTION MakeEkmp(lcpolis, lcPath, IsVisible, IsQuit, TipAcc, TipOfExp)
 DotName = 'zakl_ekmp.xls'
 IF !fso.FileExists(pTempl+'\'+DotName)
  MESSAGEBOX('ОТСУТСТВУЕТ ШАБЛОН '+CHR(13)+CHR(10)+UPPER(m.dotname),0+32,'')
  RETURN 
 ENDIF
 
 m.llcpolis = lcpolis
 
 =MakeEkmp0(m.llcpolis, lcPath, IsVisible, IsQuit, TipAcc, '4')
 =MakeEkmp0(m.llcpolis, lcPath, IsVisible, IsQuit, TipAcc, '5')
 =MakeEkmp0(m.llcpolis, lcPath, IsVisible, IsQuit, TipAcc, '6')
 =MakeEkmp0(m.llcpolis, lcPath, IsVisible, IsQuit, TipAcc, '9')

RETURN 

FUNCTION MakeEkmp0(lcpolis, lcPath, IsVisible, IsQuit, TipAcc, TipOfExp)

 m.TipOfExp = TipOfExp
 DO CASE 
  CASE m.TipOfExp = '4'
   m.ctipofexp = 'Плановая ЭКМП'
   m.vidofexp = 'Плановая ЭКМП'
  CASE m.TipOfExp = '5'
   m.ctipofexp = 'Целевая ЭКМП'
   m.vidofexp = 'Целевая ЭКМП'
  CASE m.TipOfExp = '6'
   m.ctipofexp = 'Тематическая ЭКМП'
   m.vidofexp = 'Тематическая ЭКМП'
  CASE m.TipOfExp = '9'
   m.ctipofexp = 'ЭКМП по жалобе'
   m.vidofexp = 'ЭКМП по жалобе'
  OTHERWISE 
   m.ctipofexp = ''
   m.vidofexp = ''
 ENDCASE 
 
 oal  = ALIAS()
 orec = RECNO()
 
 CREATE CURSOR curpaz (sn_pol c(25)) 
 INDEX on sn_pol TAG sn_pol 
 SET ORDER TO sn_pol

 SELECT &oal
 m.nlocked = 0 
 SCAN 
  IF !ISRLOCKED()
   LOOP 
  ENDIF 
  m.sn_pol = sn_pol
  IF !SEEK(m.sn_pol, 'curpaz')
   INSERT INTO curpaz FROM MEMVAR 
   m.nlocked = m.nlocked + 1
  ENDIF 
 ENDSCAN 
 GO (orec)
 
 m.IsMulti=0
 IF m.nlocked>0
  IF MESSAGEBOX('ФОРМИРОВАТЬ АКТЫ НА ВСЕХ ОТОБРАННЫХ?'+CHR(13)+CHR(10),4+32,'')=6
   m.IsMulti=1
  ELSE
   m.IsMulti=0
  ENDIF 
 ENDIF 

 m.docfio  = m.usrfam+' '+m.usrim+' '+m.usrot

 m.mcod       = goApp.mcod 
 m.lpuid      = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)
 m.lpuaddress = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.adres), '')
 m.lpuname    = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.name)+', '+m.mcod, '')

 IF m.IsMulti=1
  SELECT curpaz 
  SCAN 
*   m.lpolis = lcpolis
   m.lpolis = sn_pol
   =MakeEkmp1(lpolis, lcPath, IsVisible, IsQuit, TipAcc, TipOfExp)
  ENDSCAN 
 ELSE 
  m.lpolis = lcpolis
  =MakeEkmp1(lpolis, lcPath, IsVisible, IsQuit, TipAcc, TipOfExp)
 ENDIF 
 
 SELECT &oal
 GO (orec) 

RETURN 
 
FUNCTION MakeEkmp1(lpolis, lcPath, IsVisible, IsQuit, TipAcc, TipOfExp)

 oal2 = ALIAS()
 m.sn_pol = lpolis

 CREATE CURSOR curdoc (docexp c(7))
 INDEX on docexp TAG docexp 
 SET ORDER TO docexp 
 
 SELECT merror
 m.nIsDocs = 0 
 SCAN 
  IF SEEK(merror.recid, 'talon', 'recid') AND talon.sn_pol = m.sn_pol AND merror.et = m.TipOfExp
   m.nIsDocs = m.nIsDocs + 1
   m.docexp = docexp 
   IF !SEEK(m.docexp, 'curdoc')
    INSERT INTO curdoc FROM MEMVAR 
   ENDIF 
  ENDIF 
 ENDSCAN 

 m.fiopaz = IIF(SEEK(m.sn_pol, 'people', 'sn_pol'), ALLTRIM(people.fam)+' '+ALLTRIM(people.im)+' '+ALLTRIM(people.ot), {})
 m.d_beg = IIF(SEEK(m.sn_pol, 'people', 'sn_pol'), people.d_beg, {})
 m.d_end = IIF(SEEK(m.sn_pol, 'people', 'sn_pol'), people.d_end, {})
 m.sex   = IIF(SEEK(m.sn_pol, 'people', 'sn_pol'), IIF(people.w==1,'мужской','женский'), '')
 m.dr    = IIF(SEEK(m.sn_pol, 'people', 'sn_pol'), DTOC(people.dr), '')

 IF RECCOUNT('curdoc')>0
  SELECT curdoc
  SCAN 
   m.docexp = docexp
   =MakeEkmp11(lpolis, lcPath, IsVisible, IsQuit, TipAcc, TipOfExp, m.docexp)
  ENDSCAN 
 ELSE 
 ENDIF 
 USE IN curdoc 
 SELECT (oal2)

RETURN 

FUNCTION MakeEkmp11(lpolis, lcPath, IsVisible, IsQuit, TipAcc, TipOfExp, para7) && para7 - m.docexp

 oal3 = ALIAS()
 m.sn_pol = lpolis

 CREATE CURSOR currsns (reason c(1))
 INDEX on reason TAG reason
 SET ORDER TO reason
 
 SELECT merror
 m.nIsReasons = 0 
 SCAN 
  IF SEEK(merror.recid, 'talon', 'recid') AND talon.sn_pol = m.sn_pol AND merror.et = m.TipOfExp AND merror.docexp=m.para7
   m.nIsReasons = m.nIsReasons + 1
   m.reason = reason
   IF !SEEK(m.reason, 'currsns')
    INSERT INTO currsns FROM MEMVAR 
   ENDIF 
  ENDIF 
 ENDSCAN 

 IF RECCOUNT('currsns')>0
  SELECT currsns
  SCAN 
   m.reason = reason
   =MakeEkmp111(lpolis, lcPath, IsVisible, IsQuit, TipAcc, TipOfExp, m.docexp, m.reason)
  ENDSCAN 
 ELSE 
 ENDIF 
 USE IN currsns
 SELECT (oal3)

RETURN 

FUNCTION MakeEkmp111(lpolis, lcPath, IsVisible, IsQuit, TipAcc, TipOfExp, para7, para8) && para8 - m.docexp
 
 oal4 = ALIAS()
 m.sn_pol = lpolis
 m.ddocexp = para7
 m.rreason = para8

 CREATE CURSOR ttalon (recid i, sn_pol c(25), c_i c(25), docotd c(100), ds c(6), d_in d, d_u d, pcod c(10),;
  otd c(4), cod n(6), k_u n(3), d_type c(1), s_all n(11,2), err_mee c(3), ername c(250), docexp c(7), ishod n(3), is_name c(60), ;
  d_beg d, d_end d, reason c(1))
 INDEX on recid TAG recid
 SET ORDER TO recid 
 
 CREATE CURSOR curds (ds c(6), dsname c(160))
 INDEX on ds TAG ds
 SET ORDER TO ds

 CREATE CURSOR curdefs (er_c c(2), osn230 c(5), ername c(100))
 INDEX on er_c TAG er_c
 SET ORDER TO er_c

 SELECT merror

 m.TipOfMee = 0
 
 m.s_lech = 0
 m.dat1   = {31.12.2099}
 m.dat2   = {01.01.2000}
 m.dslast = 0
 m.totdefs = 0
 m.nekoplate = 0
 m.tot_straf = 0

 SCAN

  IF SEEK(recid, 'talon', 'recid') AND talon.sn_pol != m.sn_pol
   LOOP 
  ENDIF 
  IF USED('serror')
   IF SEEK(recid, 'serror')
    LOOP 
   ENDIF 
  ENDIF 
  IF et!=m.TipOfExp
   LOOP 
  ENDIF 
  IF docexp!=m.ddocexp
   LOOP 
  ENDIF 
  IF reason!=m.rreason
   LOOP 
  ENDIF 
  
  m.povod = IIF(INLIST(m.TipOfExp,'3','5'), IIF(SEEK(m.rreason, 'reasons'), reasons.name, ''), '')
 
  m.c_i    = talon.c_i
  m.s_all  = talon.s_all
  m.e_sall = 0
  m.s_lech = m.s_lech + m.s_all
  m.ds     = talon.ds
  m.dsname = IIF(SEEK(m.ds, 'mkb10'), mkb10.name_ds, '')
  m.d_u    = talon.d_u
  m.k_u    = talon.k_u
  
  m.ishod   = talon.ishod
  m.is_name = IIF(SEEK(m.ishod, 'isv', 'ishod'), isv.is_name, '')

  m.recid    = recid
  m.err_mee  = err_mee 
  m.err_mee  = LEFT(err_mee,2)
  m.er_c     = m.err_mee
  m.osn230   = osn230
  m.ername   = IIF(SEEK(m.err_mee,'errsmee'), ALLTRIM(errsmee.f_komment), '')
  m.s_1      = s_1
  m.s_2      = s_2
  
  m.totdefs = m.totdefs + IIF(m.er_c!='W0', 1, 0)
  m.nekoplate = m.nekoplate + m.s_1
  m.tot_straf = m.tot_straf + m.s_2

  m.cod        = cod 
  m.otd        = talon.otd
  m.otdname    = IIF(USED('otdel'), IIF(SEEK(m.otd,'otdel'), ALLTRIM(otdel.name), ''), '')
  m.pcod       = talon.pcod 
  m.doctorname = IIF(USED('doctor'), IIF(SEEK(m.pcod, 'doctor'), PROPER(ALLTRIM(doctor.fam))+' '+;
   PROPER(ALLTRIM(doctor.im))+' '+PROPER(ALLTRIM(doctor.ot))+', '+pcod, ''), '')

  IF IsUsl(m.cod)
   m.docotd = m.doctorname
  ELSE 
   m.docotd = m.otdname
  ENDIF 

  DO CASE 
   CASE IsMes(m.cod) OR IsVMP(m.cod)
    m.d_in = IIF(m.k_u>1, m.d_u-m.k_u-1, m.d_u-1)
   CASE IsKD(m.cod)
   CASE IsUsl(m.cod)
   OTHERWISE 
  ENDCASE 

  DO CASE 
   CASE m.TipAcc == 0 && Сводный счет
    m.dat1 = m.d_beg
    m.dat2 = m.d_end
   CASE m.TipAcc == 1 && Амбулаторный счет
    m.dat1 = MIN(m.d_u, m.dat1)
    m.dat2 = MAX(m.d_u, m.dat2)
   CASE m.TipAcc == 2 && Дневной стационар
    m.dat1 = MIN(m.d_u-m.k_u, m.dat1)
    m.dat2 = MAX(m.d_u, m.dat2)
    m.dslast = m.dslast + k_u
   CASE m.TipAcc == 3 && Стационар
    m.dat1 = MIN(m.d_u-m.k_u+1, m.dat1)
    m.dat2 = MAX(m.d_u, m.dat2)
    m.dslast = m.dslast + m.k_u
  ENDCASE 

  DO CASE 
   CASE EMPTY(m.err_mee)
   CASE LEFT(m.err_mee,2)=='W0'
    m.TipOfMee = IIF(m.TipOfMee<=1, 1, m.TipOfMee)
   OTHERWISE 
    m.TipOfMee = 2 && Есть ошибки
     m.e_sall = s_1
  ENDCASE 

  IF !SEEK(m.recid, 'ttalon')
   INSERT INTO ttalon FROM MEMVAR 
  ENDIF 

  IF !SEEK(m.ds, 'curds')
   INSERT INTO curds FROM MEMVAR 
  ENDIF 

  IF !SEEK(m.err_mee, 'curdefs')
   INSERT INTO curdefs FROM MEMVAR 
  ENDIF 

 ENDSCAN 

 m.fioexp   = IIF(SEEK(m.ddocexp, 'explist'), ;
  ALLTRIM(explist.fam)+' '+ALLTRIM(explist.im)+' '+ALLTRIM(explist.ot), '')
 m.fioexp2   = IIF(SEEK(m.ddocexp, 'explist'), ;
  ALLTRIM(explist.fam)+' '+LEFT(ALLTRIM(explist.im),1)+'.'+LEFT(ALLTRIM(explist.ot),1)+'.', '')

 m.TipOfMee = IIF(m.TipOfMee==0, 1 ,m.TipOfMee)
 
 SELECT people

 DO CASE 
  CASE m.TipAcc == 0
   m.AddToName = 'св'
   m.tipakt = 'свод'
  CASE m.TipAcc == 1
   m.AddToName = 'амб'
   m.tipakt = 'амб'
  CASE m.TipAcc == 2
   m.AddToName = 'дст'
   m.tipakt = 'дн/стац'
  CASE m.TipAcc == 3
   m.AddToName = 'ст'
   m.tipakt = 'стац'
 ENDCASE 

 ooal = ALIAS()
 SELECT recid FROM ssacts WHERE period=m.gcperiod AND mcod=m.mcod AND codexp=INT(VAL(m.TipOfExp)) AND ;
  tipacc=m.tipacc AND sn_pol=PADR(STRTRAN(m.sn_pol,' ',''),25) AND IsOk=IIF(m.TipOfMee=1, .t., .f.) AND doctyp='XLS' ;
  AND docexp=m.ddocexp INTO CURSOR rqwest NOCONSOLE 
 m.nfileid = recid
 USE 
 SELECT (ooal)
 
 IF m.nfileid>0
  m.DocName = pmee+'\ssacts\'+PADL(m.nfileid,6,'0')
 ELSE 
  INSERT INTO ssacts (period,doctyp,mcod,codexp,tipacc,isok,sn_pol,fam,im,ot,docexp,doctyp,reason) ;
   VALUES ;
  (m.gcperiod,'DOC',m.mcod,INT(VAL(m.tipofexp)),m.tipacc,IIF(m.TipOfMee=1, .t., .f.),;
   PADR(STRTRAN(m.sn_pol,' ',''),25),people.fam,people.im,people.ot,m.ddocexp,'XLS',goApp.reason)
  m.nfileid = GETAUTOINCVALUE()
  m.DocName = pmee+'\ssacts\'+PADL(m.nfileid,6,'0')
  UPDATE ssacts SET actname=PADL(m.nfileid,6,'0')+'.xls', actdate=DATETIME() WHERE recid = m.nfileid
 ENDIF 
 
 IF fso.FileExists(m.DocName+'.xls')
  oFile = fso.GetFile(m.DocName+'.xls')
  DateCreated      = TTOC(oFile.DateCreated)
  DateLastAccessed = TTOC(oFile.DateLastAccessed)
  DateLastModified = TTOC(oFile.DateLastModified)
  RELEASE oFile
  
  IF  m.IsMulti=0

  IF m.IsVisible=.t.
  IF MESSAGEBOX('ПО ВЫБРАННОМУ СЧЕТУ АКТ УЖЕ ФОРМИРОВАЛСЯ!'+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ДАТА СОЗДАНИЯ АКТА            : '+m.DateCreated+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ДАТА ПОСЛЕДНЕГО ОТКРЫТИЯ АКТА : '+m.DateLastAccessed+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ДАТА ПОСЛЕДНЕГО ИЗМЕНЕНИЯ АКТА: '+m.DateLastModified+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ВЫ ХОТИТЕ ПЕРЕФОРМИРОВАТЬ АКТ?',4+32,'') == 7 

   USE IN ttalon
   USE IN curds
   USE IN curdefs
   RETURN
  ELSE 

   UPDATE ssacts SET actdate=DATETIME() WHERE recid = m.nfileid
   
  ENDIF 
  ELSE 
   UPDATE ssacts SET actdate=DATETIME() WHERE recid = m.nfileid
  ENDIF 

  ELSE 

   UPDATE ssacts SET actdate=DATETIME() WHERE recid = m.nfileid

  ENDIF 

 ENDIF 

 m.n_akt = mcod + m.qcod + PADL(m.tMonth,2,'0') + RIGHT(STR(tYear,4),1)+'/'+ALLTRIM(STR(m.nfileid))
 m.d_akt =  IIF(m.qcod!='I3', DTOC(DATE()), '')
 m.d_exp = {}
 m.zakl_name01 = 'Экспертное заключение №'+m.n_akt+IIF(m.qcod!='P2',' от '+m.d_akt,'')
 m.zakl_name02 = '(протокол оценки качества медицинской помощи, '+m.ctipofexp+')'
 m.expname = 'Акт '
 DO CASE 
  CASE m.tipofexp = '4'
   m.expname = m.expname + 'плановой'
  CASE m.tipofexp = '5'
   m.expname = m.expname + 'целевой'
  CASE m.tipofexp = '6'
   m.expname = m.expname + 'тематической'
  CASE m.tipofexp = '9'
   m.expname = m.expname + 'целевой (по жалобе)'
 ENDCASE 
 m.expname = m.expname + ' экспертизы качества медицинской помощи'
 m.n_schet = STR(tYear,4)+PADL(m.tMonth,2,'0')
 m.dslast = IIF(!INLIST(m.TipAcc,2,3), IIF(m.dat2-m.dat1>0, m.dat2-m.dat1, 1), m.dslast)
* m.ds = talon.ds
* m.dsnam = IIF(SEEK(m.ds, 'mkb10'), ALLTRIM(mkb10.name_ds), '')
 m.pcod = talon.pcod
 m.docfam = IIF(USED('doctor'), IIF(SEEK(m.pcod, 'doctor'), ALLTRIM(doctor.Fam)+' '+ALLTRIM(doctor.Im)+' '+ALLTRIM(doctor.Ot), ''), '')

* oDoc.Bookmarks('ds').Select  
* oWord.Selection.TypeText(m.ds+', '+m.dsnam)

 IF m.TipOfMee == 1 && Все хорошо!
  m.resume = 'ДЕФЕКТОВ ОФОРМЛЕНИЯ ПЕРВИЧНОЙ МЕДИЦИНСКОЙ ДОКУМЕНТАЦИИ, НАРУШЕНИЙ ПРИ ОКАЗАНИИ МЕДИЦИНСКОЙ ПОМОЩИ НЕ ВЫЯВЛЕНО.'    
 ELSE && Есть ошибки!
  m.resume = REPLICATE('_',255)
 ENDIF 
 
 SELECT curds 
 IF RECCOUNT()<=1
 ELSE 
  SET ORDER TO 
  SCAN 
   IF ds==m.ds
    LOOP 
   ENDIF 
*   m.dsname = ds+', '+ALLTRIM(m.dsname)
*   oDoc.Tables(2).Cell(nRow,1).Select
*   oWord.Selection.InsertRows
*   oWord.Selection.TypeText(ds+', '+ALLTRIM(dsname))
  ENDSCAN 
 ENDIF 
* USE IN curds 
 
* oDoc.Bookmarks('diagname').Select  
* oWord.Selection.TypeText(m.dsname)

 m.diagname = m.dsname && !!!
* m.diagname = m.dsnam
* m.c_i='xxx'
* m.is_name = 'yyy'
* m.doctorname='doctorname'
* m.dsname = 'dsname'
 
 m.punkt1 = IIF(m.TipOfMee=1 and m.qcod!='P2','','')
 m.punkt2 = IIF(m.TipOfMee=1 and m.qcod!='P2','НЕТ','')
 m.punkt3 = IIF(m.TipOfMee=1 and m.qcod!='P2','БЕЗ ЗАМЕЧАНИЙ','')
 m.punkt4 = IIF(m.TipOfMee=1 and m.qcod!='P2','нет','')
 m.punkt5 = IIF(m.TipOfMee=1 and m.qcod!='P2','БЕЗ ЗАМЕЧАНИЙ','')
 m.punkt6 = IIF(m.TipOfMee=1 and m.qcod!='P2','НЕТ','')
 m.resume = IIF(m.qcod!='P2',m.resume,'')

 LOCAL m.lcTmpName, m.lcRepName, m.lcDbfName, m.llResult
 m.lcTmpName = pTempl+'\'+m.dotname
 m.lcRepName = m.docname+'.xls'
 m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)

 USE IN ttalon
 USE IN curds
 USE IN curdefs
 SELECT (oal4)

 WAIT CLEAR 

RETURN 
