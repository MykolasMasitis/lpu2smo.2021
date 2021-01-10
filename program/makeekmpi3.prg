FUNCTION MakeEkmpI3(lcpolis, lcPath, IsVisible, IsQuit, TipAcc, TipOfExp)

 m.usrdir = fso.GetParentFolderName(pbin) + '\'+UPPER(m.gcuser)

 DotNameWW = 'ActEKMPss.xls'
 IF !fso.FileExists(pTempl+'\'+DotNameWW)
  MESSAGEBOX('ОТСУТСТВУЕТ ШАБЛОН '+CHR(13)+CHR(10)+UPPER(m.dotnameww),0+32,'')
  RETURN 
 ENDIF
 DotNameW0 = 'ActEKMPssW0.xls'
 IF !fso.FileExists(pTempl+'\'+DotNameW0)
  MESSAGEBOX('ОТСУТСТВУЕТ ШАБЛОН '+CHR(13)+CHR(10)+UPPER(m.dotnamew0),0+32,'')
  RETURN 
 ENDIF
 
 m.llcpolis = lcpolis
 m.IsOkAct = .f.
 
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
 m.lpuboss    = IIF(SEEK(m.lpuid, 'lpudogs'), ALLTRIM(lpudogs.boss), '')

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
 m.nqwert = 0 
 COUNT FOR SEEK(merror.recid, 'talon', 'recid') AND talon.sn_pol = m.sn_pol AND merror.et = m.TipOfExp AND ;
 	merror.err_mee != 'W0' AND IIF(goApp.smoexp<>m.SuperExp, usr=goApp.smoexp, .T.) TO m.nqwert
 IF m.nqwert<=0
*  USE IN curdoc 
*  SELECT (oal2)
*  RETURN 
 ENDIF 

 m.nIsDocs = 0 
 SCAN 
  IF SEEK(merror.recid, 'talon', 'recid') AND talon.sn_pol = m.sn_pol AND merror.et = m.TipOfExp ;
  	AND IIF(goApp.smoexp<>m.SuperExp, usr=goApp.smoexp, .T.)
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
 m.nqwert = 0 

 COUNT FOR SEEK(merror.recid, 'talon', 'recid') AND talon.sn_pol = m.sn_pol AND merror.et = m.TipOfExp AND ;
 	merror.docexp=m.para7 AND IIF(m.TipOfExp != '5', merror.err_mee != 'W0', 1=1) TO m.nqwert

 IF m.nqwert<=0
*  USE IN currsns
*  SELECT (oal3)
*  RETURN 
  m.IsOkAct = .t.
  m.dotname = m.dotnamew0
 ELSE 
  m.dotname = m.dotnameww
 ENDIF 

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

 m.n_rss = 0
 m.vvir = m.mcod + DTOS(goApp.d_exp)
 IF !SEEK(m.vvir, 'rss')
  m.e_period = STR(YEAR(goApp.d_exp),4)+PADL(MONTH(goApp.d_exp),2,'0')
  INSERT INTO rss (lpu_id,mcod,d_u,e_period,smoexp,k_acts) VALUES (m.lpuid,m.mcod,goApp.d_exp,m.e_period,goApp.smoexp,1)
  m.n_rss = GETAUTOINCVALUE()
 ELSE 
  m.n_rss    = rss.recid
  m.e_period = rss.e_period
  UPDATE rss SET k_acts = k_acts+1 WHERE recid=m.n_rss
 ENDIF 

* m.vvir   = m.mcod + m.gcperiod + m.TipOfExp + m.reason
 *m.vvir   = m.mcod + m.gcperiod
 m.vvir   = goApp.smoexp + m.mcod + m.gcperiod + m.TipOfExp
 m.n_rqst = IIF(SEEK(m.vvir, 'rqst'), rqst.recid, 0)
 IF m.n_rqst>0
  m.rqstfile = PADL(m.n_rqst,6,'0')
  IF fso.FileExists(pmee+'\requests\'+m.rqstfile+'.dbf')
   m.rtal = ALIAS()
   IF OpenFile(pmee+'\requests\'+m.rqstfile, 'rrfile', 'shar', 'sn_pol')>0
    IF USED('rrfile')
     USE IN rrfile
    ENDIF 
   ENDIF 
   SELECT (rtal)
  ENDIF 
 ENDIF 
 IF fso.FileExists(pmee+'\ssacts\moves'+'.dbf')
  m.rtal = ALIAS()
  IF OpenFile(pmee+'\ssacts\moves', 'moves', 'shar')>0
   IF USED('moves')
    USE IN moves
   ENDIF 
  ENDIF 
  SELECT (rtal)
 ENDIF 


 ooal = ALIAS()
 SELECT recid, resume, conclusion, recommend FROM ssacts WHERE ;
  period=m.gcperiod AND mcod=m.mcod AND codexp=INT(VAL(m.TipOfExp)) AND ;
  tipacc=m.tipacc AND sn_pol=PADR(STRTRAN(m.sn_pol,' ',''),25) AND docexp=m.ddocexp AND reason=m.rreason ;
  INTO CURSOR rqwest NOCONSOLE 
 *SELECT recid, resume, conclusion, recommend FROM ssacts WHERE ;
  period=m.gcperiod AND mcod=m.mcod AND codexp=INT(VAL(m.TipOfExp)) AND ;
  sn_pol=PADR(STRTRAN(m.sn_pol,' ',''),25) AND docexp=m.ddocexp AND reason=m.rreason ;
  INTO CURSOR rqwest NOCONSOLE 
 m.nfileid = recid
 *m.resume     = resume
 *m.conclusion = conclusion
 *m.recommend  = recommend
 m.resume     = ''
 m.conclusion = ''
 m.recommend  = ''
 USE 
 SELECT (ooal)
 
 IF m.nfileid>0
  *m.DocName = pmee+'\ssacts\'+PADL(m.nfileid,6,'0')
  m.DocName = IIF(!IsUsrDir, m.pmee, m.usrdir)+'\ssacts\'+PADL(m.nfileid,6,'0')
  *m.n_akt = NumActOfExp(m.lpuid,m.tipofexp,goApp.reason,m.nfileid)
  m.n_akt = NumActOfExp(m.lpuid,m.tipofexp,m.rreason,m.nfileid)
 ELSE 
   SELECT TOP 1 resume, conclusion, recommend FROM ssacts ORDER BY recid DESC INTO CURSOR cqwert
   *m.resume     = resume
   *m.conclusion = conclusion
   *m.recommend  = recommend
   m.resume     = ''
   m.conclusion = ''
   m.recommend  = ''
   USE IN cqwert
   SELECT (ooal)

   *INSERT INTO ssacts (n_rss,period,doctyp,mcod,lpu_id,codexp,tipacc,sn_pol,fam,im,ot,docexp,doctyp,reason,smoexp,qr) ;
    VALUES ;
   (m.n_rss,m.gcperiod,'DOC',m.mcod,m.lpuid,INT(VAL(m.tipofexp)),m.tipacc,;
    PADR(STRTRAN(m.sn_pol,' ',''),25),people.fam,people.im,people.ot,m.ddocexp,'XLS',goApp.reason,goApp.smoexp,.t.)
   INSERT INTO ssacts (n_rss,period,mcod,lpu_id,codexp,tipacc,sn_pol,fam,im,ot,docexp,reason,smoexp,qr) ;
    VALUES ;
   (m.n_rss,m.gcperiod,m.mcod,m.lpuid,INT(VAL(m.tipofexp)),m.tipacc,;
    PADR(STRTRAN(m.sn_pol,' ',''),25),people.fam,people.im,people.ot,m.ddocexp,m.rreason,goApp.smoexp,.t.)

   m.nfileid = GETAUTOINCVALUE()
   *m.n_akt = NumActOfExp(m.lpuid,m.tipofexp,goApp.reason,m.nfileid)
   m.n_akt = NumActOfExp(m.lpuid,m.tipofexp,m.rreason,m.nfileid)

   *m.DocName = pmee+'\ssacts\'+PADL(m.nfileid,6,'0')
   m.DocName = IIF(!IsUsrDir, m.pmee, m.usrdir)+'\ssacts\'+PADL(m.nfileid,6,'0')
   UPDATE ssacts SET actname=PADL(m.nfileid,6,'0')+'.xls', actdate=DATETIME(), n_akt=m.n_akt, qr=.t. WHERE recid = m.nfileid

 ENDIF 
 
 m.lastid = 0
 m.lastet = ''
 m.maxdat = {}
 IF USED('moves')
  SELECT MAX(dat) as maxdat FROM moves GROUP BY actid INTO CURSOR curlst WHERE actid = m.nfileid
  m.maxdat = curlst.maxdat
  SELECT recid as lastid, et as lastet FROM moves INTO CURSOR curlst WHERE actid = m.nfileid AND dat=m.maxdat
  m.lastet = curlst.lastet
  m.lastid = curlst.lastid
  USE IN curlst
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

    IF USED('rrfile')
     USE IN rrfile
    ENDIF 
    IF USED('moves')
     USE IN moves
    ENDIF 
    SELECT (oal4)
    RETURN 
   ELSE 

    UPDATE ssacts SET actdate=DATETIME() WHERE recid = m.nfileid
    IF USED('moves')
     IF m.lastet = '1'
      UPDATE moves SET dat=DATETIME() WHERE recid = m.lastid
     ELSE 
      INSERT INTO moves (actid,et,usr,dat) VALUES (m.nfileid,'1',m.gcUser,DATETIME())
     ENDIF 
    ENDIF 
   
   ENDIF 
  ELSE 
   UPDATE ssacts SET actdate=DATETIME() WHERE recid = m.nfileid
   IF USED('moves')
    IF m.lastet = '1'
     UPDATE moves SET dat=DATETIME() WHERE recid = m.lastid
    ELSE 
     INSERT INTO moves (actid,et,usr,dat) VALUES (m.nfileid,'1',m.gcUser,DATETIME())
    ENDIF 
   ENDIF 
  ENDIF && IF m.IsVisible=.t.

  ELSE 

   UPDATE ssacts SET actdate=DATETIME() WHERE recid = m.nfileid
   IF USED('moves')
    IF m.lastet = '1'
     UPDATE moves SET dat=DATETIME() WHERE recid = m.lastid
    ELSE 
     INSERT INTO moves (actid,et,usr,dat) VALUES (m.nfileid,'1',m.gcUser,DATETIME())
    ENDIF 
   ENDIF 

  ENDIF && IF  m.IsMulti=0
 ELSE 
  IF USED('moves')
   IF m.lastet = '1'
    UPDATE moves SET dat=DATETIME() WHERE recid = m.lastid
   ELSE 
    INSERT INTO moves (actid,et,usr,dat) VALUES (m.nfileid,'1',m.gcUser,DATETIME())
   ENDIF 
  ENDIF 
 ENDIF && IF fso.FileExists(m.DocName+'.xls')

 *IF !m.IsOkAct
  DO FORM TxtForActs && Убрать!
 *ENDIF 

 UPDATE ssacts SET resume=ALLTRIM(m.resume), conclusion=ALLTRIM(m.conclusion), recommend=ALLTRIM(m.recommend) WHERE recid = m.nfileid

 CREATE CURSOR ttalon (nrec i AUTOINC, recid i, sn_pol c(25), c_i c(25), docotd c(100), ds c(6), d_in d, d_u d, pcod c(10),;
  otd c(4), cod n(6), k_u n(3), d_type c(1), s_all n(11,2), err_mee c(3), docexp c(7), ishod n(3), is_name c(60), ;
  osn230 c(5), reason c(1), s_1 n(11,2), s_2 n(11,2), ername c(250))
 INDEX on recid TAG recid
 SET ORDER TO recid 

 CREATE CURSOR curds (ds c(6), dsname c(160))
 INDEX on ds TAG ds
 SET ORDER TO ds

 CREATE CURSOR curdefs (er_c c(2), osn230 c(5), ername c(100))
 INDEX on er_c TAG er_c
 SET ORDER TO er_c

 CREATE CURSOR unDocOtd (pcod c(10), doctorname c(100))
 INDEX on pcod TAG pcod
 SET ORDER TO pcod


 m.TipOfMee = 0
 
 m.s_lech = 0
 m.dat1   = {31.12.2099}
 m.dat2   = {01.01.2000}
 m.dslast = 0
 m.totdefs = 0
 m.nekoplate = 0
 m.tot_straf = 0
 m.IsPlkStraf = .f.
 m.IsStacStraf = .f.
 m.karta = 'qwert'

 m.actbody = ''
 m.vidp    = '00'
 SELECT merror
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
  
*  REPLACE n_akt WITH m.nfileid, t_akt WITH 'SS'
  REPLACE n_akt WITH m.n_akt, d_akt WITH goApp.d_exp, merror.t_akt WITH 'SS' IN merror 

  IF USED('rrfile')
   IF SEEK(m.sn_pol, 'rrfile')
    REPLACE n_akt WITH m.n_akt, d_akt WITH goApp.d_exp, t_akt WITH 'SS' IN rrfile
   ENDIF 
  ENDIF 

*  m.povod = IIF(SEEK(m.rreason, 'reasons'), reasons.name, '')
  m.povod = IIF(m.TipOfExp != '6', IIF(SEEK(m.rreason, 'reasons'), reasons.name, ''), IIF(SEEK(m.gcUser+m.rreason, 'themes'), themes.name, ''))
 
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
  m.ername   = IIF(SEEK(m.er_c, 'sookod'), ALLTRIM(sookod.f_naim), '')
  m.osn230   = IIF(m.er_c!='W0', osn230, '0.0.0')
  m.ername   = IIF(SEEK(m.err_mee,'errsmee'), ALLTRIM(errsmee.f_komment), '')
  m.tip      = talon.tip
  m.s_1      = s_1
  m.straf    = straf
*  m.s_2      = s_2
  IF !EMPTY(m.tip)
   IF m.straf > 0 AND IIF(!EMPTY(m.tip) AND m.c_i!=m.karta, .t., .f.)
    m.IsStacStraf = .t.
   ELSE 
    m.IsStacStraf = .f.
   ENDIF 
   m.s_2      = IIF(m.IsStacStraf=.t., merror.s_2, 0)
  ELSE 
   m.s_2      = IIF(m.IsPlkStraf=.f., merror.s_2, 0)
   IF m.straf > 0 AND m.IsPlkStraf = .f.
    m.IsPlkStraf = .t.
   ENDIF
  ENDIF  
  m.karta = IIF(!EMPTY(m.tip), m.c_i, m.karta)
  
  m.totdefs   = m.totdefs + IIF(m.er_c!='W0', 1, 0)
  m.nekoplate = m.nekoplate + m.s_1
  m.tot_straf = m.tot_straf + m.s_2

  m.cod        = cod 
  m.otd        = talon.otd
  m.otdname    = IIF(USED('otdel'), IIF(SEEK(m.otd,'otdel'), ALLTRIM(otdel.name), ''), '')
  m.pcod       = talon.pcod 
  m.doctorname = IIF(USED('doctor'), IIF(SEEK(m.pcod, 'doctor'), PROPER(ALLTRIM(doctor.fam))+' '+;
   PROPER(ALLTRIM(doctor.im))+' '+PROPER(ALLTRIM(doctor.ot))+', '+pcod, ''), '')

*  m.actbody = m.actbody + PADR(PADL(m.cod,6,'0'),8)
  m.actbody = m.actbody + PADR(m.cod,8)
  m.actbody = m.actbody + STRTRAN(DTOC(m.d_u),'.','')
  m.actbody = m.actbody + PADR(m.osn230,6)
  m.actbody = m.actbody + PADR(ALLTRIM(TRANSFORM(m.s_1,'9999999.99')),10)
  
  IF m.vidp = '00'
   DO CASE 
    CASE IsGsp(m.cod)
     m.vidp = '03'
    CASE IsDst(m.cod)
     m.vidp = '04'
    CASE IsPlk(m.cod)
     m.vidp = '02'
    CASE Is02(m.cod)
     m.vidp = '05'
   ENDCASE 
  ENDIF 

  IF IsUsl(m.cod)
   m.docotd = m.doctorname
  ELSE 
   m.docotd = m.otdname
  ENDIF 

  IF !EMPTY(m.pcod) AND !SEEK(m.pcod, 'unDocOtd')
   INSERT INTO unDocOtd FROM MEMVAR 
  ENDIF 

  m.d_in = m.d_u
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
  
*  m.d_beg = IIF(IsMes(m.cod) OR IsVMP(m.cod), IIF(m.k_u>1, m.d_u-m.k_u-1, m.d_u-1), m.d_u)

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

 m.expname = 'Акт '
 DO CASE 
  CASE m.tipofexp = '4'
   m.podvid='0'
   m.expname = m.expname + 'плановой'
  CASE m.tipofexp = '5'
   m.podvid='1'
   m.expname = m.expname + 'целевой'
  CASE m.tipofexp = '6'
   m.podvid='Т'
   m.expname = m.expname + 'тематической'
  CASE m.tipofexp = '9'
   m.podvid='1'
   m.expname = m.expname + 'целевой (по жалобе)'
  OTHERWISE 
   m.podvid='0'
 ENDCASE 
* m.n_akt = mcod + m.qcod + PADL(m.tMonth,2,'0') + RIGHT(STR(tYear,4),1)+'/'+ALLTRIM(STR(m.nfileid))
* m.n_akt = m.qcod+STR(m.lpuid,4)+IIF(INLIST(m.TipOfExp,'2','3','7'),'1','2')+;
   IIF(INLIST(m.TipOfExp,'2','4','6','7','8'),'1','2')+m.podvid+ALLTRIM(STR(m.nfileid))
 m.d_akt =  IIF(m.qcod!='I3', DTOC(DATE()), '')
 m.d_exp = {}
 m.zakl_name01 = 'Экспертное заключение №'+m.n_akt+IIF(m.qcod!='P2',' от '+m.d_akt,'')
 m.zakl_name02 = '(протокол оценки качества медицинской помощи, '+m.ctipofexp+')'
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

 *UPDATE ssacts SET s_all=m.s_lech, s_def=m.nekoplate, s_fee=m.tot_straf WHERE recid = m.nfileid
 UPDATE ssacts SET s_all=m.s_lech, s_def=m.nekoplate, s_fee=m.tot_straf,; 
	n_st=IIF(m.vidp='03',1,0), n_dst=IIF(m.vidp='04',1,0), n_plk=IIF(m.vidp='02',1,0), n_02=IIF(m.vidp='05',1,0) ;
	WHERE recid = m.nfileid
 
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
* m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)
 SELECT ttalon
 SUM s_1 TO m.nekoplate
 IF m.qcod!='I3'
  m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)
 ELSE 
  m.llResult = X_Report(m.lcTmpName, m.lcRepName, .F.)
  PUBLIC oExcel AS Excel.Application
  WAIT "Запуск MS Excel..." WINDOW NOWAIT 
  TRY 
   oExcel=GETOBJECT(,"Excel.Application")
  CATCH 
   oExcel=CREATEOBJECT("Excel.Application")
  ENDTRY 
  WAIT CLEAR 

  m.acthead = '2'+PADR(STRTRAN(m.n_akt,'Т','T'),16)+PADR(goApp.smoexp,7)+PADR(m.ddocexp,7)+PADR(m.sn_pol,20)+;
  	IIF(m.tipofexp='5', m.rreason, ' ')+m.vidp+PADR(ALLTRIM(TRANSFORM(m.tot_straf,'9999999.99')),10)+;
  	PADR(ALLTRIM(TRANSFORM(m.nekoplate,'9999999.99')),10)
  	

  m.barcode = m.acthead+m.actbody

  lcQRImage = loFbc.QRBarcodeImage(m.barcode,pMee+'\ssacts\'+PADL(m.nfileid,6,'0')+'.png',6,2)

  OneBook  = oExcel.Workbooks.Add(m.docname)
  oSheet=OneBook.ActiveSheet
  IF m.dotname <> m.dotnamew0 && !m.IsOkAct
   oPic = oSheet.Range("A1").Parent.Pictures.Insert(pMee+'\ssacts\'+PADL(m.nfileid,6,'0')+'.png')
  ENDIF 
  IF fso.FileExists(pMee+'\ssacts\'+PADL(m.nfileid,6,'0')+'.png')
   fso.DeleteFile(pMee+'\ssacts\'+PADL(m.nfileid,6,'0')+'.png')
  ENDIF 
  
  IF m.dotname <> m.dotnamew0 && !m.IsOkAct
  WITH oSheet.Range("AA1:AA5")
   oPic.Top = .Top
   oPic.Left = .Left
   oPic.Height = .Height
  ENDWITH 
  ENDIF 

  osheet.Protect('qwerty',,.t.)
 
  IF fso.FileExists(m.docname+'.xls')
   TRY 
    fso.DeleteFile(m.docname+'.xls')
    OneBook.SaveAs(m.docname,18)
    *oBook.Close && Убрать!
   CATCH  
    MESSAGEBOX('ФАЙЛ '+m.docname+'.XLS ОКТРЫТ!',0+64,'')
   ENDTRY 
  ELSE 
   oBook.SaveAs(m.docname,18)
   oBook.Close && Убрать!
  ENDIF 
  oExcel.Visible = m.IsVisible
 ENDIF 

 IF USED('rrfile')
  USE IN rrfile
 ENDIF 
 IF USED('moves')
  USE IN moves
 ENDIF 
 USE IN ttalon
 USE IN curds
 USE IN curdefs
 USE IN unDocOtd
 SELECT (oal4)

 WAIT CLEAR 

RETURN 
