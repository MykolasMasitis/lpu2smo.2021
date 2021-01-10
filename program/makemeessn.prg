FUNCTION MakeMEESSn(para1, IsVisible, IsQuit, TipAcc, para5) && sn_pol, .t., .t., goApp.TipAcc, '2'

 DotName = 'MEESS_N.xls'
 IF !fso.FileExists(pTempl+'\'+DotName)
  MESSAGEBOX('ОТСУТСТВУЕТ ФАЙЛ ШАБЛОН ОТЧЕТА'+CHR(13)+CHR(10)+;
   'ActMEEssI3.xls',0+32,'')
  RETURN 
 ENDIF 

 m.usrdir = fso.GetParentFolderName(pbin) + '\'+UPPER(m.gcuser)
 
 m.sn_pol   = para1
 m.TipOfExp = para5

 m.d_beg    = IIF(SEEK(m.sn_pol, 'people', 'sn_pol'), people.d_beg, {})
 m.d_end    = IIF(SEEK(m.sn_pol, 'people', 'sn_pol'), people.d_end, {})
 
 m.dr = DTOC(IIF(SEEK(m.sn_pol, 'people', 'sn_pol'), people.dr, {}) )
 m.w  = IIF(SEEK(m.sn_pol, 'people', 'sn_pol'), people.w, 0)
 m.sex = IIF(m.w=1, 'муж', IIF(m.w=2, 'жен', ''))

 DO CASE 
  CASE m.TipOfExp = '2'
   m.ctipofexp = 'плановой'
  CASE m.TipOfExp = '3'
   m.ctipofexp = 'целевой'
  CASE m.TipOfExp = '7'
   m.ctipofexp = 'тематической'
  CASE m.TipOfExp = '8'
   m.ctipofexp = 'по жалобе'
  OTHERWISE 
   m.ctipofexp = ''
 ENDCASE 

 IF !EMPTY(goApp.d_exp)
  m.d_exp = DTOC(goApp.d_exp)
 ELSE 
  m.d_exp = ''
 ENDIF 
 m.d_exp = ''

 m.flcod       = goApp.flcod
 m.mcod        = goApp.mcod 
 m.lpuid       = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)

 m.lpuname     = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.fullname)+', '+m.mcod, '')
 m.lpuaddress  = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.adres), '')
 m.lpuname     = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.fullname)+', '+m.mcod+', '+m.lpuaddress, '')
 m.lpudog      = IIF(SEEK(m.lpuid, 'lpudogs'), ALLTRIM(lpudogs.dogs), '')
 m.lpuddog     = IIF(SEEK(m.lpuid, 'lpudogs'), lpudogs.ddogs, {})
 m.lpudog      = 'в соответствии с Договором '+m.lpudog+' от '+DTOC(m.lpuddog)
 m.TipOfPeriod = IIF(EMPTY(m.flcod),0,1) && 0-локальный период, 1 - сводный!
 m.lpuboss     = IIF(SEEK(m.lpuid, 'lpudogs'), ALLTRIM(lpudogs.boss), '')

 m.exp_dat1 = '01.'+PADL(tMonth,2,'0')+'.'+STR(tYear,4)
 m.exp_dat2 = DTOC(GOMONTH(CTOD(m.exp_dat1),1)-1)

 m.ppolis = STRTRAN(ALLTRIM(m.sn_pol),' ','') && Для названия Акта
 
 old_al = ALIAS() 
 *SELECT COUNT(*) AS cnt FROM talon a, merror b WHERE a.recid=b.recid AND b.et=m.TipOfExp AND a.sn_pol=m.sn_pol;
 	INTO CURSOR curcur
 SELECT b.reason as reason, b.err_mee as err_mee, count(*) as cnt FROM talon a, merror b WHERE a.recid=b.recid AND b.et=m.TipOfExp AND a.sn_pol=m.sn_pol;
 	GROUP BY reason, err_mee INTO CURSOR curcur
 m.nIsExps = _tally

 IF m.nIsExps == 0
  USE 
  SELECT (old_al)
  MESSAGEBOX(CHR(13)+CHR(10)+'ПО ВЫБРАННОМУ СЧЕТУ МЭЭ'+CHR(13)+CHR(10)+;
  	'ДАННОГО ТИПА НЕ ПРОВОДИЛОСЬ!'+CHR(13)+CHR(10), 0+64, '')
  RETURN
 ENDIF 
 
 IF m.TipOfExp != '3' && МЭЭ Целевая
  CALCULATE SUM(cnt) FOR err_mee!='W0' TO m.nIsExpsS
 * SELECT merror.reason as reason, count(*) as cnt FROM talon a, merror b WHERE a.recid=b.recid AND b.et=m.TipOfExp AND a.sn_pol=m.sn_pol;
 * 	AND b.err_mee!='W0' INTO CURSOR curcur
 ELSE
  CALCULATE SUM(cnt) TO m.nIsExpsS
 * SELECT merror.reason as reason, count(*) as cnt FROM talon a, merror b WHERE a.recid=b.recid AND b.et=m.TipOfExp AND a.sn_pol=m.sn_pol ;
 * 	INTO CURSOR curcur
 ENDIF 

 *m.nIsExpsS = curcur.cnt
 IF INLIST(m.TipOfExp,'6','7')
  m.theme  = curcur.reason
 ELSE 
  m.reason = curcur.reason
 ENDIF 

 *USE 
 *SELECT talon 
 SELECT (old_al)

 IF m.nIsExpsS == 0
  MESSAGEBOX(CHR(13)+CHR(10)+'ПО ВЫБРАННОМУ СЧЕТУ ОШИБОК НЕ ВЫЯВЛЕНО!'+CHR(13)+CHR(10)+;
  	'РЕЗУЛЬТАТЫ (W0) ПОПАДУТ ТОЛЬКО В СВОДНЫЙ АКТ!',0+64,'')
  RETURN
 ENDIF 
 
 SELECT curcur
 SCAN 
  m.reason = reason
  m.cnt    = cnt
  
  IF m.cnt<=0
   LOOP 
  ENDIF 
  
  IF m.TipOfExp != '3' && МЭЭ Целевая
   IF err_mee = 'W0'
    LOOP 
   ENDIF 
  ENDIF 

  =MakeMeeSS1(m.sn_pol, m.IsVisible, m.IsQuit, m.TipAcc, m.TipOfExp, IIF(m.TipOfExp='3', m.reason, ''))
  
 ENDSCAN 
 USE IN curcur 
 SELECT (old_al)
 
RETURN 

FUNCTION MakeMeeSS1(para1, para2, para3, para4, para5, para6)
 m.sn_pol    = para1
 m.IsVisible = para2
 m.IsQuit    = para3
 m.TipAcc    = para4
 m.TipOfExp  = para5
 m.reason    = para6

 * Добавлено 15.12.2018 
 * Не все акты попадают в реестр, искажая результаты экспертизы
 IF EMPTY(goApp.d_exp)
  MESSAGEBOX('НЕ ЗАДАНА ДАТА ЭКСПЕРТИЗЫ!'+CHR(13)+CHR(10)+'ФОРМИРОВАНИЕ АКТА НЕВОЗМОЖНО!',0+64,'')
  RETURN 
 ENDIF 

 m.n_rss = 0
 m.vvir = m.mcod + DTOS(goApp.d_exp) && Это вопрос!!!
 IF !SEEK(m.vvir, 'rss')
  m.e_period = STR(YEAR(goApp.d_exp),4)+PADL(MONTH(goApp.d_exp),2,'0')
  INSERT INTO rss (lpu_id, mcod, d_u, e_period, smoexp, k_acts) VALUES (m.lpuid, m.mcod, goApp.d_exp, m.e_period, goApp.smoexp, 1)
  m.n_rss = GETAUTOINCVALUE()
 ELSE 
  m.n_rss    = rss.recid
  m.e_period = rss.e_period
  UPDATE rss SET k_acts = k_acts+1 WHERE recid=m.n_rss
 ENDIF 

* m.vvir   = m.mcod + m.gcperiod + m.TipOfExp + IIF(INLIST(m.TipOfExp,'6','7'), m.theme, m.reason)
 *m.vvir   = m.mcod + m.gcperiod
 m.vvir   = goApp.smoexp + m.mcod + m.gcperiod + m.TipOfExp
 m.n_rqst = IIF(SEEK(m.vvir, 'rqst'), rqst.recid, 0)
 IF m.n_rqst>0
  m.rqstfile = PADL(m.n_rqst,6,'0')
  IF fso.FileExists(pmee+'\requests\'+m.rqstfile+'.dbf')
   m.rtal = ALIAS()
   IF OpenFile(pmee+'\requests\'+m.rqstfile, 'rrfile', 'shar')>0
    IF USED('rrfile')
     USE IN rrfile
    ENDIF 
   ENDIF 
   SELECT rrfile
   IF ATAGINFO(qqww,'')<1 OR (ATAGINFO(qqww,'')=1 AND qqww(1,1)!='SN_POL')
    USE IN rrfile
    IF OpenFile(pmee+'\requests\'+m.rqstfile, 'rrfile', 'excl')>0
     IF USED('rrfile')
      USE IN rrfile
     ENDIF 
    ELSE 
     SELECT rrfile
     INDEX on sn_pol TAG sn_pol
     USE 
     IF OpenFile(pmee+'\requests\'+m.rqstfile, 'rrfile', 'shar', 'sn_pol')>0
      IF USED('rrfile')
       USE IN rrfile
      ENDIF 
     ENDIF 
    ENDIF 
   ELSE 
    SET ORDER TO SN_POL
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

 SELECT recid, resume, conclusion, recommend FROM ssacts WHERE period=m.gcperiod AND mcod=m.mcod AND flcod=m.flcod AND ;
  codexp=INT(VAL(m.TipOfExp)) AND tipacc=m.tipacc AND sn_pol=PADR(STRTRAN(m.sn_pol,' ',''),25) ;
  AND reason=m.reason INTO CURSOR rqwest NOCONSOLE 
 m.nfileid    = recid
 m.resume     = resume
 m.conclusion = conclusion
 m.recommend  = recommend
 USE 
 SELECT (old_al)
 
 IF m.nfileid>0
  DocName = IIF(!IsUsrDir, m.pmee, m.usrdir)+'\ssacts\'+PADL(m.nfileid,6,'0')
  m.n_akt = NumActOfExp(m.lpuid, m.tipofexp, m.reason, m.nfileid)
  
 ELSE 


  SELECT TOP 1 resume, conclusion, recommend FROM ssacts ORDER BY recid DESC INTO CURSOR cqwert
  m.resume     = resume
  m.conclusion = conclusion
  m.recommend  = recommend
  USE IN cqwert

  *INSERT INTO ssacts (n_rss,period,mcod,lpu_id,codexp,flcod,e_period,smoexp,tipacc,sn_pol,fam,im,ot,reason,qr) ;
    VALUES ;
   (m.n_rss,m.gcperiod,m.mcod,m.lpuid,INT(VAL(m.tipofexp)),m.flcod,m.e_period,goApp.smoexp,;
  	m.tipacc, PADR(STRTRAN(m.sn_pol,' ',''),25),people.fam,people.im,people.ot,;
  IIF(!INLIST(goApp.etap,'6','7'), m.reason, m.theme),.t.)
  INSERT INTO ssacts (n_rss,period,mcod,lpu_id,codexp,flcod,smoexp,tipacc,sn_pol,fam,im,ot,reason,qr) ;
    VALUES ;
   (m.n_rss,m.gcperiod,m.mcod,m.lpuid,INT(VAL(m.tipofexp)),m.flcod,goApp.smoexp,;
  	m.tipacc, PADR(STRTRAN(m.sn_pol,' ',''),25),people.fam,people.im,people.ot,;
  IIF(!INLIST(goApp.etap,'6','7'), m.reason, m.theme),.t.)

  m.nfileid = GETAUTOINCVALUE()
  m.n_akt = NumActOfExp(m.lpuid,m.tipofexp,m.reason,m.nfileid)

  DocName = IIF(!IsUsrDir, m.pmee, m.usrdir)+'\ssacts\'+PADL(m.nfileid,6,'0')
  UPDATE ssacts SET actname=PADL(m.nfileid,6,'0')+'.xls', actdate=DATETIME(), n_akt=m.n_akt WHERE recid = m.nfileid
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
 
 IF fso.FileExists(DocName+'.xls')
  oFile = fso.GetFile(DocName+'.xls')
  DateCreated      = TTOC(oFile.DateCreated)
  DateLastAccessed = TTOC(oFile.DateLastAccessed)
  DateLastModified = TTOC(oFile.DateLastModified)
  RELEASE oFile
  
  IF m.IsVisible
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
    USE IN ttalon
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
  ENDIF && IF m.IsVisible
 ELSE 
  IF USED('moves')
   INSERT INTO moves (actid,et,usr,dat) VALUES (m.nfileid,'1',m.gcUser,DATETIME())
  ENDIF 
 ENDIF && IF fso.FileExists(DocName+'.xls')

 DO FORM TxtForActs
 UPDATE ssacts SET resume=ALLTRIM(m.resume), conclusion=ALLTRIM(m.conclusion), recommend=ALLTRIM(m.recommend) WHERE recid = m.nfileid

 CREATE CURSOR concls (nrec n(5), concl c(250))
 
 CREATE CURSOR ttalon (nrec i AUTOINC, sn_pol c(25), c_i c(30), ds c(6), tip c(1), d_u d, pcod c(10), otd c(4), cod n(6), k_u n(3),;
  d_type c(1), s_all n(11,2), e_cod n(6), e_ku n(3), e_tip c(1), err_mee c(3), e_period c(6), et c(1), s_1 n(11,2), s_2 n(11,2),;
  koeff n(4,2), ee c(1), straf n(11,2),osn230 c(5),docotd c(100), reason c(1))

 m.TipOfMee = 0
 
 m.s_lech = 0
 m.dat1 = {31.12.2099}
 m.dat2 = {01.01.2000}
 m.dslast = 0
 m.IsPlkStraf = .f.
 m.IsStacStraf = .f.
 m.karta = 'qwert'
* m.karta = c_i
 
 * Сделано закомментированного ниже модуля
 SELECT merror
 SCAN
  IF et!=m.TipOfExp
   LOOP 
  ENDIF 
  m.recid = recid
  IF USED('serror')
   IF SEEK(m.recid, 'serror')
    LOOP 
   ENDIF 
  ENDIF 
  IF !SEEK(m.recid, 'talon', 'recid')
   LOOP 
  ENDIF 
  m.polis = talon.sn_pol
  IF m.polis != m.sn_pol
   LOOP 
  ENDIF 
  IF reason!=IIF(m.TipOfExp='3', m.reason, '')
   LOOP 
  ENDIF 
  
  m.c_i    = talon.c_i
  m.ds     = talon.ds
  m.pcod   = talon.pcod
  m.otd    = talon.otd
  m.cod    = talon.cod
  m.tip    = talon.tip
  m.d_u    = talon.d_u
  m.k_u    = talon.k_u
  m.d_type = talon.d_type
  m.s_all  = talon.s_all
  
  m.et       = et
  m.ee       = ee
  m.e_cod    = e_cod
  m.e_ku     = e_ku
  m.e_tip    = e_tip
  m.err_mee  = err_mee 
  m.osn230   = IIF(err_mee ='W0', '0.0.0', osn230)
  m.koeff    = koeff
  m.e_period = e_period
  m.straf    = straf
  m.s_1      = s_1

  m.otdname    = IIF(USED('otdel'), IIF(SEEK(m.otd,'otdel'), ALLTRIM(otdel.name), ''), '')
  m.doctorname = IIF(USED('doctor'), IIF(SEEK(m.pcod, 'doctor'), PROPER(ALLTRIM(doctor.fam))+' '+;
   PROPER(ALLTRIM(doctor.im))+' '+PROPER(ALLTRIM(doctor.ot))+', '+pcod, ''), '')

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
  m.reason   = merror.reason
  
  DO CASE 
   CASE IsPlk(m.cod)
    m.docotd = m.doctorname
   CASE IsGsp(m.cod)
    m.docotd = m.otdname
   CASE Is02(m.cod)
    m.docotd = ALLTRIM(m.otd)+'/'+ALLTRIM(m.pcod)
  ENDCASE 
   
  m.s_lech = m.s_lech + m.s_all
  DO CASE 
   CASE m.TipAcc == 0 && Сводный счет
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
    m.dslast = m.dslast + k_u
  ENDCASE 
  IF m.et==m.TipOfExp
   DO CASE 
    CASE EMPTY(merror.err_mee)
    CASE LEFT(merror.err_mee,2)=='W0'
     m.TipOfMee = IIF(m.TipOfMee<=1, 1, m.TipOfMee)
    OTHERWISE 
     m.TipOfMee = 2 && Есть ошибки
   ENDCASE 
  ENDIF 

  REPLACE n_akt WITH m.n_akt, d_akt WITH goApp.d_exp, t_akt WITH 'SS'

  IF USED('rrfile')
   IF SEEK(m.sn_pol, 'rrfile')
    REPLACE n_akt WITH m.n_akt, d_akt WITH goApp.d_exp, t_akt WITH 'SS' IN rrfile
   ENDIF 
  ENDIF 

  INSERT INTO ttalon FROM MEMVAR 
 ENDSCAN 
 SELECT (old_al)
 * Сделано закомментированного ниже модуля
 
 m.dat1 = IIF(m.TipAcc==3, m.dat1, m.dat1)
 
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
 
* Алгоритм нумерации актов изменен по просьбе Ингосстрах-М с целью унификации внутренней нумерации и нумерации актов, 
* подаваемой в МГФОМС. 
* Необходимо изменить связку тема тематической экспертизы и повода остальных видов экспертизы.

* m.d_akt = IIF(m.qcod!='I3', DTOC(DATE()), '')
 m.d_akt = ''
* m.n_akt = mcod + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)+'/'+ALLTRIM(STR(m.nfileid))
 m.cpredps = 'Предписание № '+m.n_akt+' от '+m.d_akt
 m.vidofexp = ''
 DO CASE 
  CASE m.TipOfExp = '2'
   m.vidofexp='Плановая МЭЭ'
  CASE m.TipOfExp = '3'
   m.vidofexp='Целевая МЭЭ'
  CASE m.TipOfExp = '7'
   m.vidofexp='Тематическая МЭЭ'
  CASE m.TipOfExp = '8'
   m.vidofexp='МЭЭ по жалобе'
  OTHERWISE 
 ENDCASE 

 m.povod = IIF(m.TipOfExp != '7', IIF(SEEK(m.reason, 'reasons'), reasons.name, ''), IIF(SEEK(m.gcUser+m.reason, 'themes'), themes.name, ''))

 IF m.TipOfPeriod=0
*  m.n_akt = m.qcod+STR(m.lpuid,4)+IIF(INLIST(m.TipOfExp,'2','3','7'),'1','2')+;
   IIF(INLIST(m.TipOfExp,'2','4','6','7','8'),'1','2')+m.podvid+ALLTRIM(STR(m.nfileid))
 ELSE 
*  m.n_akt = m.qcod+STR(m.lpuid,4)+IIF(INLIST(m.TipOfExp,'2','3','7'),'1','2')+;
   IIF(INLIST(m.TipOfExp,'2','4','6','7','9'),'1','2')+m.podvid+ALLTRIM(STR(m.nfileid))
 ENDIF 
 
* m.d_akt = IIF(m.qcod!='I3', DTOC(DATE()), '')
 m.d_akt = ''
 m.d_akt2 = m.d_akt
 m.n_schet = STR(tYear,4)+PADL(tMonth,2,'0')
 IF m.TipAcc == 0
  m.dat1 = IIF(SEEK(m.sn_pol, 'people'), people.d_beg, {})
  m.dat2 = IIF(SEEK(m.sn_pol, 'people'), people.d_end, {})
 ENDIF 
 m.dslast = IIF(!INLIST(m.TipAcc,2,3), IIF(m.dat2-m.dat1>0, m.dat2-m.dat1, 1), m.dslast)
 m.ds = ds
 m.dsnam = IIF(SEEK(m.ds, 'mkb10'), ALLTRIM(mkb10.name_ds), '')
 m.pcod = pcod
 m.docfam = IIF(USED('doctor'), IIF(SEEK(m.pcod, 'doctor'), ALLTRIM(doctor.Fam)+' '+ALLTRIM(doctor.Im)+' '+ALLTRIM(doctor.Ot), ''), '')
 m.fioexp  = ''
 IF !EMPTY(goApp.smoexp)
  IF USED('users')
   IF SEEK(ALLTRIM(goApp.smoexp), 'users', 'name')
     m.fioexp  = ALLTRIM(users.fam)+' '+ALLTRIM(users.im)+' '+ALLTRIM(users.ot)
   ENDIF 
  ENDIF 
 ELSE 
 m.fioexp  = m.usrfam+' '+m.usrim+' '+m.usrot
 ENDIF 
 m.fioexp2 = LEFT(m.usrim,1)+'.'+LEFT(m.usrot,1)+'.'+m.usrfam

 SELECT ttalon

 nRow = 2
 m.tot_badsum = 0
 m.tot_straf = 0
 m.tot_goodsum = 0
 m.actbody = ''
 m.vidp = '00'
 
 SCAN 
*  m.actbody = m.actbody + PADR(PADL(cod,6,'0'),8)
  m.actbody = m.actbody + PADR(cod,8)
  m.actbody = m.actbody + STRTRAN(DTOC(d_u),'.','')
  
  m.er_c = err_mee
  m.osn230 = IIF(SEEK(LEFT(UPPER(m.er_c),2), 'sookod'), sookod.osn230, '')	
  m.actbody = m.actbody + PADR(m.osn230,6)

  m.actbody = m.actbody + PADR(ALLTRIM(TRANSFORM(s_1,'9999999.99')),10)

  m.tot_badsum = m.tot_badsum + s_1
  m.tot_goodsum = m.tot_goodsum + (s_all-s_1)

  nRow = nRow + 1
    
*  m.tot_straf = m.tot_straf + IIF( m.tot_straf<=0, straf*m.ynorm, 0)
  m.tot_straf = m.tot_straf + s_2

  IF m.vidp = '00'
   DO CASE 
    CASE IsGsp(cod)
     m.vidp = '03'
    CASE IsDst(cod)
     m.vidp = '04'
    CASE IsPlk(cod)
     m.vidp = '02'
    CASE Is02(cod)
     m.vidp = '05'
   ENDCASE 
  ENDIF 

 ENDSCAN 
* USE 


 m.koplate   = m.s_lech - m.tot_badsum
 m.nekoplate = m.tot_badsum
 m.saystraf = cpr(INT(m.tot_straf))+' '+PADL(INT((m.tot_straf-INT(m.tot_straf))*100),2,'0')+' КОП.'

 UPDATE ssacts SET s_all=m.s_lech, s_def=m.nekoplate, s_fee=m.tot_straf, ;
	n_st=IIF(m.vidp='03',1,0), n_dst=IIF(m.vidp='04',1,0), n_plk=IIF(m.vidp='02',1,0), n_02=IIF(m.vidp='05',1,0) ;
 	WHERE recid = m.nfileid

 IF USED('rrfile')
  IF SEEK(m.sn_pol, 'rrfile')
   REPLACE n_akt WITH m.n_akt, d_akt WITH goApp.d_exp, t_akt WITH 'SS' IN rrfile
  ENDIF 
 ENDIF 
 IF USED('rrfile')
  USE IN rrfile
 ENDIF 
 IF USED('moves')
  USE IN moves
 ENDIF 

 m.resume = ''
 IF m.TipOfMee=1
  m.resume = 'По представленному счету замечаний нет.'
  m.vivod  = 'Счет подлежит оплате в полном объеме из средств ОМС.'
 ENDIF 

 LOCAL m.lcTmpName, m.lcRepName, m.lcDbfName, m.llResult
 m.lcTmpName = pTempl+'\'+m.dotname
 m.lcRepName = m.docname+'.xls'

 m.conclusion   = ALLTRIM(m.conclusion)
 m.str_len = 90
 IF LEN(m.conclusion)>0 && если вообще что-то есть
  IF LEN(m.conclusion)>m.str_len
   m.s_str = ''    
   FOR m.n_ch=1 TO LEN(m.conclusion)
    m.s_str = m.s_str+ SUBSTR(m.conclusion,m.n_ch,1)
    DO CASE 
     CASE LEN(m.s_str)>=m.str_len
      IF RIGHT(m.s_str,1)=SPACE(1)
       INSERT INTO concls (concl) VALUES (m.s_str)
       m.s_str = ''    
      ELSE 
       IF RAT(SPACE(1),m.s_str)>0
        INSERT INTO concls (concl) VALUES (LEFT(m.s_str,RAT(SPACE(1),m.s_str)-1))
        m.s_str = SUBSTR(m.s_str,RAT(SPACE(1),m.s_str)+1)
       ELSE 
        INSERT INTO concls (concl) VALUES (m.s_str)
        m.s_str = ''    
       ENDIF 
      ENDIF 
     CASE SUBSTR(m.conclusion,m.n_ch,2) = CHR(13)+CHR(10)
      m.n_ch = m.n_ch + 1
      INSERT INTO concls (concl) VALUES (LEFT(m.s_str,LEN(m.s_str)-1))
      m.s_str = ''    
     OTHERWISE 
    ENDCASE 
   ENDFOR 
   INSERT INTO concls (concl) VALUES (m.s_str)
  ELSE && если строка <90 символов
   *INSERT INTO concls (concl) VALUES (m.conclusion)
   m.s_str = ''    
   FOR m.n_ch=1 TO LEN(m.conclusion)
    m.s_str = m.s_str+ SUBSTR(m.conclusion,m.n_ch,1)
    DO CASE 
     CASE LEN(m.s_str)>=m.str_len
      IF RIGHT(m.s_str,1)=SPACE(1)
       INSERT INTO concls (concl) VALUES (m.s_str)
       m.s_str = ''    
      ELSE 
       IF RAT(SPACE(1),m.s_str)>0
        INSERT INTO concls (concl) VALUES (LEFT(m.s_str,RAT(SPACE(1),m.s_str)-1))
        m.s_str = SUBSTR(m.s_str,RAT(SPACE(1),m.s_str)+1)
       ELSE 
        INSERT INTO concls (concl) VALUES (m.s_str)
        m.s_str = ''    
       ENDIF 
      ENDIF 
     CASE SUBSTR(m.conclusion,m.n_ch,2) = CHR(13)+CHR(10)
      m.n_ch = m.n_ch + 1
      INSERT INTO concls (concl) VALUES (LEFT(m.s_str,LEN(m.s_str)-1))
      m.s_str = ''    
     OTHERWISE 
    ENDCASE 
   ENDFOR 
   INSERT INTO concls (concl) VALUES (m.s_str)
  ENDIF 
 ELSE 
  INSERT INTO concls (concl) VALUES (SPACE(10))
 ENDIF 

  m.llResult = X_Report(m.lcTmpName, m.lcRepName, .F.)
  PUBLIC oExcel AS Excel.Application
  WAIT "Запуск MS Excel..." WINDOW NOWAIT 
  TRY 
   oExcel=GETOBJECT(,"Excel.Application")
  CATCH 
   oExcel=CREATEOBJECT("Excel.Application")
  ENDTRY 
  WAIT CLEAR 
 
  m.acthead = '2'+PADR(STRTRAN(m.n_akt,'Т','T'),16)+PADR(goApp.smoexp,7)+PADR(' ',7)+PADR(m.sn_pol,20)+;
  	IIF(m.tipofexp='3', m.reason, ' ')+m.vidp+PADR(ALLTRIM(TRANSFORM(m.tot_straf,'9999999.99')),10)+;
  	PADR(ALLTRIM(TRANSFORM(m.tot_badsum,'9999999.99')),10) 
  	

  m.barcode = m.acthead+m.actbody

  lcQRImage = loFbc.QRBarcodeImage(m.barcode,pMee+'\ssacts\'+PADL(m.nfileid,6,'0')+'.png',6,2)
  OneBook  = oExcel.Workbooks.Add(m.docname)
  oSheet=OneBook.ActiveSheet

  oPic = oSheet.Range("A1").Parent.Pictures.Insert(pMee+'\ssacts\'+PADL(m.nfileid,6,'0')+'.png')
  IF fso.FileExists(pMee+'\ssacts\'+PADL(m.nfileid,6,'0')+'.png')
   fso.DeleteFile(pMee+'\ssacts\'+PADL(m.nfileid,6,'0')+'.png')
  ENDIF 
  
  * WITH oSheet.Range("O2:O6")
  WITH oSheet.Range("AK1:AK5")
   oPic.Top    = .Top
   oPic.Left   = .Left
   oPic.Height = .Height
  ENDWITH 

  osheet.Protect('qwerty',,.t.)
 
  IF fso.FileExists(m.docname+'.xls')
   TRY 
    fso.DeleteFile(m.docname+'.xls')
    OneBook.SaveAs(m.docname,18)
   CATCH  
    MESSAGEBOX('ФАЙЛ '+m.docname+'.XLS ОКТРЫТ!',0+64,'')
   ENDTRY 
  ELSE 
   oBook.SaveAs(m.docname,18)
  ENDIF 
  oExcel.Visible = m.IsVisible

 WAIT CLEAR 

RETURN 