FUNCTION MakeMEESSXls(para1, IsVisible, IsQuit, TipAcc, para5)
 DotName = 'ActMEEss.xls'
 IF !fso.FileExists(pTempl+'\'+DotName)
  MESSAGEBOX('ОТСУТСТВУЕТ ФАЙЛ ШАБЛОН ОТЧЕТА'+CHR(13)+CHR(10)+;
   'ActMEEss.xls',0+32,'')
  RETURN 
 ENDIF 

 IF IsUsrDir=.T.
  m.usrdir = fso.GetParentFolderName(pbin) + '\'+UPPER(m.gcuser)
  IF !fso.FolderExists(m.usrdir)
   MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+UPPER(ALLTRIM(m.usrdir))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
  IF !fso.FolderExists(m.usrdir+'\SSACTS')
   MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+UPPER(ALLTRIM(m.usrdir+'\SSACTS'))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
  IF !fso.FolderExists(m.usrdir+'\SVACTS')
   MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+UPPER(ALLTRIM(m.usrdir+'\SSACTS'))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
 ELSE 
  IF !fso.FolderExists(pmee)
   MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+UPPER(ALLTRIM(pmee))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
 ENDIF 
 
 m.sn_pol   = para1
 m.TipOfExp = para5

 DO CASE 
  CASE m.TipOfExp = '2'
   m.ctipofexp = 'плановой'
  CASE m.TipOfExp = '3'
   m.ctipofexp = 'целевой'
  OTHERWISE 
   m.ctipofexp = ''
 ENDCASE 

 IF !EMPTY(goApp.d_exp)
  m.d_exp = DTOC(goApp.d_exp)
 ELSE 
  m.d_exp = IIF(m.qcod!='I3', DTOC(DATE()), '')
 ENDIF 

 m.flcod   = goApp.flcod
 m.mcod    = goApp.mcod 
 m.lpuid   = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)
 m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.fullname)+', '+m.mcod, '')
 m.TipOfPeriod = IIF(EMPTY(m.flcod),0,1) && 0-локальный период, 1 - сводный!

 m.exp_dat1 = '01.'+PADL(tMonth,2,'0')+'.'+STR(tYear,4)
 m.exp_dat2 = DTOC(GOMONTH(CTOD(m.exp_dat1),1)-1)

 m.ppolis = STRTRAN(ALLTRIM(m.sn_pol),' ','') && Для названия Акта
 
 SELECT talon 
 
 COUNT FOR sn_pol = m.sn_pol AND SEEK(recid, 'merror', 'recid') AND merror.et = m.TipOfExp TO m.nIsExps

 IF m.nIsExps == 0
  MESSAGEBOX(CHR(13)+CHR(10)+'ПО ВЫБРАННОМУ СЧЕТУ МЭЭ НЕ ПРОВОДИЛОСЬ!'+CHR(13)+CHR(10),0+64,'')
  RETURN
 ELSE 
 ENDIF 

 CREATE CURSOR ttalon (nrec i AUTOINC, sn_pol c(25),c_i c(30),ds c(6),tip c(1),d_u d,pcod c(10),otd c(4),cod n(6),k_u n(3),;
  d_type c(1),s_all n(11,2),e_cod n(6),e_ku n(3),e_tip c(1),err_mee c(3),e_period c(6),et c(1),s_1 n(11,2),;
  koeff n(4,2), ee c(1), straf n(11,2))

 SELECT talon 

 m.TipOfMee = 0
 
 m.s_lech = 0
 m.dat1 = {31.12.2099}
 m.dat2 = {01.01.2000}
 m.dslast = 0
 SCAN FOR sn_pol == m.sn_pol
  IF USED('serror')
   IF SEEK(recid, 'serror')
    LOOP 
   ENDIF 
  ENDIF 
  IF !SEEK(PADL(recid,6,'0')+m.TipOfExp, 'merror', 'id_et') && PADL(recid,6,'0')+et+docexp+reason
   LOOP 
  ENDIF 

  SCATTER FIELDS EXCEPT recid,mcod,period,sn_pol,q,novor,ds_s,ds_p,profil,rslt,prvs,ord,;
   ishod,recid_lpu,ispr MEMVAR 
  
  m.et       = merror.et
  m.ee       = merror.ee
  m.e_cod    = merror.e_cod
  m.e_ku     = merror.e_ku
  m.e_tip    = merror.e_tip
  m.err_mee  = merror.err_mee 
  m.osn230   = merror.osn230
  m.koeff    = merror.koeff
  m.e_period = merror.e_period
  m.straf    = merror.straf
  m.s_1      = merror.s_1
  m.s_2      = merror.s_2
   
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
  
  INSERT INTO ttalon FROM MEMVAR 
 ENDSCAN 
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
 
 SELECT recid FROM ssacts WHERE period=m.gcperiod AND mcod=m.mcod AND flcod=m.flcod AND ;
  codexp=INT(VAL(m.TipOfExp)) AND tipacc=m.tipacc AND sn_pol=PADR(STRTRAN(m.sn_pol,' ',''),25) ;
  INTO CURSOR rqwest NOCONSOLE 
 m.nfileid = recid
 USE 
 
 IF m.nfileid>0
  DocName = IIF(!IsUsrDir, m.pmee, m.usrdir)+'\ssacts\'+PADL(m.nfileid,6,'0')
 ELSE 
  INSERT INTO ssacts (period,mcod,flcod,codexp,tipacc,sn_pol,fam,im,ot) ;
   VALUES ;
  (m.gcperiod,m.mcod,m.flcod,INT(VAL(m.tipofexp)),m.tipacc,;
   PADR(STRTRAN(m.sn_pol,' ',''),25),people.fam,people.im,people.ot)
  m.nfileid = GETAUTOINCVALUE()
  DocName = IIF(!IsUsrDir, m.pmee, m.usrdir)+'\ssacts\'+PADL(m.nfileid,6,'0')
  UPDATE ssacts SET actname=PADL(m.nfileid,6,'0')+'.xls', actdate=DATETIME() WHERE recid = m.nfileid
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

   USE IN ttalon
   RETURN
  ELSE 
   UPDATE ssacts SET actdate=DATETIME() WHERE recid = m.nfileid
  ENDIF 
  ELSE 
   UPDATE ssacts SET actdate=DATETIME() WHERE recid = m.nfileid
  ENDIF 

 ENDIF 

 m.n_akt = mcod + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)+'/'+ALLTRIM(STR(m.nfileid))
 DO CASE 
  CASE m.TipOfExp = '2'
   m.podvid='0'
  CASE m.TipOfExp = '3'
   m.podvid='1'
  CASE m.TipOfExp = '4'
   m.podvid='0'
  CASE m.TipOfExp = '5'
   m.podvid='1'
  CASE m.TipOfExp = '6'
   m.podvid='Т'
  CASE m.TipOfExp = '7'
   m.podvid='Т'
  OTHERWISE 
   m.podvid='0'
 ENDCASE 
 IF m.TipOfPeriod=0
  m.n_akt = m.qcod+STR(m.lpuid,4)+IIF(INLIST(m.TipOfExp,'2','3','7'),'1','2')+;
   IIF(INLIST(m.TipOfExp,'2','4','6','7','8'),'1','2')+m.podvid+ALLTRIM(STR(m.nfileid))
 ELSE 
  m.n_akt = m.qcod+STR(m.lpuid,4)+IIF(INLIST(m.TipOfExp,'2','3','7'),'1','2')+;
   IIF(INLIST(m.TipOfExp,'2','4','6','7','9'),'1','2')+m.podvid+ALLTRIM(STR(m.nfileid))
 ENDIF 
 m.d_akt = IIF(m.qcod!='I3', DTOC(DATE()), '')
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
 SCAN 
  m.er_c = err_mee
  m.osn230 = IIF(SEEK(LEFT(UPPER(m.er_c),2), 'sookod'), sookod.osn230, '')	

  m.tot_badsum = m.tot_badsum + s_1
  m.tot_goodsum = m.tot_goodsum + (s_all-s_1)

  nRow = nRow + 1
    
  m.tot_straf = m.tot_straf + IIF( m.tot_straf<=0, straf*m.ynorm, 0)

 ENDSCAN 
* USE 

 m.koplate   = m.s_lech - m.tot_badsum
 m.nekoplate = m.tot_badsum

 IF m.TipOfMee=1
  m.resume = 'По представленному счету замечаний нет.'
  m.vivod  = 'Счет подлежит оплате в полном объеме из средств ОМС.'
 ENDIF 

 LOCAL m.lcTmpName, m.lcRepName, m.lcDbfName, m.llResult
 m.lcTmpName = pTempl+'\'+m.dotname
 m.lcRepName = m.docname+'.xls'

 m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)

 WAIT CLEAR 

RETURN 

