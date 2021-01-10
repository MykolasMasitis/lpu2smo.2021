FUNCTION MakeMEESV(lcPath, para3, IsQuit, para4, para5) && ¬ÂÒËˇ ‰Îˇ »Ì„ÓÒÒ‡!
 
 m.IsVisible   = para3
 m.tipofexp    = para4
 m.TipOfPeriod = para5 && 0-ÎÓÍ‡Î¸Ì˚È ÔÂËÓ‰, 1 - Ò‚Ó‰Ì˚È!
 
 IF IsUsrDir=.T.
  m.usrdir = fso.GetParentFolderName(pbin) + '\'+UPPER(m.gcuser)
  IF !fso.FolderExists(m.usrdir)
   MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ '+UPPER(ALLTRIM(m.usrdir))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
  IF !fso.FolderExists(m.usrdir+'\SSACTS')
   MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ '+UPPER(ALLTRIM(m.usrdir+'\SSACTS'))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
  IF !fso.FolderExists(m.usrdir+'\SVACTS')
   MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ '+UPPER(ALLTRIM(m.usrdir+'\SSACTS'))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
 ELSE 
  IF !fso.FolderExists(pmee)
   MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ '+UPPER(ALLTRIM(pmee))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
 ENDIF 

 DotName = 'ActMEEsvI3.xls'
 IF !fso.FileExists(pTempl+'\'+DotName)
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À ÿ¿¡ÀŒÕ Œ“◊≈“¿'+CHR(13)+CHR(10)+;
   'ActMEEsvI3.xls',0+32,'')
  RETURN 
 ENDIF 
 
 oal     = ALIAS() && Aisoms
 m.mcod  = SUBSTR(lcpath,RAT('\',lcpath)+1)
 
 IF m.TipOfPeriod=0 && ÎÓÍ‡Î¸Ì˚È ÔÂËÓ‰
  m.exp_dat1 = '01.'+PADL(tMonth,2,'0')+'.'+STR(tYear,4)
  m.exp_dat2 = DTOC(GOMONTH(CTOD(m.exp_dat1),1)-1)
 ELSE  && ÔÓËÁ‚ÓÎ¸Ì˚È ÔÂËÓ‰
  m.exp_dat1 = DTOC(flmindate(m.flcod))
  m.exp_dat2 = DTOC(flmaxdate(m.flcod))
 ENDIF 

 IF !EMPTY(goApp.d_exp)
  m.edat1    = DTOC(goApp.d_exp)
  m.edat2    = m.edat1  
 ELSE 
  m.edat1    = DTOC(DATE())
  m.edat2    = m.edat1  
 ENDIF 

 m.lpuid   = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)
 m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
 m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.fullname)+', '+m.mcod, '')
 m.lpuboss    = IIF(SEEK(m.lpuid, 'lpudogs'), ALLTRIM(lpudogs.boss), '')
 
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

 m.lpudog  = IIF(SEEK(m.lpuid, 'lpudogs'), ALLTRIM(lpudogs.dogs), '')
 m.lpuddog = IIF(SEEK(m.lpuid, 'lpudogs'), lpudogs.ddogs, {})
 m.lpucdog = m.lpudog+' ÓÚ '+DTOC(m.lpuddog)

 m.dschet    = IIF(FIELD('processed', 'aisoms')='PROCESSED', TTOC(aisoms.processed), '')+', ÌÓÏÂ Ò˜ÂÚ‡ '+STR(tYear,4)+PADL(tMonth,2,'0')

 IF m.TipOfPeriod=0
  m.flcod = ''
  pPath = pBase+'\'+gcPeriod+'\'+m.mcod
  TFile = 'talon'
  mFile = 'm'+m.mcod
 ELSE 
  m.flcod = flcod
  pPath = pBase+'\'+gcPeriod+'\0000000\'+m.mcod
  TFile = 't'+m.flcod
  mFile = 'm'+m.flcod
 ENDIF 

 IF OpenFile(pPath+'\'+TFile, 'Talon', 'SHARED', 'recid')>0
  IF USED('talon')
   USE IN talon 
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pPath+'\'+mFile, 'merror', 'SHARED')>0
  IF USED('merror')
   USE IN merror
  ENDIF 
  IF USED('talon')
   USE IN talon 
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT merror
 m.nexps=0

 COUNT FOR et=m.tipofexp TO m.nexps

 IF m.nexps=0
  MESSAGEBOX(CHR(13)+CHR(10)+'œŒ ¬€¡–¿ÕÕŒÃ” Àœ” › —œ≈–“»«¿'+CHR(13)+CHR(10)+;
   '«¿ƒ¿ÕÕŒ√Œ “»œ¿ Õ≈ œ–Œ¬Œƒ»À¿—‹!',0+64,'')
  IF USED('talon')
   USE IN talon 
  ENDIF 
  IF USED('merror')
   USE IN merror
  ENDIF 
  SELECT &oal
  RETURN 
 ENDIF 
 
 SELECT a.sn_pol FROM talon a, merror b;
	WHERE a.recid=b.recid AND b.et=m.TipOfExp GROUP BY sn_pol INTO CURSOR curallaccs READWRITE 
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 SELECT a.sn_pol FROM talon a, merror b;
	WHERE a.recid=b.recid AND b.et=m.TipOfExp AND err_mee<>'W0' GROUP BY sn_pol INTO CURSOR curdefaccs READWRITE 
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR curokaccs (sn_pol c(25))
 SELECT curokaccs
 INDEX on sn_pol TAG sn_pol 

 SELECT curallaccs
 SET RELATION TO sn_pol INTO curdefaccs
 SCAN FOR EMPTY(curdefaccs.sn_pol)
  SCATTER MEMVAR 
  INSERT INTO curokaccs FROM MEMVAR 
 ENDSCAN 
 SET RELATION OFF INTO curdefaccs
 USE IN curdefaccs
 USE IN curallaccs
 
 IF RECCOUNT('curokaccs') <= 0
  MESSAGEBOX(CHR(13)+CHR(10)+'œŒ ¬€¡–¿ÕÕŒÃ” Àœ” › —œ≈–“»«¿'+CHR(13)+CHR(10)+;
   '«¿ƒ¿ÕÕŒ√Œ “»œ¿ ¡≈«Œÿ»¡Œ◊Õ€’ —◊≈“Œ¬ Õ≈ ¬€ﬂ¬»À¿!'+CHR(13)+CHR(10)+;
   '—‘Œ–Ã»–”…“≈ ¿ “€ —“–¿’Œ¬€’ —À”◊¿≈¬!',0+64,'')
  USE IN curokaccs
  IF USED('talon')
   USE IN talon 
  ENDIF 
  IF USED('merror')
   USE IN merror
  ENDIF 
  SELECT &oal
  RETURN 
 ENDIF 

 *SELECT a.docexp FROM merror a, talon b, curokaccs c WHERE a.recid=b.recid AND b.sn_pol=c.sn_pol AND ;
 	et=m.TipOfExp GROUP BY docexp INTO CURSOR curexps
 SELECT a.usr FROM merror a, talon b, curokaccs c WHERE a.recid=b.recid AND b.sn_pol=c.sn_pol AND ;
 	et=m.TipOfExp GROUP BY usr INTO CURSOR curexps
 	
 SCAN    && ÷ËÍÎ ÔÓ ˝ÍÒÔÂÚ‡Ï —ÃŒ
  *m.docexp = docexp
  m.usr = usr
  *m.docfio = IIF(SEEK(m.docexp, 'explist'), ;
   ALLTRIM(explist.fam)+' '+ALLTRIM(explist.im)+' '+ALLTRIM(explist.ot)+', ÍÓ‰ '+m.docexp, '')
  m.docfio = ''

  *=OneSvActA(m.docexp)
  IF m.usr<>goApp.smoexp AND goApp.smoexp<>m.SuperExp
   LOOP 
  ENDIF 

  *MESSAGEBOX(m.usr,0+64,goApp.smoexp)
  =OneSvActA(m.usr)

 ENDSCAN && ÷ËÍÎ ÔÓ ˝ÍÒÔÂÚ‡Ï —ÃŒ
 USE IN curexps
 USE IN curokaccs
 
 IF USED('talon') 
  USE IN talon 
 ENDIF 
 IF USED('merror') 
  USE IN merror
 ENDIF 

 SELECT &oal

RETURN 

FUNCTION OneSvActA(m.paraexp)

 *m.docexp = m.paraexp
 m.usr = m.paraexp
 *SELECT a.reason AS reason FROM merror a, talon b, curokaccs c WHERE a.recid=b.recid AND b.sn_pol=c.sn_pol AND ;
 	et=m.TipOfExp AND docexp=m.docexp GROUP BY reason INTO CURSOR currss
 SELECT a.reason AS reason FROM merror a, talon b, curokaccs c WHERE a.recid=b.recid AND b.sn_pol=c.sn_pol AND ;
 	et=m.TipOfExp AND usr=m.usr GROUP BY reason INTO CURSOR currss
 
 SELECT currss
 SCAN    && ÷ËÍÎ ÔÓ ÔÓ‚Ó‰‡Ï/ÚÂÏ‡Ï ˝ÍÒÔÂÚËÁ˚ ‚ÌÛÚË Ó‰ÌÓ„Ó ÍÓ‰‡ ˝ÍÒÔÂÚ‡ —ÃŒ
  m.reason = reason
  *=OneSvActAB(m.docexp, m.reason)
  =OneSvActAB(m.usr, m.reason)
 ENDSCAN && ÷ËÍÎ ÔÓ ÔÓ‚Ó‰‡Ï/ÚÂÏ‡Ï ˝ÍÒÔÂÚËÁ˚ ‚ÌÛÚË Ó‰ÌÓ„Ó ÍÓ‰‡ ˝ÍÒÔÂÚ‡ —ÃŒ
 USE IN currss

 SELECT curexps

RETURN 

FUNCTION OneSvActAB(m.paraexp, m.paraexp2)

 *PRIVATE m.docexp, m.reason
 PRIVATE m.usr, m.reason
 
 *m.docexp = m.paraexp
 m.usr    = m.paraexp
 m.reason = m.paraexp2

 ooal = ALIAS()

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
 IF fso.FileExists(pmee+'\svacts\moves'+'.dbf')
  m.rtal = ALIAS()
  IF OpenFile(pmee+'\svacts\moves', 'moves', 'shar')>0
   IF USED('moves')
    USE IN moves
   ENDIF 
  ENDIF 
  SELECT (rtal)
 ENDIF 
 
 *SELECT recid, resume, conclusion, recommend FROM svacts WHERE period=m.gcperiod AND mcod=m.mcod AND codexp=INT(VAL(m.TipOfExp)) ;
   AND docexp=m.docexp AND flcod=m.flcod AND reason=m.reason INTO CURSOR rqwest NOCONSOLE  
 SELECT recid, resume, conclusion, recommend FROM svacts WHERE period=m.gcperiod AND mcod=m.mcod AND codexp=INT(VAL(m.TipOfExp)) ;
   AND smoexp=m.usr AND flcod=m.flcod AND reason=m.reason INTO CURSOR rqwest NOCONSOLE  
 
 *MESSAGEBOX('period: '+m.gcperiod+CHR(13)+CHR(10)+;
 	'mcod: '+m.mcod+CHR(13)+CHR(10)+;
 	'codexp: '+m.TipOfExp+CHR(13)+CHR(10)+;
 	'smoexp: '+m.docexp+CHR(13)+CHR(10)+;
 	'flcod: '+m.flcod+CHR(13)+CHR(10)+;
 	'reason: '+m.reason,0+64,'')
 
 m.nfileid    = recid
 m.resume     = resume
 m.conclusion = conclusion
 m.recommend  = recommend

 USE 
 SELECT (ooal)

 IF m.nfileid>0
  DocName = IIF(!IsUsrDir, m.pmee, m.usrdir)+'\svacts\'+PADL(m.nfileid,6,'0')
  m.n_akt = NumActOfExp(m.lpuid,m.tipofexp,m.reason,m.nfileid)
 ELSE 
   SELECT TOP 1 resume, conclusion, recommend FROM svacts ORDER BY recid DESC INTO CURSOR cqwert
   m.resume     = resume
   m.conclusion = conclusion
   m.recommend  = recommend
   USE IN cqwert
   SELECT (ooal)

   *INSERT INTO svacts (n_rss,period,mcod,lpu_id,codexp,flcod,e_period,smoexp,et,reason,qr,status) ;
    VALUES ;
   (m.n_rss,m.gcperiod,m.mcod,m.lpuid,INT(VAL(m.tipofexp)),m.flcod,m.e_period,goApp.smoexp,m.tipofexp,m.reason,.T.,'1')
   INSERT INTO svacts (n_rss,period,mcod,lpu_id,codexp,flcod,smoexp,et,reason,qr,status) ;
    VALUES ;
   (m.n_rss,m.gcperiod,m.mcod,m.lpuid,INT(VAL(m.tipofexp)),m.flcod,goApp.smoexp,m.tipofexp,m.reason,.T.,'1')

   m.nfileid = GETAUTOINCVALUE()
   m.n_akt = NumActOfExp(m.lpuid,m.tipofexp,m.reason,m.nfileid)

   DocName = IIF(!IsUsrDir, m.pmee, m.usrdir)+'\svacts\'+PADL(m.nfileid,6,'0')
   UPDATE svacts SET actname=PADL(m.nfileid,6,'0')+'.xls', actdate=DATETIME(), n_akt=m.n_akt, qr=.t. WHERE recid = m.nfileid
 ENDIF 
  
 IF fso.FileExists(DocName+'.xls')
  oFile = fso.GetFile(DocName+'.xls')
  DateCreated      = TTOC(oFile.DateCreated)
  DateLastAccessed = TTOC(oFile.DateLastAccessed)
  DateLastModified = TTOC(oFile.DateLastModified)
  RELEASE oFile
  
  IF MESSAGEBOX('œŒ ¬€¡–¿ÕÕŒÃ” Àœ” ¿ “ ”∆≈ ‘Œ–Ã»–Œ¬¿À—ﬂ!'+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ƒ¿“¿ —Œ«ƒ¿Õ»ﬂ ¿ “¿            : '+m.DateCreated+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ƒ¿“¿ œŒ—À≈ƒÕ≈√Œ Œ“ –€“»ﬂ ¿ “¿ : '+m.DateLastAccessed+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ƒ¿“¿ œŒ—À≈ƒÕ≈√Œ »«Ã≈Õ≈Õ»ﬂ ¿ “¿: '+m.DateLastModified+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   '¬€ ’Œ“»“≈ œ≈–≈‘Œ–Ã»–Œ¬¿“‹ ¿ “?',4+32,'') == 7 
   
   USE IN talon 
   USE IN merror
   IF USED('rrfile')
    USE IN rrfile
   ENDIF 
   IF USED('moves')
    USE IN moves
   ENDIF 
   SELECT aisoms
   RETURN
  ELSE && IF MESSAGEBOX('œŒ ¬€¡–¿ÕÕŒÃ” Àœ”
   IF m.TipOfPeriod=0
    UPDATE svacts SET actdate=DATETIME() WHERE recid = m.nfileid
    IF USED('moves')
     UPDATE moves SET dat=DATETIME() WHERE actid = m.nfileid AND et='1'
     IF _tally=0
      INSERT INTO moves (actid,et,usr,dat) VALUES (m.nfileid,'1',m.gcUser,DATETIME())
     ENDIF 
    ENDIF 
   ENDIF 
  ENDIF 
 ELSE && IF fso.FileExists(DocName+'.xls')
  IF USED('moves')
   INSERT INTO moves (actid,et,usr,dat) VALUES (m.nfileid,'1',m.gcUser,DATETIME())
  ENDIF 
 ENDIF 

 DO FORM TxtForActs

 UPDATE svacts SET resume=ALLTRIM(m.resume), conclusion=ALLTRIM(m.conclusion), recommend=ALLTRIM(m.recommend) WHERE recid = m.nfileid

 m.IsExpMee = .f.

 m.checked_tot = 0
 m.checked_amb = 0
 m.checked_dst = 0
 m.checked_st  = 0
 m.checked_02  = 0

 *m.bad_kol   = 0
 *m.bad_sum   = 0
 m.opl_tot   = 0

 IF INLIST(m.TipOfExp,'7')
  m.povod    = IIF(SEEK(m.gcUser+m.reason, 'themes'), themes.name, '')
 ELSE 
  m.povod    = IIF(SEEK(m.reason, 'reasons'), reasons.name, '')
 ENDIF 

 m.aktname  = ''
 m.vidofexp = ''
 DO CASE 
  CASE m.TipOfExp = '2'
   m.aktname='¿ÍÚ ÔÎ‡ÌÓ‚ÓÈ ÏÂ‰ËÍÓ-˝ÍÓÌÓÏË˜ÂÒÍÓÈ ˝ÍÒÔÂÚËÁ˚ π'
   m.vidofexp='ÔÎ‡ÌÓ‚‡ˇ Ã››'
  CASE m.TipOfExp = '3'
   m.aktname='¿ÍÚ ˆÂÎÂ‚ÓÈ ÏÂ‰ËÍÓ-˝ÍÓÌÓÏË˜ÂÒÍÓÈ ˝ÍÒÔÂÚËÁ˚ π'
   m.vidofexp='ˆÂÎÂ‚‡ˇ Ã››'
  CASE m.TipOfExp = '7'
   m.aktname='¿ÍÚ ÚÂÏ‡ÚË˜ÂÒÍÓÈ ÏÂ‰ËÍÓ-˝ÍÓÌÓÏË˜ÂÒÍÓÈ ˝ÍÒÔÂÚËÁ˚ π'
   m.vidofexp='ÚÂÏ‡ÚË˜ÂÒÍ‡ˇ Ï˝˝'
  CASE m.TipOfExp = '8'
   m.aktname='¿ÍÚ ÏÂ‰ËÍÓ-˝ÍÓÌÓÏË˜ÂÒÍÓÈ ˝ÍÒÔÂÚËÁ˚ ÔÓ Ê‡ÎÓ·Â π'
   m.vidofexp='Ï˝˝ ÔÓ Ê‡ÎÓ·Â'
  OTHERWISE 
 ENDCASE 

 m.d_akt = IIF(m.qcod!='I3', DTOC(DATE()), '')
  
 m.cpredps = 'œÂ‰ÔËÒ‡ÌËÂ π ' + m.n_akt + ' ÓÚ ' + m.d_akt
 
 CREATE CURSOR qwert (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
  
 CREATE CURSOR qwertamb (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR qwertst (c_i c(30))
 INDEX on c_i TAG c_i
 SET ORDER TO c_i

 CREATE CURSOR qwert02 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR qwertdst (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR curdata (nrec n(5), sn_pol c(25), c_i c(30), d_beg d, d_end d, ds c(6), cod n(6), s_all n(11,2), ;
  osn230 c(5), er_c c(3), delta n(11,2), straf n(11,2), cmnt c(50))

 SELECT merror
 SET RELATION TO recid INTO talon 
 
 m.nrec  = 1
 SCAN 
  *IF !(et=m.TipOfExp AND docexp=m.docexp AND reason=m.reason)
  IF !(et=m.TipOfExp AND usr=m.usr AND reason=m.reason)
   LOOP 
  ENDIF 
  m.er_c = err_mee
  IF LEFT(UPPER(m.er_c),2) != 'W0'
   LOOP 
  ENDIF 
  
  REPLACE s_1 WITH 0, s_2 WITH 0

  m.sn_pol = talon.sn_pol
  IF !SEEK(m.sn_pol, 'curokaccs')
   LOOP 
  ENDIF 
  m.cod    = cod
  m.c_i    = talon.c_i

  REPLACE n_akt WITH m.n_akt, d_akt WITH goApp.d_exp, t_akt WITH 'SV'
  
  IF USED('rrfile')
   IF SEEK(m.sn_pol, 'rrfile')
    REPLACE n_akt WITH m.n_akt, d_akt WITH goApp.d_exp, t_akt WITH 'SV' IN rrfile
   ENDIF 
  ENDIF 

  IF !SEEK(m.sn_pol, 'qwert')
   INSERT INTO qwert (sn_pol) VALUES (m.sn_pol)
   m.checked_tot = m.checked_tot + 1
  ENDIF 
   
  IF IsUsl(m.cod) AND !SEEK(m.sn_pol, 'qwertamb')
   INSERT INTO qwertamb (sn_pol) VALUES (m.sn_pol)
   m.checked_amb = m.checked_amb + 1
  ENDIF 
   
  IF IsKD(m.cod) AND !SEEK(m.sn_pol, 'qwertdst')
   INSERT INTO qwertdst (sn_pol) VALUES (m.sn_pol)
   m.checked_dst = m.checked_dst + 1
  ENDIF 
   
  IF (IsMes(m.cod) OR IsVMP(m.cod)) AND !SEEK(m.c_i, 'qwertst')
   INSERT INTO qwertst (c_i) VALUES (m.c_i)
   m.checked_st = m.checked_st + 1
  ENDIF 

  IF (Is02(m.cod)) AND !SEEK(m.sn_pol, 'qwert02')
   INSERT INTO qwert02 (sn_pol) VALUES (m.sn_pol)
   m.checked_02 = m.checked_02 + 1
  ENDIF 

  m.osn230 = osn230
  m.d_beg  = IIF(!IsMes(m.cod) and !IsVMP(m.cod), talon.d_u, talon.d_u-talon.k_u+1)
  m.d_end  = talon.d_u

  m.ds     = talon.ds   
  m.ns_all = 0
  m.s_all  = s_all

  m.opl_tot = m.opl_tot + s_all && »ÁÏÂÌÂÌÓ 12.09.12 ÔÓ Á‡ÏÂ˜‡ÌË˛ —Œ√¿« - ÌÂ‚ÂÌ‡ˇ ÒÛÏÏ‡ ‚ Ò‚Ó‰ÌÓÏ ‡ÍÚÂ "Í ÓÔÎ‡ÚÂ"
  
  m.cmnt = '«‡ÏÂ˜‡ÌËÈ ÌÂÚ'

  INSERT INTO curdata FROM MEMVAR 

  m.nrec = m.nrec+1

 ENDSCAN 
 SET RELATION OFF INTO talon 

 USE IN qwert
 USE IN qwertamb
 USE IN qwertst
 USE IN qwertdst
 USE IN qwert02
 
 SELECT (ooal)

 m.checked_tot = m.checked_amb + m.checked_st + m.checked_dst + m.checked_02

 *UPDATE svacts SET n_ss=m.checked_tot, n_st=m.checked_st, n_dst=m.checked_dst, n_plk=m.checked_amb WHERE recid = m.nfileid
 UPDATE svacts SET n_ss=m.checked_tot, n_st=m.checked_st, n_dst=m.checked_dst, n_plk=m.checked_amb, n_02=m.checked_02;
 	WHERE recid = m.nfileid
 
 LOCAL m.lcTmpName, m.lcRepName, m.lcDbfName, m.llResult
 m.lcTmpName = pTempl+'\'+m.dotname
 m.lcRepName = m.docname+'.xls'

 IF m.qcod!='I3'
  m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)
 ELSE 
  m.llResult = X_Report(m.lcTmpName, m.lcRepName, .F.)
  PUBLIC oExcel AS Excel.Application
  WAIT "«‡ÔÛÒÍ MS Excel..." WINDOW NOWAIT 
  TRY 
   oExcel=GETOBJECT(,"Excel.Application")
  CATCH 
   oExcel=CREATEOBJECT("Excel.Application")
  ENDTRY 
  WAIT CLEAR 

  *m.acthead = '1'+PADR(STRTRAN(m.n_akt,'“','T'),16)+PADR(goApp.smoexp,7)+PADR(m.docexp,7)+IIF(m.tipofexp='3', m.reason, ' ')
  m.acthead = '1'+PADR(STRTRAN(m.n_akt,'“','T'),16)+PADR(goApp.smoexp,7)+PADR(m.usr,7)+IIF(m.tipofexp='3', m.reason, ' ')
  m.actbody = ''
  m.actbody = m.actbody + IIF(m.checked_amb>0, '02'+PADR(m.checked_amb,4), '')
  m.actbody = m.actbody + IIF(m.checked_st>0, '03'+PADR(m.checked_st,4), '')
  m.actbody = m.actbody + IIF(m.checked_dst>0, '04'+PADR(m.checked_dst,4), '')
  m.actbody = m.actbody + IIF(m.checked_02>0, '05'+PADR(m.checked_02,4), '')

  m.barcode = m.acthead+m.actbody

  lcQRImage = loFbc.QRBarcodeImage(m.barcode,pMee+'\svacts\'+PADL(m.nfileid,6,'0')+'.png',6,2)
 
  OneBook  = oExcel.Workbooks.Add(m.docname)
  oSheet=OneBook.ActiveSheet

  oPic = oSheet.Range("A1").Parent.Pictures.Insert(pMee+'\svacts\'+PADL(m.nfileid,6,'0')+'.png')
  IF fso.FileExists(pMee+'\svacts\'+PADL(m.nfileid,6,'0')+'.png')
   fso.DeleteFile(pMee+'\svacts\'+PADL(m.nfileid,6,'0')+'.png')
  ENDIF 
  
  WITH oSheet.Range("I2:I6")
   oPic.Top = .Top
   oPic.Left = .Left
   oPic.Height = .Height
  ENDWITH 
  
  osheet.Protect('qwerty',,.t.)
 
  IF fso.FileExists(m.docname+'.xls')
   TRY 
    fso.DeleteFile(m.docname+'.xls')
    OneBook.SaveAs(m.docname,18)
   CATCH  
    MESSAGEBOX('‘¿…À '+m.docname+'.XLS Œ “–€“!',0+64,'')
   ENDTRY 
  ELSE 
   oBook.SaveAs(m.docname,18)
  ENDIF 
  oExcel.Visible = m.IsVisible
 ENDIF 


 USE IN curdata 
 IF USED('rrfile')
  USE IN rrfile
 ENDIF 
 IF USED('moves')
  USE IN moves
 ENDIF 

 SELECT aisoms

RETURN 