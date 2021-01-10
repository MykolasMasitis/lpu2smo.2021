FUNCTION MakeEKMPSVn(parra1, parra2, parra3, parra4, parra5)

 m.lcpath      = parra1
 m.IsVisible   = parra2
 m.IsQuit      = parra3
 m.tipofexp    = parra4
 m.TipOfPeriod = parra5  && 0-ÎÓÍ‡Î¸Ì˚È ÔÂËÓ‰, 1 - Ò‚Ó‰Ì˚È!
 
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

 m.expname  = '¿ÍÚ '
 m.vidofexp = ''

 DO CASE 
  CASE m.tipofexp = '4'
   m.expname = m.expname + 'ÔÎ‡ÌÓ‚ÓÈ'
   m.vidofexp = 'ÔÎ‡ÌÓ‚‡ˇ › Ãœ'
  CASE m.tipofexp = '5'
   m.expname = m.expname + 'ˆÂÎÂ‚ÓÈ'
   m.vidofexp = 'ˆÂÎÂ‚‡ˇ › Ãœ'
  CASE m.tipofexp = '6'
   m.expname = m.expname + 'ÚÂÏ‡ÚË˜ÂÒÍÓÈ'
   m.vidofexp = 'ÚÂÏ‡ÚË˜ÂÒÍ‡ˇ › Ãœ'
  CASE m.tipofexp = '9'
   m.expname = m.expname + 'ÚÂÏ‡ÚË˜ÂÒÍÓÈ'
   m.vidofexp = '› Ãœ ÔÓ Ê‡ÎÓ·Â'
 ENDCASE 

 m.expname = m.expname + ' ˝ÍÒÔÂÚËÁ˚ Í‡˜ÂÒÚ‚‡ ÏÂ‰ËˆËÌÒÍÓÈ ÔÓÏÓ˘Ë'
 m.flcod   = goApp.flcod

 DotName = 'EKMP_N.xls'

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

 IF !fso.FileExists(pTempl+'\'+DotName)
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À ÿ¿¡ÀŒÕ Œ“◊≈“¿'+CHR(13)+CHR(10)+;
   'EKMP_N.xls',0+32,'')
  RETURN 
 ENDIF 
 
 m.mcod       = goApp.mcod 
 m.Is02 = IIF(m.mcod='0371001', .T., .F.)

 m.lpuid      = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)
 m.IsVed      = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
 m.lpuaddress = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.adres), '')
 m.lpuname    = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.fullname)+', '+m.mcod+', '+m.lpuaddress, '')
 m.lpudog     = IIF(SEEK(m.lpuid, 'lpudogs'), ALLTRIM(lpudogs.dogs), '')
 m.lpuddog    = IIF(SEEK(m.lpuid, 'lpudogs'), lpudogs.ddogs, {})
 m.lpudog     = '‚ ÒÓÓÚ‚ÂÚÒÚ‚ËË Ò ƒÓ„Ó‚ÓÓÏ '+m.lpudog+' ÓÚ '+DTOC(m.lpuddog)
 m.lpuboss    = IIF(SEEK(m.lpuid, 'lpudogs'), ALLTRIM(lpudogs.boss), '')
 IF FIELD('SENT')='SENT'
  m.sent       = sent
 ELSE 
  m.sent       = DATETIME()
 ENDIF 
 m.dexp1 = DTOC(m.tdat1)
 m.dexp2 = DTOC(m.tdat2)

 IF m.TipOfPeriod=0
  pPath = pBase+'\'+gcPeriod+'\'+m.mcod
  TFile = 'talon'
  PFile = 'people'
  mFile = 'm'+m.mcod
 ELSE 
  m.flcod = aisoms.flcod
  pPath = pBase+'\'+gcPeriod+'\0000000\'+m.mcod
  TFile = 't'+m.flcod
  TFile = 'p'+m.flcod
  mFile = 'm'+m.flcod
 ENDIF 

 IF OpenFile(pPath+'\'+TFile, 'Talon', 'SHARED', 'recid')>0
  IF USED('talon')
   USE IN talon 
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pPath+'\'+PFile, 'People', 'SHARED', 'sn_pol')>0
  USE IN talon 
  IF USED('people')
   USE IN people
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
  IF USED('people')
   USE IN people
  ENDIF 
  RETURN 
 ENDIF 

 SELECT merror
 
 COUNT FOR !EMPTY(err_mee) AND et=m.TipOfExp TO m.nIsEkmp
 
 IF m.nIsEkmp<=0
  MESSAGEBOX(CHR(13)+CHR(10)+'œŒ ¬€¡–¿ÕÕŒÃ” Àœ” › Ãœ Õ≈ œ–Œ¬Œƒ»À¿—‹!'+CHR(13)+CHR(10),0+32,'')
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('talon')
   USE IN talon 
  ENDIF 
  IF USED('merror')
   USE IN merror
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 

 SELECT a.sn_pol FROM talon a, merror b;
	WHERE a.recid=b.recid AND b.et=m.TipOfExp GROUP BY sn_pol INTO CURSOR curallaccs READWRITE 
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 SELECT a.sn_pol FROM talon a, merror b;
	WHERE a.recid=b.recid AND b.et=m.TipOfExp AND err_mee!='W0' GROUP BY sn_pol INTO CURSOR curdefaccs READWRITE 
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
   '«¿ƒ¿ÕÕŒ√Œ “»œ¿ ¡≈«Œÿ»¡Œ◊Õ€’ —◊≈“Œ¬ Õ≈¬€ﬂ¬»À¿!'+CHR(13)+CHR(10)+;
   '—‘Œ–Ã»–”…“≈ ¿ “€ —“–¿’Œ¬€’ —À”◊¿≈¬!',0+64,'')
  USE IN curokaccs
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('talon')
   USE IN talon 
  ENDIF 
  IF USED('merror')
   USE IN merror
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
 
 *SELECT a.docexp,a.reason as reason FROM merror a, talon b, curokaccs c WHERE a.recid=b.recid AND b.sn_pol=c.sn_pol AND ;
 	et=m.TipOfExp GROUP BY docexp,reason INTO CURSOR curexps
 SELECT a.usr,a.docexp,a.reason as reason FROM merror a, talon b, curokaccs c WHERE ;
 	a.usr=goApp.smoexp AND a.recid=b.recid AND b.sn_pol=c.sn_pol AND et=m.TipOfExp ;
 	GROUP BY usr,docexp,reason INTO CURSOR curexps
 
 m.t_reason = goApp.reason
 SELECT curexps
 SCAN 
  IF INLIST(m.TipOfExp,'6')
   m.theme   = curexps.reason
   m.povod    = IIF(SEEK(m.theme, 'themes'), themes.name, '')
   goApp.reason = m.theme
  ELSE 
   m.reason   = curexps.reason
   m.povod    = IIF(SEEK(m.reason, 'reasons'), reasons.name, '')
   goApp.reason = m.reason
  ENDIF 
  
  
  m.docexp = docexp
  m.docfio = IIF(SEEK(m.docexp, 'explist'), ;
   ALLTRIM(explist.fam)+' '+ALLTRIM(explist.im)+' '+ALLTRIM(explist.ot)+', ÍÓ‰ '+m.docexp, '')
  m.usr = usr

  IF m.usr<>goApp.smoexp AND goApp.smoexp<>m.SuperExp
   LOOP 
  ENDIF 
  =OneSvAct(m.docexp)

 ENDSCAN 
 goApp.reason = m.t_reason

 USE IN curexps
 USE IN curokaccs
 IF USED('people')
  USE IN people
 ENDIF 
 IF USED('talon') 
  USE IN talon 
 ENDIF 
 IF USED('merror') 
  USE IN merror
 ENDIF 

 SELECT aisoms

RETURN 

FUNCTION OneSvAct(paraexp)
 PRIVATE m.docexp
 
 m.docexp = m.paraexp

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

 SELECT recid, resume, conclusion, recommend FROM svacts WHERE period=m.gcperiod AND mcod=m.mcod AND codexp=INT(VAL(m.TipOfExp)) ;
  AND docexp=m.docexp AND flcod=m.flcod AND reason=goApp.reason INTO CURSOR rqwest NOCONSOLE  
 m.nfileid = recid
 m.resume     = resume
 m.conclusion = conclusion
 m.recommend  = recommend
 USE IN rqwest

 IF m.nfileid>0
  DocName = IIF(!IsUsrDir, m.pmee, m.usrdir)+'\svacts\'+PADL(m.nfileid,6,'0')
  m.n_akt = NumActOfExp(m.lpuid,m.tipofexp,goApp.reason,m.nfileid)
 ELSE 
   SELECT TOP 1 resume, conclusion, recommend FROM svacts ORDER BY recid DESC INTO CURSOR cqwert
   m.resume     = resume
   m.conclusion = conclusion
   m.recommend  = recommend
   USE IN cqwert

   *INSERT INTO svacts (n_rss,period,mcod,lpu_id,codexp,docexp,flcod,e_period,smoexp,et,qr) ;
    VALUES ;
   (m.n_rss,m.gcperiod,m.mcod,m.lpuid,INT(VAL(m.tipofexp)),m.docexp,m.flcod,m.e_period,goApp.smoexp,m.tipofexp,.t.)
   INSERT INTO svacts (n_rss,period,mcod,lpu_id,codexp,docexp,flcod,smoexp,et,qr,reason) ;
    VALUES ;
   (m.n_rss,m.gcperiod,m.mcod,m.lpuid,INT(VAL(m.tipofexp)),m.docexp,m.flcod,goApp.smoexp,m.tipofexp,.t.,goApp.reason)

   m.nfileid = GETAUTOINCVALUE()
   m.n_akt = NumActOfExp(m.lpuid,m.tipofexp,goApp.reason,m.nfileid)

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

   IF USED('rrfile')
    USE IN rrfile
   ENDIF 
   IF USED('moves')
    USE IN moves
   ENDIF 
   SELECT curexps
   RETURN
  ELSE 
   IF m.TipOfPeriod=0
    UPDATE svacts SET actdate=DATETIME() WHERE recid = m.nfileid
    IF USED('moves')
     UPDATE moves SET dat=DATETIME() WHERE actid = m.nfileid
    ENDIF 
   ENDIF 
  ENDIF 
 ELSE 
  IF USED('moves')
   INSERT INTO moves (actid,et,usr,dat) VALUES (m.nfileid,'1',m.gcUser,DATETIME())
  ENDIF 
 ENDIF 
  
 DO FORM TxtForActs

 UPDATE svacts SET resume=ALLTRIM(m.resume), conclusion=ALLTRIM(m.conclusion), recommend=ALLTRIM(m.recommend) WHERE recid = m.nfileid

 m.checked_tot = 0
 m.checked_amb = 0
 m.checked_dst = 0
 m.checked_st  = 0
 m.checked_02  = 0

 *m.def_amb = 0
 *m.def_dst = 0
 *m.def_st  = 0

 *m.nepredst    = 0
 m.checked     = 0
 *m.totdefs     = 0
 *m.sumneoplata = 0
 *m.tot_straf   = 0
 *m.kol_straf   = 0
 m.tot_s_all   = 0
  
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
   m.podvid='“'
  CASE m.TipOfExp = '7'
   m.podvid='“'
  OTHERWISE 
   m.podvid='0'
 ENDCASE 

 m.d_akt = IIF(m.qcod!='I3', DTOC(DATE()), '')
 
 m.nakt = 'π ' + m.n_akt + ' ÓÚ ' + m.d_akt
 m.cpredps = 'œÂ‰ÔËÒ‡ÌËÂ π '+m.n_akt+' ÓÚ '+m.d_akt

 m.exp_dat1 = '01.'+PADL(tMonth,2,'0')+'.'+STR(tYear,4)
 m.exp_dat2 = DTOC(GOMONTH(CTOD(m.exp_dat1),1)-1)

 CREATE CURSOR curpaz (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 
 CREATE CURSOR qwertamb (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR qwertst (c_i c(30))
 INDEX on c_i TAG c_i
 SET ORDER TO c_i

 CREATE CURSOR qwertdst (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR qwert02 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR concls (nrec n(5), concl c(250))
 CREATE CURSOR rcmnds (nrec n(5), rcmnd c(250))

 CREATE CURSOR workcurs (nrec n(5), sn_pol c(25), c_i c(30), ds c(6), d_u d, cod n(6), d_u0 d, s_all n(11,2), ;
 	d_beg d, d_end d, er_c c(2), osn230 c(6), koeff n(4,2), straf n(4,2), cmnt c(50))
 INDEX ON sn_pol TAG sn_pol
 INDEX ON c_i TAG c_i
 
 *CREATE CURSOR curdata2 (nrec2 n(5), sn_pol c(25), c_i c(30), d_beg d, d_end d, ds c(6), cod n(6), s_all n(11,2), ;
 * osn230 c(5), er_c c(3), delta n(11,2), koeff n(3,2), straf n(11,2), cmnt c(50))
 *SELECT curdata2
 *INDEX on sn_pol TAG sn_pol
 *SET ORDER TO sn_pol

 *CREATE CURSOR curdata3 (nrec3 n(5), sn_pol c(25), c_i c(30), d_beg d, d_end d, ds c(6), cod n(6), s_all n(11,2), ;
 * osn230 c(5), er_c c(3), delta n(11,2), koeff n(3,2), straf n(11,2), cmnt c(50))
 *SELECT curdata3
 *INDEX on sn_pol TAG sn_pol
 *SET ORDER TO sn_pol

 IF USED('rrfile')
  UPDATE rrfile SET n_akt='', d_akt={}, t_akt='SV' WHERE n_akt=m.n_akt
 ENDIF 

 SELECT merror
 SET RELATION TO recid INTO talon 
 *m.nrec2 = 0 
 *m.nrec3 = 0 
 SCAN 

  IF !(et=m.TipOfExp AND docexp=m.docexp)
   LOOP 
  ENDIF 
  IF LEFT(err_mee,2)<>'W0'
   LOOP 
  ENDIF 
  m.sn_pol      = talon.sn_pol
  IF !SEEK(m.sn_pol, 'curokaccs')
   LOOP 
  ENDIF 
  IF reason<>goApp.reason
   LOOP 
  ENDIF 
  
  REPLACE n_akt WITH m.n_akt, d_akt WITH goApp.d_exp, t_akt WITH 'SV'
  
  IF USED('rrfile')
   IF SEEK(m.sn_pol, 'rrfile')
    REPLACE n_akt WITH m.n_akt, d_akt WITH goApp.d_exp, t_akt WITH 'SV' IN rrfile
   ENDIF 
  ENDIF 

  m.s_1         = s_1
  m.s_2         = s_2
  m.koeff       = koeff
  m.straf       = 0
  m.c_i         = talon.c_i
  m.er_c        = UPPER(LEFT(err_mee,2))
  m.osn230      = osn230
  m.s_all       = s_all
  m.cod         = cod
  m.ds          = talon.ds
  m.tot_s_all   = m.tot_s_all + m.s_all
  m.d_beg       = IIF(SEEK(m.sn_pol, 'people', 'sn_pol'), people.d_beg, {})
  m.d_end       = IIF(SEEK(m.sn_pol, 'people', 'sn_pol'), people.d_end, {})
  m.d_u         = talon.d_u
  m.d_u0        = IIF(!IsMes(m.cod) and !IsVMP(m.cod), talon.d_u, talon.d_u-talon.k_u+1)

  m.cmnt = PADL(m.cod,6,'0')

  INSERT INTO workcurs (sn_pol, c_i, s_all, ds, d_u, d_u0, d_beg, d_end, er_c, osn230, koeff, straf, cod, cmnt) VALUES ;
   (m.sn_pol, m.c_i, m.s_all, m.ds, m.d_u, m.d_u0, m.d_beg, m.d_end, m.er_c, m.osn230, m.koeff, m.straf, m.cod, m.cmnt)
 
  IF !m.Is02
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
   ELSE 
    *IF (Is02(m.cod)) AND !SEEK(m.sn_pol, 'qwert02')
    * INSERT INTO qwert02 (sn_pol) VALUES (m.sn_pol)
    * m.checked_02 = m.checked_02 + 1
    *ENDIF 
    m.checked_02 = m.checked_02 + 1
   ENDIF 

 ENDSCAN 
 SET RELATION OFF INTO talon 
 
 *m.defs = ''
 SELECT workcurs
 SET ORDER TO sn_pol
 GO TOP 
 m.nrec = IIF(m.mcod='0371001', 0, 1)
 m.polis = sn_pol
 SCAN
  *m.sumneoplata = m.sumneoplata + s_all
  m.sn_pol = sn_pol
  *m.defs = m.defs+osn230+'; '
  IF m.sn_pol!=m.polis OR m.mcod='0371001'
   m.polis = m.sn_pol
   m.nrec = m.nrec + 1
  ENDIF 
  REPLACE nrec WITH m.nrec
 ENDSCAN 
 
 *SELECT curdata2
 *IF RECCOUNT('curdata2')<=0
 * SCATTER MEMVAR 
 * INSERT INTO curdata2 FROM MEMVAR 
 *ENDIF 
 *SELECT curdata3
 *IF RECCOUNT('curdata3')<=0
 * SCATTER MEMVAR 
 * INSERT INTO curdata3 FROM MEMVAR 
 *ENDIF 
* SET ORDER TO 
* REPLACE ALL nrec2 WITH RECNO()

 m.checked = m.checked_amb + m.checked_dst + m.checked_st + m.checked_02
 *m.totdefs = IIF(INLIST(m.qcod,'P2','I3'), m.def_amb+m.def_dst+m.def_st, m.totdefs)
 *m.totdefs = m.def_amb+m.def_dst+m.def_st
 
 IF !m.Is02
  UPDATE svacts SET n_ss=m.checked_tot, n_st=m.checked_st, n_dst=m.checked_dst, n_plk=m.checked_amb;
 	WHERE recid = m.nfileid
 ELSE 
  UPDATE svacts SET n_ss=m.checked_tot, n_02=m.checked_02;
 	WHERE recid = m.nfileid
 ENDIF 
 LOCAL m.lcTmpName, m.lcRepName, m.lcDbfName, m.llResult
 m.lcTmpName = pTempl+'\'+m.dotname
 m.lcRepName = m.docname+'.xls'

 m.conclusion   = ALLTRIM(m.conclusion)
 m.str_len = 90
 IF LEN(m.conclusion)>0 && ÂÒÎË ‚ÓÓ·˘Â ˜ÚÓ-ÚÓ ÂÒÚ¸
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
  ELSE && ÂÒÎË ÒÚÓÍ‡ <90 ÒËÏ‚ÓÎÓ‚
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

 m.recommend   = ALLTRIM(m.recommend)
 m.str_len = 90
 IF LEN(m.recommend)>0 && ÂÒÎË ‚ÓÓ·˘Â ˜ÚÓ-ÚÓ ÂÒÚ¸
  IF LEN(m.recommend)>m.str_len
   m.s_str = ''    
   FOR m.n_ch=1 TO LEN(m.recommend)
    m.s_str = m.s_str+ SUBSTR(m.recommend,m.n_ch,1)
    DO CASE 
     CASE LEN(m.s_str)>=m.str_len
      IF RIGHT(m.s_str,1)=SPACE(1)
       INSERT INTO rcmnds (rcmnd) VALUES (m.s_str)
       m.s_str = ''    
      ELSE 
       IF RAT(SPACE(1),m.s_str)>0
        INSERT INTO rcmnds (rcmnd) VALUES (LEFT(m.s_str,RAT(SPACE(1),m.s_str)-1))
        m.s_str = SUBSTR(m.s_str,RAT(SPACE(1),m.s_str)+1)
       ELSE 
        INSERT INTO rcmnds (rcmnd) VALUES (m.s_str)
        m.s_str = ''    
       ENDIF 
      ENDIF 
     CASE SUBSTR(m.recommend,m.n_ch,2) = CHR(13)+CHR(10)
      m.n_ch = m.n_ch + 1
      INSERT INTO rcmnds (rcmnd) VALUES (LEFT(m.s_str,LEN(m.s_str)-1))
      m.s_str = ''    
     OTHERWISE 
    ENDCASE 
   ENDFOR 
   INSERT INTO rcmnds (rcmnd) VALUES (m.s_str)
  ELSE && ÂÒÎË ÒÚÓÍ‡ <90 ÒËÏ‚ÓÎÓ‚
   *INSERT INTO rcmnds (rcmnd) VALUES (m.recommend)
   m.s_str = ''    
   FOR m.n_ch=1 TO LEN(m.recommend)
    m.s_str = m.s_str+ SUBSTR(m.recommend,m.n_ch,1)
    DO CASE 
     CASE LEN(m.s_str)>=m.str_len
      IF RIGHT(m.s_str,1)=SPACE(1)
       INSERT INTO rcmnds (rcmnd) VALUES (m.s_str)
       m.s_str = ''    
      ELSE 
       IF RAT(SPACE(1),m.s_str)>0
        INSERT INTO rcmnds (rcmnd) VALUES (LEFT(m.s_str,RAT(SPACE(1),m.s_str)-1))
        m.s_str = SUBSTR(m.s_str,RAT(SPACE(1),m.s_str)+1)
       ELSE 
        INSERT INTO rcmnds (rcmnd) VALUES (m.s_str)
        m.s_str = ''    
       ENDIF 
      ENDIF 
     CASE SUBSTR(m.recommend,m.n_ch,2) = CHR(13)+CHR(10)
      m.n_ch = m.n_ch + 1
      INSERT INTO rcmnds (rcmnd) VALUES (LEFT(m.s_str,LEN(m.s_str)-1))
      m.s_str = ''    
     OTHERWISE 
    ENDCASE 
   ENDFOR 
   INSERT INTO rcmnds (rcmnd) VALUES (m.s_str)
  ENDIF 
 ELSE 
  INSERT INTO rcmnds (rcmnd) VALUES (SPACE(10))
 ENDIF 

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
 
*  m.barcode = '1 '+SUBSTR(m.n_akt,3) + ' '+ALLTRIM(STR(m.checked))
*  loFbc.BarcodeImage(m.barcode, pMee+'\svacts\'+PADL(m.nfileid,6,'0')+'.jpg')

*  m.acthead = '1'+PADR(STRTRAN(m.n_akt,'“','T'),16)+PADR(goApp.smoexp,7)+PADR(m.docexp,7)+STRTRAN(m.reason,'0',' ')
  m.acthead = '1'+PADR(STRTRAN(m.n_akt,'“','T'),16)+PADR(goApp.smoexp,7)+PADR(m.docexp,7)+IIF(m.tipofexp='5', m.reason, ' ')
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

  *WITH oSheet.Range("I2:J6")
  WITH oSheet.Range("AB1:AB5")
   oPic.Top = .Top
   oPic.Left = .Left
   oPic.Height = .Height
  ENDWITH 
  IF 1=2
  * —Í˚‚‡ÂÏ –‡Á‰ÂÎ 3  
  oRange = oExcel.Rows("26:40")
  oRange.EntireRow.Hidden = .t.
  osheet.Cells(25,1).Value="III. «‡ÔÓÎÌˇÂÚÒˇ ÔË Ì‡ÎË˜ËË ‰‡ÌÌ˚ı"
  * —Í˚‚‡ÂÏ –‡Á‰ÂÎ 3  

  * —Í˚‚‡ÂÏ –‡Á‰ÂÎ 6
  oRange = oExcel.Rows("72:97")
  oRange.EntireRow.Hidden = .t.
  osheet.Cells(71,1).Value="VI. «‡ÔÓÎÌˇÂÚÒˇ ÔË Ì‡ÎË˜ËË ‰‡ÌÌ˚ı"
  * —Í˚‚‡ÂÏ –‡Á‰ÂÎ 6  
  ENDIF 

  IF m.is02
   osheet.Range("ambstdst").EntireRow.Hidden = .T.
   osheet.Range("checked02").EntireRow.Hidden = .F.
  ELSE 
   osheet.Range("ambstdst").EntireRow.Hidden = .F.
   osheet.Range("checked02").EntireRow.Hidden = .T.
  ENDIF 

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

 USE IN concls
 USE IN rcmnds
 USE IN workcurs
 USE IN qwertamb
 SELECT qwertst
 USE IN qwertdst
 USE IN qwert02
 *USE IN badamb
 *SELECT badst
 *USE IN baddst
 IF USED('rrfile')
  USE IN rrfile
 ENDIF 
 IF USED('moves')
  USE IN moves
 ENDIF 

 SELECT curexps

RETURN 