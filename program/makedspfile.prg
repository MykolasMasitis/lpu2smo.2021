PROCEDURE MakeDspFile(para1, para2)
 m.NeedOpen = para1
 m.IsSilent = para2
 IF PARAMETERS()>0
  m.NeedOpen = para1
 ENDIF 
 IF PARAMETERS()>1
  m.IsSilent = para2
 ENDIF 
 
 m.q = m.qcod
 
 IF !m.IsSilent
  IF MESSAGEBOX(CHR(13)+CHR(10)+'ÂÛ ÕÎÒÈÒÅ ÑÔÎÐÌÈÐÎÂÀÒÜ ÔÀÉË DSP-ÔÀÉË?'+CHR(13)+CHR(10),4+32,'')=7
   RETURN 
  ENDIF 
 ENDIF 
 
 IF !fso.FileExists(pcommon+'\dspcodes.dbf')
  IF !m.IsSilent
   MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË DSPCODES.DBF!'+CHR(13)+CHR(10),0+16,'')
  ENDIF 
  RETURN 
 ENDIF 

 m.prperiod = STR(IIF(tmonth=1, tyear-1, tyear),4) + PADL(IIF(tmonth=1, 12, tmonth-1),2,'0')
* IF tdat1>{01.01.2014}
  IF !m.IsSilent
   IF !fso.FileExists(pbase+'\'+m.prperiod+'\dsp.dbf')
    IF MESSAGEBOX(CHR(13)+CHR(10)+'ÍÅ ÑÔÎÐÌÈÐÎÂÀÍ DSP-ÔÀÉË ÇÀ ÏÐÅÄÛÄÓÙÈÉ ÏÅÐÈÎÄ!'+CHR(13)+CHR(10)+;
    	'ÑÔÎÐÌÈÐÎÂÀÒÜ ÏÓÑÒÎÉ ÔÀÉË Â ÒÅÊÓÙÅÌ ÏÅÐÈÎÄÅ?',0+16,'') = 7
     RETURN
    ELSE
     CREATE TABLE &pBase\&gcPeriod\dsp (recid i, q c(2), period c(6), mcod c(7), sn_pol c(17), c_i c(30), ds c(6),;
     	fam c(20), im c(20), ot c(20), w n(1), dr d, ages n(2), cod n(6), rslt n(3), ;
     	d_u d, s_all n(11,2), k_u2 n(3), s_all2 n(11,2), k_u2ok n(3), s_all2ok n(11,2), er c(3), tip n(1))
  	 INDEX ON period+mcod+PADL(recid,6,'0') TAG uniqq && èñïîëüçóåòñÿ â formdds
  	 *INDEX on mcod+sn_pol+PADL(cod,6,"0") TAG exptag 
     INDEX on sn_pol+PADL(tip,1,'0') TAG exptag
     INDEX on sn_pol+PADL(cod,6,'0') TAG un_tag
  	 INDEX on c_i TAG c_i 
  	 USE 
    ENDIF
   ELSE 
    IF tmonth>1
     fso.CopyFile(pbase+'\'+m.prperiod+'\dsp.dbf', pbase+'\'+m.gcperiod+'\dsp.dbf')
     fso.CopyFile(pbase+'\'+m.prperiod+'\dsp.cdx', pbase+'\'+m.gcperiod+'\dsp.cdx')
    ELSE 
     CREATE TABLE &pBase\&gcPeriod\dsp (recid i, q c(2), period c(6), mcod c(7), sn_pol c(17), c_i c(30), ds c(6),;
     	fam c(20), im c(20), ot c(20), w n(1), dr d, ages n(2), cod n(6), rslt n(3), ;
     	d_u d, s_all n(11,2), k_u2 n(3), s_all2 n(11,2), k_u2ok n(3), s_all2ok n(11,2), er c(3), tip n(1))
  	 INDEX ON period+mcod+PADL(recid,6,'0') TAG uniqq && èñïîëüçóåòñÿ â formdds
  	 *INDEX on mcod+sn_pol+PADL(cod,6,"0") TAG exptag 
     INDEX on sn_pol+PADL(tip,1,'0') TAG exptag
     INDEX on sn_pol+PADL(cod,6,'0') TAG un_tag
  	 INDEX on c_i TAG c_i 
  	 USE 
    ENDIF 
   ENDIF 
  ELSE && IF !m.IsSilent
   IF !fso.FileExists(pbase+'\'+m.prperiod+'\dsp.dbf')
    CREATE TABLE &pBase\&gcPeriod\dsp (recid i, period c(6), q c(2), mcod c(7), sn_pol c(17), c_i c(30), ds c(6), ;
     	fam c(20), im c(20), ot c(20), w n(1), dr d, ages n(2), cod n(6), rslt n(3), ;
     	d_u d, s_all n(11,2), k_u2 n(3), s_all2 n(11,2), k_u2ok n(3), s_all2ok n(11,2), er c(3), tip n(1))
  	INDEX ON period+mcod+PADL(recid,6,'0') TAG uniqq
  	*INDEX on mcod+sn_pol+PADL(cod,6,"0") TAG exptag 
    INDEX on sn_pol+PADL(tip,1,'0') TAG exptag
    INDEX on sn_pol+PADL(cod,6,'0') TAG un_tag
  	INDEX on c_i TAG c_i 
  	USE 
   ELSE 
    fso.CopyFile(pbase+'\'+m.prperiod+'\dsp.dbf', pbase+'\'+m.gcperiod+'\dsp.dbf')
    fso.CopyFile(pbase+'\'+m.prperiod+'\dsp.cdx', pbase+'\'+m.gcperiod+'\dsp.cdx')
   ENDIF 
  ENDIF && IF !m.IsSilent
* ENDIF 

 m.lcpath = pbase+'\'+m.gcperiod
 m.period = m.gcperiod

 IF OpenFile(pcommon+'\dspcodes', 'dspcodes', 'shar', 'cod')>0
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\dsp', 'dsp', 'excl')>0
  IF USED('dsp')
   USE IN dsp 
  ENDIF 
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  RETURN 
 ELSE 
  SELECT dsp 
  DELETE TAG ALL 
  INDEX ON period+mcod+PADL(recid,6,'0') TAG uniqq
  INDEX on sn_pol+PADL(tip,1,'0') TAG exptag
  INDEX on sn_pol+PADL(cod,6,'0') TAG un_tag
  INDEX on c_i TAG c_i 
  *INDEX ON period+mcod+PADL(recid,6,'0') TAG uniqq
  *INDEX on mcod+sn_pol+PADL(cod,6,'0') TAG exptag 
  *INDEX on c_i TAG c_i 
  USE 
 ENDIF 

 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\dsp', 'dsp', 'shar', 'uniqq')>0
  IF USED('dsp')
   USE IN dsp 
  ENDIF 
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT dsp 
 SET RELATION TO cod INTO dspcodes 
 REPLACE tip WITH dspcodes.tip ALL
 SET RELATION OFF INTO  dspcodes 
 
 IF !fso.FileExists(m.pbase+'\'+m.gcperiod+'\aisoms.dbf')
  IF USED('dsp')
   USE IN dsp 
  ENDIF 
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  RETURN 
 ENDIF 
 
 IF m.NeedOpen
 IF OpenFile(m.pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('dsp')
   USE IN dsp 
  ENDIF 
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  RETURN 
 ENDIF 
 ENDIF 
 
 SELECT aisoms
 SCAN 
  m.mcod = mcod
  IF !fso.FolderExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF RECCOUNT('talon')<=0
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'errsv', 'shar', 'rid')>0
   IF USED('errsv')
    USE IN errsv
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
 
  WAIT "ÎÁÐÀÁÎÒÊÀ..." WINDOW NOWAIT 
  SELECT talon
  SET RELATION TO sn_pol INTO people
  SET RELATION TO recid INTO errsv ADDITIVE 
  m.nusl = 0
  SCAN 
   m.cod = cod
   IF !SEEK(m.cod, 'dspcodes')
    LOOP
   ENDIF 
  
   m.tipofcod = dspcodes.tip
   m.rslt     = rslt
  
   DO CASE 
    CASE m.tipofcod = 1 && Äèñïàñåðèçàöèÿ âçðîñëûõ, ñ èþëÿ
     IF !(INLIST(m.rslt,317,318,353,357,358) OR BETWEEN(m.rslt,355,356)) && !BETWEEN(m.rslt,316,319) AND !BETWEEN(m.rslt,352,358) 
      LOOP 
     ENDIF 
    CASE m.tipofcod = 2 && Ïðîôîñìîòðû âçðîñëûõ, ñ ñåíòÿáðÿ
     *IF !BETWEEN(m.rslt,343,345) AND !INLIST(m.rslt,373,374)
     IF !BETWEEN(m.rslt,343,344) AND !INLIST(m.rslt,373,374)
      LOOP 
     ENDIF 
    CASE m.tipofcod = 3 && Äèñïàñåðèçàöèÿ äåòåé-ñèðîò, ñ èþëÿ
     IF !(BETWEEN(m.rslt,347,351) OR BETWEEN(m.rslt,369,372)) AND ;
     	!(BETWEEN(m.rslt,321,325) OR BETWEEN(m.rslt,365,368)) && !BETWEEN(m.rslt,321,325) AND m.rslt!=320 AND m.rslt!=390 AND !BETWEEN(m.rslt,347,351) && Ýòî - óñûíîâëåííûå ñèðîòû!
    * IF !BETWEEN(m.rslt,321,325) AND m.rslt!=320 AND m.rslt!=390 AND !BETWEEN(m.rslt,347,351) && Ýòî - óñûíîâëåííûå ñèðîòû!
      LOOP 
     ENDIF 
    CASE m.tipofcod = 4 && Ïðîôîñìîòðû íåñîâåðøåííîëåòíèõ, ñ ñåíòÿáðÿ
     *IF !BETWEEN(m.rslt,332,336) AND m.rslt!=326
     IF !(BETWEEN(m.rslt,332,336) OR BETWEEN(m.rslt,361,364)) && !BETWEEN(m.rslt,332,336) AND m.rslt!=326
      LOOP 
     ENDIF 
    CASE m.tipofcod = 5 && Ïðåäâàðèòåëüíûå ïðîôîñìîòðû íåñîâåðøåííîëåòíèõ, ñ ñåíòÿáðÿ
     IF !BETWEEN(m.rslt,337,341) AND !INLIST(m.rslt,326,396)
      LOOP 
     ENDIF 
    CASE m.tipofcod = 6 && Ïåðèîäè÷åñêèå ïðîôîñìîòðû íåñîâåðøåííîëåòíèõ, ñ ñåíòÿáðÿ
     IF m.rslt!=342  AND m.rslt!=326
      LOOP 
     ENDIF 
    CASE m.tipofcod = 7 && ïðîôîñìîòðû, ñ ñåíòÿáðÿ
    OTHERWISE 
     LOOP 
   ENDCASE 

   m.nusl = m.nusl + 1

   m.mcod  = mcod
   m.recid = recid
   m.key   = m.period + m.mcod + PADL(m.recid,6,'0')
   IF SEEK(m.key, 'dsp')
    LOOP 
   ENDIF 

   m.d_u   = d_u
   m.k_u   = k_u
   m.s_all = s_all
   m.er    = ''
   m.ds    = ds

   m.id_smo = recid
   m.sn_pol = sn_pol
   m.fam    = people.fam
   m.im     = people.im
   m.ot     = people.ot
   m.w      = people.w
   m.dr     = people.dr
   m.er     = errsv.c_err
   
   m.ages   = YEAR(tdat1) - YEAR(m.dr)
  
   m.c_i    = c_i 
   
   m.tip = m.tipofcod

   INSERT INTO dsp FROM MEMVAR 
   
  ENDSCAN 
  SET RELATION OFF INTO people
  SET RELATION OFF INTO errsv

  USE IN errsv
  USE IN talon
  USE IN people

  SELECT aisoms
  WAIT CLEAR 
 ENDSCAN && aisoms
 
 IF !m.IsSilent
 MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÍÀÐÓÆÅÍÎ '+TRANSFORM(m.nusl, '99999')+' ÓÑËÓÃ!',0+64,'')
 ENDIF 
 
 WAIT "ÐÀÑ×ÅÒ ÂÒÎÐÎÃÎ ÝÒÀÏÀ (ÍÎÂÛÉ ÀËÃÎÐÈÒÌ)..." WINDOW NOWAIT 
 SELECT dsp
 SET ORDER TO c_i
 
 SELECT aisoms
 SCAN 
  m.mcod = mcod
  IF !fso.FolderExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF RECCOUNT('talon')<=0
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(m.pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'errsv', 'shar', 'rid')>0
   IF USED('errsv')
    USE IN errsv
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
 
  SELECT talon 
  SET ORDER TO c_i 
  SET RELATION TO LEFT(c_i,25) INTO dsp 
  SET RELATION TO recid INTO errsv ADDITIVE 

  m.oc_i     = c_i 
  m.k_u2     = 0
  m.s_all2   = 0
  m.k_u2ok   = 0
  m.s_all2ok = 0

  SCAN 
   IF EMPTY(dsp.c_i)
    LOOP 
   ENDIF 
   m.rslt = dsp.rslt 
   IF !INLIST(m.rslt,320,326,352,353,357,358,361,362,363,364,369,370,371,372,390,396)
    LOOP 
   ENDIF 
   IF dsp.k_u2>0
    *LOOP && Ïî÷åìó?
   ENDIF 	

   IF d_u < dsp.d_u 
    LOOP 
   ENDIF 
   IF cod = dsp.cod 
    LOOP 
   ENDIF 
  
   IF c_i != m.oc_i
    IF m.k_u2>0
     REPLACE k_u2 WITH m.k_u2, s_all2 WITH m.s_all2, k_u2ok WITH m.k_u2ok, s_all2ok WITH m.s_all2ok IN dsp 
    ENDIF 
   
    m.oc_i = c_i 
   
    m.k_u2     = 0
    m.s_all2   = 0
    m.k_u2ok   = 0
    m.s_all2ok = 0
   ENDIF 
  
   m.k_u2   = m.k_u2 + k_u
   m.s_all2 = m.s_all2 + s_all
    
   IF EMPTY(errsv.rid)
    m.k_u2ok   = m.k_u2ok + k_u
    m.s_all2ok = m.s_all2ok + s_all
   ENDIF 

  ENDSCAN 
  SET RELATION OFF INTO dsp 
  SET RELATION OFF INTO errsv
  WAIT CLEAR 

  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('errsv')
   USE IN errsv
  ENDIF 

 ENDSCAN 

 IF USED('dsp')
  USE IN dsp
 ENDIF 
 IF USED('dspcodes')
  USE IN dspcodes
 ENDIF 
 
 IF m.NeedOpen
  IF USED('aisoms')
   USE IN aisoms 
  ENDIF 
 ENDIF 
 
 IF !m.IsSilent 
  MESSAGEBOX('OK!',0+64,'')
 ENDIF 

RETURN 
