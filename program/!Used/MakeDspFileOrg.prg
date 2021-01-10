PROCEDURE MakeDspFileOrg(para1, para2)
 m.NeedOpen = .t.
 m.IsSilent = .f.
 IF PARAMETERS()>0
  m.NeedOpen = para1
 ENDIF 
 IF PARAMETERS()>1
  m.IsSilent = para2
 ENDIF 
 
* IF tdat1<{01.01.2014}
*  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒ×ÅÒ ÌÎÆÅÒ ÔÎÐÌÈÐÎÂÀÒÜÑß ÍÀ×ÈÍÀß Ñ ßÍÂÀÐß 2014!'+CHR(13)+CHR(10),0+16,'')
*  RETURN 
* ENDIF 
 IF !m.IsSilent
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÂÛ ÕÎÒÈÒÅ ÑÔÎÐÌÈÐÎÂÀÒÜ ÔÀÉË DSP-ÔÀÉË?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 ENDIF 
 
 m.prperiod = STR(IIF(tmonth=1, tyear-1, tyear),4) + PADL(IIF(tmonth=1, 12, tmonth-1),2,'0')
 IF tdat1>{01.01.2014}
  IF !m.IsSilent
   IF !fso.FileExists(pbase+'\'+m.prperiod+'\dsp.dbf')
    MESSAGEBOX(CHR(13)+CHR(10)+'ÍÅ ÑÔÎÐÌÈÐÎÂÀÍ DSP-ÔÀÉË ÇÀ ÏÐÅÄÛÄÓÙÈÉ ÏÅÐÈÎÄ!'+CHR(13)+CHR(10),0+16,'')
    *RETURN 
   ENDIF 
  ENDIF 
 ENDIF 
 
 lcpath = pbase+'\'+m.gcperiod
 m.period = m.gcperiod
 
 IF !fso.FileExists(lcpath+'\talon.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÍÅ ÑÔÎÐÌÈÐÎÂÀÍ ÑÂÎÄÍÛÉ Ñ×ÅÒ ÇÀ ÏÅÐÈÎÄ!'+CHR(13)+CHR(10),0+16,'talon.dbf')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(lcpath+'\people.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÍÅ ÑÔÎÐÌÈÐÎÂÀÍ ÑÂÎÄÍÛÉ Ñ×ÅÒ ÇÀ ÏÅÐÈÎÄ!'+CHR(13)+CHR(10),0+16,'people.dbf')
  RETURN 
 ENDIF 

 IF !fso.FileExists(lcpath+'\e'+m.gcperiod+'.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÍÅ ÑÔÎÐÌÈÐÎÂÀÍ ÑÂÎÄÍÛÉ Ñ×ÅÒ ÇÀ ÏÅÐÈÎÄ!'+CHR(13)+CHR(10),0+16,'\e'+m.gcperiod+'.dbf')
  RETURN 
 ENDIF 

 IF !fso.FileExists(pcommon+'\dspcodes.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË DSPCODES.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 IF OpenFile(lcpath+'\talon', 'talon', 'shar')>0
  IF USED('talon')
   USE IN talon
  ENDIF 
  RETURN 
 ENDIF 
 
 IF OpenFile(lcpath+'\people', 'people', 'shar', 'sn_pol')>0
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(lcpath+'\e'+m.gcperiod, 'errsv', 'shar', 'rid')>0
  IF USED('errsv')
   USE IN errsv
  ENDIF 
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  RETURN 
 ENDIF 

 IF RECCOUNT('talon')<=0
  IF USED('talon')
   USE IN talon
  ENDIF 
  MESSAGEBOX(CHR(13)+CHR(10)+'ÑÂÎÄÍÛÉ ÔÀÉË TALON.DBF ÏÓÑÒ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 IF RECCOUNT('people')<=0
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  MESSAGEBOX(CHR(13)+CHR(10)+'ÑÂÎÄÍÛÉ ÔÀÉË PEOPLE.DBF ÏÓÑÒ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 IF OpenFile(pcommon+'\dspcodes', 'dspcodes', 'shar', 'cod')>0
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  RETURN 
 ENDIF 
 
 dspfile = 'dsp'
 IF fso.FileExists(lcpath+'\'+dspfile+'.dbf')
  fso.DeleteFile(lcpath+'\'+dspfile+'.dbf')
 ENDIF  
 IF !fso.FileExists(lcpath+'\'+dspfile+'.dbf')
  IF INLIST(m.gcperiod,'201401','201501','201601')
  CREATE TABLE &lcpath\&dspfile (recid i, period c(6), mcod c(7), sn_pol c(17), c_i c(30), ;
   fam c(20), im c(20), ot c(20), w n(1), dr d, ages n(2), cod n(6), rslt n(3), ;
    d_u d, s_all n(11,2), k_u2 n(3), s_all2 n(11,2), k_u2ok n(3), s_all2ok n(11,2), er c(3))
  INDEX ON period+mcod+PADL(recid,6,'0') TAG uniqq
  INDEX on mcod+sn_pol+PADL(cod,6,"0") TAG exptag 
  INDEX on c_i TAG c_i 
  USE 
  ELSE 
   fso.CopyFile(pbase+'\'+m.prperiod+'\dsp.dbf', pbase+'\'+m.gcperiod+'\dsp.dbf')
   fso.CopyFile(pbase+'\'+m.prperiod+'\dsp.cdx', pbase+'\'+m.gcperiod+'\dsp.cdx')
  ENDIF
 ENDIF 

 IF OpenFile(lcpath+'\'+dspfile, 'dsp', 'excl')>0
  IF USED('dsp')
   USE IN dsp 
  ENDIF 
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  RETURN 
 ELSE 
  SELECT dsp 
  DELETE TAG ALL 
  INDEX ON period+mcod+PADL(recid,6,'0') TAG uniqq
  INDEX on mcod+sn_pol+PADL(cod,6,"0") TAG exptag 
  INDEX on c_i TAG c_i 
  USE 
 ENDIF 

 IF OpenFile(lcpath+'\'+dspfile, 'dsp', 'shar', 'uniqq')>0
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  RETURN 
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
    IF !BETWEEN(m.rslt,316,319) AND !BETWEEN(m.rslt,352,358)
     LOOP 
    ENDIF 
   CASE m.tipofcod = 2 && Ïðîôîñìîòðû âçðîñëûõ, ñ ñåíòÿáðÿ
    IF !BETWEEN(m.rslt,343,345)
     LOOP 
    ENDIF 
   CASE m.tipofcod = 3 && Äèñïàñåðèçàöèÿ äåòåé-ñèðîò, ñ èþëÿ
    IF !BETWEEN(m.rslt,321,325) AND m.rslt!=320 AND m.rslt!=390 AND !BETWEEN(m.rslt,347,351) && Ýòî - óñûíîâëåííûå ñèðîòû!
     LOOP 
    ENDIF 
   CASE m.tipofcod = 4 && Ïðîôîñìîòðû íåñîâåðøåííîëåòíèõ, ñ ñåíòÿáðÿ
    IF !BETWEEN(m.rslt,332,336) AND m.rslt!=326
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

  INSERT INTO dsp FROM MEMVAR 
   
 ENDSCAN 
 SET ORDER TO sn_pol
 WAIT CLEAR 

 MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÍÀÐÓÆÅÍÎ '+TRANSFORM(m.nusl, '99999')+' ÓÑËÓÃ!',0+64,'')
 
 WAIT "ÐÀÑ×ÅÒ ÂÒÎÐÎÃÎ ÝÒÀÏÀ (ÍÎÂÛÉ ÀËÃÎÐÈÒÌ)..." WINDOW NOWAIT 
 SELECT dsp
 SET ORDER TO c_i
 SELECT talon 
 SET ORDER TO c_i 
 SET RELATION TO LEFT(c_i,25) INTO dsp 

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
  IF !INLIST(m.rslt,320,326,352,353,357,358,390,396)
   LOOP 
  ENDIF 
  IF dsp.k_u2>0
   LOOP 
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
 WAIT CLEAR 

 IF USED('dsp')
  USE IN dsp
 ENDIF 
 IF USED('talon')
  USE IN talon
 ENDIF 
 IF USED('people')
  USE IN people
 ENDIF 
 IF USED('errsv')
  USE IN errsv
 ENDIF 
 IF USED('dspcodes')
  USE IN dspcodes
 ENDIF 

RETURN 
