PROCEDURE MakeDeads
 
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÂÛ ÕÎÒÈÒÅ ÑÔÎÐÌÈÐÎÂÀÒÜ D-ÔÀÉË?'+CHR(13)+CHR(10), 4+32, '')=7
  RETURN 
 ENDIF 

 m.prperiod = STR(IIF(tmonth=1, tyear-1, tyear),4) + PADL(IIF(tmonth=1, 12, tmonth-1),2,'0')

 IF !fso.FileExists(pbase+'\'+m.prperiod+'\deads.dbf')
  IF MESSAGEBOX(CHR(13)+CHR(10)+'ÍÅ ÑÔÎÐÌÈÐÎÂÀÍ D-ÔÀÉË ÇÀ ÏÐÅÄÛÄÓÙÈÉ ÏÅÐÈÎÄ!'+CHR(13)+CHR(10)+;
  	'ÑÔÎÐÌÈÐÎÂÀÒÜ ÏÓÑÒÎÉ ÔÀÉË Â ÒÅÊÓÙÅÌ ÏÅÐÈÎÄÅ?',0+16,'') = 7
   RETURN
  ELSE
   CREATE TABLE &pBase\&gcPeriod\deads (recid i, period c(6), mcod c(7), sn_pol c(17), c_i c(30), ds c(6),;
   	fam c(20), im c(20), ot c(20), w n(1), dr d, ages n(3), cod n(6), tip c(1), rslt n(3), d_type c(5), d_u d)
   INDEX ON sn_pol TAG sn_pol
   USE 
  ENDIF
 ELSE 
  fso.CopyFile(pbase+'\'+m.prperiod+'\deads.dbf', pbase+'\'+m.gcperiod+'\deads.dbf')
  fso.CopyFile(pbase+'\'+m.prperiod+'\deads.cdx', pbase+'\'+m.gcperiod+'\deads.cdx')
 ENDIF

 m.lcpath = pbase+'\'+m.gcperiod
 m.period = m.gcperiod

 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\deads', 'deads', 'shar', 'sn_pol')>0
  IF USED('deads')
   USE IN deads 
  ENDIF 
  RETURN 
 ENDIF 
 
 IF fso.FileExists(m.pbase+'\'+m.gcperiod+'\nsi\outs.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\outs', 'enp', 'shar', 'enp')>0
   IF USED('enp')
    USE IN enp 
   ENDIF 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\outs', 'kms', 'shar', 'kms', 'again')>0
   IF USED('kms')
    USE IN kms 
   ENDIF 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\outs', 'vsn', 'shar', 'vsn', 'again')>0
   IF USED('vsn')
    USE IN vsn 
   ENDIF 
  ENDIF 
 ENDIF 

 SELECT deads
 *SCAN 
  *m.sn_pol = sn_pol && ðàçìåðíîñòü 17!
  *DO CASE 
  * CASE IsKms(m.sn_pol)
  *  IF SEEK(m.sn_pol, 'kms')
  *   DELETE 
  *  ENDIF 
  * CASE IsVs(m.sn_pol)
  *  IF SEEK(SUBSTR(m.sn_pol,7,9), 'vsn')
  *   DELETE 
  *  ENDIF 
  * OTHERWISE 
  *  IF SEEK(LEFT(m.sn_pol,16), 'enp')
  *   DELETE 
  *  ENDIF 
  *ENDCASE 
 *ENDSCAN 
 
 IF !fso.FileExists(m.pbase+'\'+m.gcperiod+'\aisoms.dbf')
  IF USED('deads')
   USE IN deads 
  ENDIF 
  RETURN 
 ENDIF 
 
 IF OpenFile(m.pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('deads')
   USE IN deads 
  ENDIF 
  RETURN 
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
 
  WAIT m.mcod + "..." WINDOW NOWAIT 
  SELECT talon
  SET RELATION TO sn_pol INTO people
  SET RELATION TO recid INTO errsv ADDITIVE 
  SCAN 
   SCATTER MEMVAR 
   IF !INLIST(m.rslt,10,11,12,105,106,205,206,313)
    LOOP 
   ENDIF
   *IF m.tip<>'5' AND m.d_type<>'5' AND m.cod<>1561 AND !INLIST(m.rslt,10,11,12,313)
   * LOOP 
   *ENDIF
   IF SEEK(m.sn_pol, 'deads')
    LOOP 
   ENDIF 
   IF OCCURS('#', m.c_i)>=3
    LOOP 
   ENDIF 

   *DO CASE 
   * CASE IsKms(m.sn_pol)
   *  IF SEEK(m.sn_pol, 'kms')
   *   LOOP 
   *  ENDIF 
   * CASE IsVs(sn_pol)
   *  IF SEEK(SUBSTR(m.sn_pol,7,9), 'vsn')
   *   LOOP  
   *  ENDIF 
   * OTHERWISE 
   *  IF SEEK(LEFT(m.sn_pol,16), 'enp')
   *   LOOP  
   *  ENDIF 
   *ENDCASE 

   m.fam    = people.fam
   m.im     = people.im
   m.ot     = people.ot
   m.w      = people.w
   m.dr     = people.dr
   m.er     = errsv.c_err
   
   m.adj = CTOD(STRTRAN(DTOC(people.dr), STR(YEAR(people.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
   m.ages = (YEAR(m.d_u) - YEAR(people.dr)) - IIF(m.adj>0, 1, 0)
  
   INSERT INTO deads FROM MEMVAR 
   
  ENDSCAN 
  SET RELATION OFF INTO people
  SET RELATION OFF INTO errsv

  USE IN errsv
  USE IN talon
  USE IN people

  SELECT aisoms
  WAIT CLEAR 
 ENDSCAN && aisoms
 
 IF USED('deads')
  USE IN deads
 ENDIF 
 IF USED('enp')
  USE IN enp
 ENDIF 
 IF USED('kms')
  USE IN kms
 ENDIF 
 IF USED('vsn')
  USE IN vsn
 ENDIF 
 
 MESSAGEBOX('OK!',0+64,'')

RETURN 
