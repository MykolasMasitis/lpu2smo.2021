PROCEDURE ProfRepT3
 IF MESSAGEBOX('ÑÂÅÄÅÍÈß ÎÁ ÎÁÚÅÌÀÕ È ÑÒÎÈÌÎÑÒÈ'+CHR(13)+CHR(10)+;
 	'ÄÈÑÏÀÍÑÅÐÍÎÃÎ ÍÀÁËÞÄÅÍÈß ÂÇÐÎÑËÎÃÎ ÍÀÑÅËÅÍÈß?'+CHR(13)+CHR(10)+'',4+64, 'ÒÀÁËÈÖÀ 3')=7
 	RETURN 
 ENDIF 
 IF !fso.FileExists(pTempl+'\Prof_t3.xls')
  MESSAGEBOX('ØÀÁËÎÍ '+UPPER('Prof_t3.xls')+' ÍÅ ÍÀÉÄÅÍ!',4+64,'')
  RETURN 
 ENDIF 
 
 CREATE CURSOR p_ppl (sn_pol c(25), vozr n(6,2))
 SELECT p_ppl
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 
 DIMENSION dimdata(3,3)
 dimdata = 0
 
 FOR m.nmonth=1 TO m.tmonth
  m.lcperiod = LEFT(m.gcperiod,4)+PADL(m.nmonth,2,'0')
  m.lcmonth  = PADL(m.nmonth,2,'0')
  IF !fso.FolderExists(m.pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pbase+'\'+m.lcperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(m.pbase+'\'+m.lcperiod+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisom
   ENDIF 
   LOOP 
  ENDIF 
  WAIT m.lcperiod + '...' WINDOW NOWAIT 
  SELECT aisoms
  SCAN 
   m.mcod = mcod 
   IF !fso.FolderExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
    IF USED('talon')
     USE IN talon 
    ENDIF 
    SELECT aisoms 
    LOOP 
   ENDIF 
   IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
    USE IN talon 
    IF USED('people')
     USE IN people
    ENDIF 
    SELECT aisoms 
    LOOP 
   ENDIF 
   SELECT talon 
   IF FIELD('p_cel')!='P_CEL'
   *IF FIELD('dn')!='DN'
    USE IN talon 
    USE IN people 
    SELECT aisoms 
    LOOP 
   ENDIF 
   SET RELATION TO sn_pol INTO people

   SCAN 
    m.p_cel = p_cel
    IF m.p_cel!='1.3'
     LOOP 
    ENDIF 
    *m.dn = dn
    *IF !INLIST(m.dn,1,2) && EMPTY(m.dn)
    * LOOP 
    *ENDIF 
    && Çäåñü ðàáî÷èé êîä!
    m.sn_pol = sn_pol
    m.vozr   = ROUND((m.tdat1 - people.dr)/365.25,2)
    m.s_all  = s_all

    IF !SEEK(m.sn_pol, 'p_ppl')
     INSERT INTO p_ppl FROM MEMVAR 
     dimdata(1,3) = dimdata(1,3) + 1
     IF m.vozr=65
      dimdata(2,3) = dimdata(2,3) + 1
     ENDIF 
     IF m.vozr>65
      dimdata(3,3) = dimdata(3,3) + 1
     ENDIF 
    ENDIF 

    dimdata(1,2) = dimdata(1,2) + m.s_all
    IF m.vozr=65
     dimdata(2,2) = dimdata(2,2) + m.s_all
    ENDIF 
    IF m.vozr>65
     dimdata(3,2) = dimdata(3,2) + m.s_all
    ENDIF 


    && Çäåñü ðàáî÷èé êîä!
   ENDSCAN 
   SET RELATION OFF INTO people
   USE IN talon 
   USE IN people 
   SELECT aisoms
   
  
  ENDSCAN 
  USE IN aisoms
  WAIT CLEAR 
  
 ENDFOR 
 
 m.lcTmpName = pTempl+'\Prof_t3.xls'
 m.lcRepName = pBase+'\'+m.gcPeriod+'\Prof_t3.xls'
 m.IsVisible = .T.
 
 CREATE CURSOR curdata (recid i)
 m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)
 USE IN curdata
 
RETURN 
