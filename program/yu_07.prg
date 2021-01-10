PROCEDURE yu_07
 IF MESSAGEBOX('ÑÔÎÐÌÈÐÎÂÀÒÜ ÎÒ×ÅÒ Þ-07?',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pTempl+'\yu_07.xls')
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ØÀÁËÎÍ'+pTempl+'\yu_07.xls',0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pTempl+'\yu_07.dbf')
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÑÏÐÀÂÎ×ÍÈÊ'+pTempl+'\yu_07.dbf',0+64,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pTempl+'\yu_07', 'yu', 'shar', 'cod')>0
  IF USED('yu')
   USE IN yu
  ENDIF 
  RETURN 
 ENDIF 
 
 DIMENSION curdata(6,8)
 curdata = 0
 
 CREATE CURSOR curs_1 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER to sn_pol
 CREATE CURSOR curs_2 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER to sn_pol
 CREATE CURSOR curs_3 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER to sn_pol
 CREATE CURSOR curs_4 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER to sn_pol
 CREATE CURSOR curs_5 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER to sn_pol
 CREATE CURSOR curs_6 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER to sn_pol
 

 m.nPeriod=2020
 FOR m.l=1 TO 12
  =yu_07_one(STR(m.nPeriod,4)+PADL(m.l,2,'0'))
 ENDFOR
 
 CREATE CURSOR curdata (recid i)
 INSERT INTO curdata (recid) VALUES (0)
 m.llResult = X_Report(pTempl+'\yu_07.xls', pBase+'\yu_07.xls', .T.)
 USE IN curdata 
 
 USE IN curs_1
 USE IN curs_2
 USE IN curs_3
 USE IN curs_4
 USE IN curs_5
 USE IN curs_6
 
 USE IN yu
   
RETURN 
 
FUNCTION yu_07_one(para1)
 m.lcPeriod = para1
 
 IF !fso.FolderExists(pBase+'\'+m.lcPeriod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.lcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.lcPeriod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms 
  ENDIF 
  RETURN 
 ENDIF 

 WAIT m.lcPeriod WINDOW NOWAIT 
 
 SELECT aisoms
 SCAN
  m.mcod = mcod
  IF !fso.FolderExists(pBase+'\'+m.lcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
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
  IF OpenFile(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\e'+m.mcod, 'serr', 'shar', 'rid')>0
   USE IN talon 
   IF USED('serr')
    USE IN serr 
   ENDIF 
   SELECT aisoms 
   LOOP 
  ENDIF 
  
  SELECT talon 
  SET RELATION TO recid INTO serr
  SCAN  
   m.sn_pol = sn_pol
   m.cod   = cod  
   m.otd   = otd 
   m.k_u   = k_u 
   m.s_all = s_all+IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
   
   m.c_err = LEFT(serr.c_err,2)
   
   m.iserr = IIF(!EMPTY(serr.c_err), .T., .F.)

   m.isonk = IIF(INLIST(SUBSTR(m.otd,4,3),'018','060'),.T.,.F.)
   
   *IF !m.iserr
   * LOOP 
   *ENDIF 
   
   IF !SEEK(m.cod, 'yu')
    LOOP 
   ENDIF 
   
   m.n_str = yu.n
   m.al = 'curs_'+STR(m.n_str,1)
   
   curdata(m.n_str,6) = curdata(m.n_str,6) + m.k_u
   IF !SEEK(m.sn_pol, al)
    INSERT INTO yu FROM MEMVAR 
    curdata(m.n_str,7) = curdata(m.n_str,7) + 1
   ENDIF 
   curdata(m.n_str,8) = curdata(m.n_str,8) + m.s_all
   
  ENDSCAN 
  SET RELATION OFF INTO serr
  USE IN talon 
  USE IN serr 
  SELECT aisoms 
  
 ENDSCAN 
 USE IN aisoms 
 WAIT CLEAR 
 
RETURN 