PROCEDURE yu_06
 IF MESSAGEBOX('ÑÔÎÐÌÈÐÎÂÀÒÜ ÎÒ×ÅÒ Þ-06?',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pTempl+'\yu_06.xls')
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ØÀÁËÎÍ'+pTempl+'\yu_06.xls',0+64,'')
  RETURN 
 ENDIF 
 
 DIMENSION curdata(41,32)
 curdata = 0

 m.nPeriod=2015
 FOR m.i=0 TO 5
  *m.nPeriod = m.nPeriod + m.i
  FOR m.l=1 TO 12
   IF (m.nPeriod+m.i = 2020) AND m.l>6
    EXIT 
   ENDIF 
   =yu_06_one(STR(m.nPeriod+m.i,4)+PADL(m.l,2,'0'), 3+(7*m.i))
  ENDFOR 
 ENDFOR
 
 CREATE CURSOR curdata (recid i)
 INSERT INTO curdata (recid) VALUES (0)
 m.llResult = X_Report(pTempl+'\yu_06.xls', pBase+'\yu_06.xls', .T.)
 USE IN curdata 
  
RETURN 
 
FUNCTION yu_06_one(para1,para2)
 m.lcPeriod = para1
 m.nRow     = para2
 
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
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\profot', 'profot', 'share', 'otd')>0
  USE IN aisoms 
  IF USED('profot')
   USE IN profot
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sookodxx', 'sookod', 'share', 'er_c')>0
  USE IN aisoms 
  USE IN profot
  IF USED('sookod')
   USE IN sookod
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
  IF RECCOUNT('serr')<=0
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
   m.cod   = cod  
   m.otd   = otd 
   m.k_u   = k_u 
   m.s_all = s_all+IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
   
   m.c_err = LEFT(serr.c_err,2)
   m.osn230 = IIF(!EMPTY(serr.c_err) AND SEEK(m.c_err, 'sookod'), LEFT(sookod.osn230,5), '')
   
   m.iserr = IIF(!EMPTY(serr.c_err), .T., .F.)

   m.usl_ok  = IIF(SEEK(SUBSTR(m.otd,2,2), 'profot'), INT(VAL(profot.usl_ok)), 0)
   m.isonk = IIF(INLIST(SUBSTR(m.otd,4,3),'018','060'),.T.,.F.)
   
   *IF !m.iserr
   * LOOP 
   *ENDIF 
   
   IF m.mcod='0371001'
    curdata(m.nRow+3,3)  = curdata(m.nRow+3,3)  + IIF(!m.iserr, m.k_u, 0)
    curdata(m.nRow+3,4)  = curdata(m.nRow+3,4)  + IIF(!m.iserr, m.s_all, 0)
    curdata(m.nRow+3,5)  = curdata(m.nRow+3,5)  + IIF(m.iserr, m.k_u, 0)
    curdata(m.nRow+3,6)  = curdata(m.nRow+3,6)  + IIF(m.iserr, m.s_all, 0)
    curdata(m.nRow+3,17) = curdata(m.nRow+3,17) + IIF(LEFT(m.osn230,3)='5.1', m.k_u, 0)
    curdata(m.nRow+3,18) = curdata(m.nRow+3,18) + IIF(LEFT(m.osn230,3)='5.1', m.s_all, 0)
    curdata(m.nRow+3,19) = curdata(m.nRow+3,19) + IIF(LEFT(m.osn230,3)='5.2', m.k_u, 0)
    curdata(m.nRow+3,20) = curdata(m.nRow+3,20) + IIF(LEFT(m.osn230,3)='5.2', m.s_all, 0)
    curdata(m.nRow+3,21) = curdata(m.nRow+3,21) + IIF(LEFT(m.osn230,3)='5.3' and LEFT(m.osn230,5)!='5.3.2', m.k_u, 0)
    curdata(m.nRow+3,22) = curdata(m.nRow+3,22) + IIF(LEFT(m.osn230,3)='5.3' and LEFT(m.osn230,5)!='5.3.2', m.s_all, 0)
    curdata(m.nRow+3,23) = curdata(m.nRow+3,23) + IIF(LEFT(m.osn230,3)='5.4', m.k_u, 0)
    curdata(m.nRow+3,24) = curdata(m.nRow+3,24) + IIF(LEFT(m.osn230,3)='5.4', m.s_all, 0)
    curdata(m.nRow+3,25) = curdata(m.nRow+3,25) + IIF(LEFT(m.osn230,3)='5.5', m.k_u, 0)
    curdata(m.nRow+3,26) = curdata(m.nRow+3,26) + IIF(LEFT(m.osn230,3)='5.5', m.s_all, 0)
    curdata(m.nRow+3,27) = curdata(m.nRow+3,27) + IIF(LEFT(m.osn230,3)='5.6', m.k_u, 0)
    curdata(m.nRow+3,28) = curdata(m.nRow+3,28) + IIF(LEFT(m.osn230,3)='5.6', m.s_all, 0)
    curdata(m.nRow+3,29) = curdata(m.nRow+3,29) + IIF(LEFT(m.osn230,3)='5.7', m.k_u, 0)
    curdata(m.nRow+3,30) = curdata(m.nRow+3,30) + IIF(LEFT(m.osn230,3)='5.7', m.s_all, 0)
    curdata(m.nRow+3,31) = curdata(m.nRow+3,31) + IIF(LEFT(m.osn230,3)='5.8', m.k_u, 0)
    curdata(m.nRow+3,32) = curdata(m.nRow+3,32) + IIF(LEFT(m.osn230,3)='5.8', m.s_all, 0)
   ELSE 
    DO CASE 
     CASE m.usl_ok=3 && ÀÏÏ
   	  curdata(m.nRow,3)  = curdata(m.nRow,3)  + IIF(!m.iserr, m.k_u, 0)
      curdata(m.nRow,4)  = curdata(m.nRow,4)  + IIF(!m.iserr, m.s_all, 0)
   	  curdata(m.nRow,5)  = curdata(m.nRow,5)  + IIF(m.iserr, m.k_u, 0)
      curdata(m.nRow,6)  = curdata(m.nRow,6)  + IIF(m.iserr, m.s_all, 0)
      curdata(m.nRow,17) = curdata(m.nRow,17) + IIF(LEFT(m.osn230,3)='5.1', m.k_u, 0)
      curdata(m.nRow,18) = curdata(m.nRow,18) + IIF(LEFT(m.osn230,3)='5.1', m.s_all, 0)
      curdata(m.nRow,19) = curdata(m.nRow,19) + IIF(LEFT(m.osn230,3)='5.2', m.k_u, 0)
      curdata(m.nRow,20) = curdata(m.nRow,20) + IIF(LEFT(m.osn230,3)='5.2', m.s_all, 0)
      curdata(m.nRow,21) = curdata(m.nRow,21) + IIF(LEFT(m.osn230,3)='5.3' and LEFT(m.osn230,5)!='5.3.2', m.k_u, 0)
      curdata(m.nRow,22) = curdata(m.nRow,22) + IIF(LEFT(m.osn230,3)='5.3' and LEFT(m.osn230,5)!='5.3.2', m.s_all, 0)
      curdata(m.nRow,23) = curdata(m.nRow,23) + IIF(LEFT(m.osn230,3)='5.4', m.k_u, 0)
      curdata(m.nRow,24) = curdata(m.nRow,24) + IIF(LEFT(m.osn230,3)='5.4', m.s_all, 0)
      curdata(m.nRow,25) = curdata(m.nRow,25) + IIF(LEFT(m.osn230,3)='5.5', m.k_u, 0)
      curdata(m.nRow,26) = curdata(m.nRow,26) + IIF(LEFT(m.osn230,3)='5.5', m.s_all, 0)
      curdata(m.nRow,27) = curdata(m.nRow,27) + IIF(LEFT(m.osn230,3)='5.6', m.k_u, 0)
      curdata(m.nRow,28) = curdata(m.nRow,28) + IIF(LEFT(m.osn230,3)='5.6', m.s_all, 0)
      curdata(m.nRow,29) = curdata(m.nRow,29) + IIF(LEFT(m.osn230,3)='5.7', m.k_u, 0)
      curdata(m.nRow,30) = curdata(m.nRow,30) + IIF(LEFT(m.osn230,3)='5.7', m.s_all, 0)
      curdata(m.nRow,31) = curdata(m.nRow,31) + IIF(LEFT(m.osn230,3)='5.8', m.k_u, 0)
      curdata(m.nRow,32) = curdata(m.nRow,32) + IIF(LEFT(m.osn230,3)='5.8', m.s_all, 0)

     CASE m.usl_ok=2 && ÄÑÏ
   	  curdata(m.nRow+1,3)  = curdata(m.nRow+1,3)  + IIF(!m.iserr, m.k_u, 0)
      curdata(m.nRow+1,4)  = curdata(m.nRow+1,4)  + IIF(!m.iserr, m.s_all, 0)
   	  curdata(m.nRow+1,5)  = curdata(m.nRow+1,5)  + IIF(m.iserr, m.k_u, 0)
      curdata(m.nRow+1,6)  = curdata(m.nRow+1,6)  + IIF(m.iserr, m.s_all, 0)
      curdata(m.nRow+1,17) = curdata(m.nRow+1,17) + IIF(LEFT(m.osn230,3)='5.1', m.k_u, 0)
      curdata(m.nRow+1,18) = curdata(m.nRow+1,18) + IIF(LEFT(m.osn230,3)='5.1', m.s_all, 0)
      curdata(m.nRow+1,19) = curdata(m.nRow+1,19) + IIF(LEFT(m.osn230,3)='5.2', m.k_u, 0)
      curdata(m.nRow+1,20) = curdata(m.nRow+1,20) + IIF(LEFT(m.osn230,3)='5.2', m.s_all, 0)
      curdata(m.nRow+1,21) = curdata(m.nRow+1,21) + IIF(LEFT(m.osn230,3)='5.3' and LEFT(m.osn230,5)!='5.3.2', m.k_u, 0)
      curdata(m.nRow+1,22) = curdata(m.nRow+1,22) + IIF(LEFT(m.osn230,3)='5.3' and LEFT(m.osn230,5)!='5.3.2', m.s_all, 0)
      curdata(m.nRow+1,23) = curdata(m.nRow+1,23) + IIF(LEFT(m.osn230,3)='5.4', m.k_u, 0)
      curdata(m.nRow+1,24) = curdata(m.nRow+1,24) + IIF(LEFT(m.osn230,3)='5.4', m.s_all, 0)
      curdata(m.nRow+1,25) = curdata(m.nRow+1,25) + IIF(LEFT(m.osn230,3)='5.5', m.k_u, 0)
      curdata(m.nRow+1,26) = curdata(m.nRow+1,26) + IIF(LEFT(m.osn230,3)='5.5', m.s_all, 0)
      curdata(m.nRow+1,27) = curdata(m.nRow+1,27) + IIF(LEFT(m.osn230,3)='5.6', m.k_u, 0)
      curdata(m.nRow+1,28) = curdata(m.nRow+1,28) + IIF(LEFT(m.osn230,3)='5.6', m.s_all, 0)
      curdata(m.nRow+1,29) = curdata(m.nRow+1,29) + IIF(LEFT(m.osn230,3)='5.7', m.k_u, 0)
      curdata(m.nRow+1,30) = curdata(m.nRow+1,30) + IIF(LEFT(m.osn230,3)='5.7', m.s_all, 0)
      curdata(m.nRow+1,31) = curdata(m.nRow+1,31) + IIF(LEFT(m.osn230,3)='5.8', m.k_u, 0)
      curdata(m.nRow+1,32) = curdata(m.nRow+1,32) + IIF(LEFT(m.osn230,3)='5.8', m.s_all, 0)
   
     CASE m.usl_ok=1&& ÊÑ
   	  curdata(m.nRow+2,3)  = curdata(m.nRow+2,3)  + IIF(!m.iserr, m.k_u, 0)
      curdata(m.nRow+2,4)  = curdata(m.nRow+2,4)  + IIF(!m.iserr, m.s_all, 0)
   	  curdata(m.nRow+2,5)  = curdata(m.nRow+2,5)  + IIF(m.iserr, m.k_u, 0)
      curdata(m.nRow+2,6)  = curdata(m.nRow+2,6)  + IIF(m.iserr, m.s_all, 0)
      curdata(m.nRow+2,17) = curdata(m.nRow+2,17) + IIF(LEFT(m.osn230,3)='5.1', m.k_u, 0)
      curdata(m.nRow+2,18) = curdata(m.nRow+2,18) + IIF(LEFT(m.osn230,3)='5.1', m.s_all, 0)
      curdata(m.nRow+2,19) = curdata(m.nRow+2,19) + IIF(LEFT(m.osn230,3)='5.2', m.k_u, 0)
      curdata(m.nRow+2,20) = curdata(m.nRow+2,20) + IIF(LEFT(m.osn230,3)='5.2', m.s_all, 0)
      curdata(m.nRow+2,21) = curdata(m.nRow+2,21) + IIF(LEFT(m.osn230,3)='5.3' and LEFT(m.osn230,5)!='5.3.2', m.k_u, 0)
      curdata(m.nRow+2,22) = curdata(m.nRow+2,22) + IIF(LEFT(m.osn230,3)='5.3' and LEFT(m.osn230,5)!='5.3.2', m.s_all, 0)
      curdata(m.nRow+2,23) = curdata(m.nRow+2,23) + IIF(LEFT(m.osn230,3)='5.4', m.k_u, 0)
      curdata(m.nRow+2,24) = curdata(m.nRow+2,24) + IIF(LEFT(m.osn230,3)='5.4', m.s_all, 0)
      curdata(m.nRow+2,25) = curdata(m.nRow+2,25) + IIF(LEFT(m.osn230,3)='5.5', m.k_u, 0)
      curdata(m.nRow+2,26) = curdata(m.nRow+2,26) + IIF(LEFT(m.osn230,3)='5.5', m.s_all, 0)
      curdata(m.nRow+2,27) = curdata(m.nRow+2,27) + IIF(LEFT(m.osn230,3)='5.6', m.k_u, 0)
      curdata(m.nRow+2,28) = curdata(m.nRow+2,28) + IIF(LEFT(m.osn230,3)='5.6', m.s_all, 0)
      curdata(m.nRow+2,29) = curdata(m.nRow+2,29) + IIF(LEFT(m.osn230,3)='5.7', m.k_u, 0)
      curdata(m.nRow+2,30) = curdata(m.nRow+2,30) + IIF(LEFT(m.osn230,3)='5.7', m.s_all, 0)
      curdata(m.nRow+2,31) = curdata(m.nRow+2,31) + IIF(LEFT(m.osn230,3)='5.8', m.k_u, 0)
      curdata(m.nRow+2,32) = curdata(m.nRow+2,32) + IIF(LEFT(m.osn230,3)='5.8', m.s_all, 0)
     ENDCASE 
   ENDIF 
   
  ENDSCAN 
  SET RELATION OFF INTO serr
  USE IN talon 
  USE IN serr 
  SELECT aisoms 
  
 ENDSCAN 
 USE IN aisoms 
 USE IN profot
 USE IN sookod 
 WAIT CLEAR 
 
RETURN 