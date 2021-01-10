PROCEDURE  MakeOnk01
 IF MESSAGEBOX('ÑÔÎÐÌÈÐÎÂÀÒÜ ÔÎÐÌÓ ÎÍÊ-01,',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(ptempl+'\IF01.xls')
  RETURN 
 ENDIF 
 
 CREATE CURSOR curr (nrec i, lpuid n(6), mcod c(7), lpuname c(150), s_all n(11,2), s_mek n(11,2), s_mee n(11,2), s_ekmp n(11,2))
 SELECT curr
 INDEX on lpuid TAG lpuid
 INDEX on mcod TAG mcod 
 SET ORDER TO mcod
 
 FOR m.nmonth = 1 TO 8
  m.lcmonth  = PADL(m.nmonth,2,'0')
  m.lcperiod = LEFT(m.gcperiod,4) + m.lcmonth

  IF !fso.FolderExists(pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 

  =IF01(m.lcperiod)
  
 ENDFOR 
 
 SELECT curr
 
 m.llResult = X_Report(pTempl+'\IF01.xls', pBase+'\'+m.gcperiod+'\F01.xls', .T.)
 
 USE 
 
RETURN 

FUNCTION IF01(para01)
 PRIVATE m.lcperiod
 m.lcperiod = para01
 IF OpenFile(pBase+'\'+m.lcperiod+'\aisoms', 'aisoms', 'shar', 'lpuid')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.lcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
  USE IN aisoms
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT aisoms
 SCAN 
  m.mcod = mcod 
  IF INT(VAL(SUBSTR(m.mcod,3,2)))<41
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  m.lpuid = lpuid
  m.lpuname = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.fullname, '')
  m.e_mee  = e_mee
  m.e_ekmp = e_ekmp
  IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\e'+m.mcod, 'err', 'shar', 'rid')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   IF USED('err')
    USE IN err
   ENDIF 
   LOOP 
  ENDIF 
  m.s_all = 0 
  m.s_mek = 0 
  SELECT talon 
  SET RELATION TO recid INTO err

  SCAN 
   m.cod = cod 
   IF !IsGsp(m.cod)
    LOOP 
   ENDIF 
   m.s_sum = s_all
   m.s_all = m.s_all + m.s_sum
   m.s_mek = m.s_mek + IIF(!EMPTY(err.c_err), m.s_sum, 0)
  ENDSCAN 

  SET RELATION OFF INTO err
  USE IN err 
  USE IN talon 
  IF !SEEK(m.lpuid, 'curr', 'lpuid')
   INSERT INTO curr (mcod,lpuid,lpuname,s_all,s_mek,s_mee,s_ekmp) VALUES (m.mcod,m.lpuid,m.lpuname,m.s_all,m.s_mek,m.e_mee,m.e_ekmp)
  ELSE 
   m.os_all  = curr.s_all
   m.os_mek  = curr.s_mek
   UPDATE curr SET s_all = m.os_all + m.s_all, s_mek = m.os_mek+m.s_mek,s_mee = s_mee+m.e_mee, s_ekmp = s_ekmp+m.e_ekmp  WHERE lpuid=m.lpuid
  ENDIF 
  SELECT aisoms
 ENDSCAN 
 USE IN aisoms
 USE IN sprlpu

 MESSAGEBOX(m.lcperiod,0+64,'')
 
RETURN 