FUNCTION  MakeCtrl(lcPath)
 fso.CopyFile(pTempl+'\Ctrl.dbf', lcPath+'\Ctrl'+m.qcod+'.dbf', .t.)

 m.mcod   = SUBSTR(lcPath, RAT('\', lcPath)+1)
 m.lpu_id = 0

 IF !USED('aisoms')
  IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shared')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   RETURN 
  ENDIF 
  m.lpu_id = IIF(SEEK(m.mcod, 'aisoms', 'mcod'), aisoms.lpuid, 0)
  USE IN aisoms
 ELSE
  m.lpu_id = IIF(SEEK(m.mcod, 'aisoms', 'mcod'), aisoms.lpuid, 0)
 ENDIF  

 EFile    = lcPath + '\e' + m.mcod
 nEFile   = lcPath + '\e'
 People   = lcPath + '\People'
 Talon    = lcPath + '\Talon'

 Ctrl = lcPath + '\Ctrl' + m.qcod
 
 pnResult = 0
 pnResult = pnResult + OpenFile("&People", "People", "SHARED", "recid")
 pnResult = pnResult + OpenFile("&Talon", "Talon", "SHARED", "recid")
 pnResult = pnResult + OpenFile("&EFile", "Err", "SHARED")
 pnResult = pnResult + OpenFile("&Ctrl", "Ctrl", "excl")
 m.lSooKodUsed = .T.
 IF !USED('sookod')
  m.lSooKodUsed = .F.
  pnResult = pnResult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sookodxx', "sookod", "SHARED", "er_c")
 ENDIF 

 IF pnResult>0
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('err')
   USE IN err
  ENDIF 
  IF USED('ctrl')
   USE IN ctrl
  ENDIF 
  IF !m.lSooKodUsed
   IF USED('sookod')
    USE IN sookod
   ENDIF 
  ENDIF 
  RETURN 
 ENDIF 


 SELECT Err
 SET RELATION TO RId INTO People
 SET RELATION TO RId INTO Talon ADDITIVE 
 SET RELATION TO LEFT(c_err,2) INTO sookod ADDITIVE 

 SCAN FOR !DELETED()
  m.mmy = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)
  m.file = IIF(f == 'R', 'R' + m.qcod + '.' + m.mmy, 'S' + m.qcod  + '.' + m.mmy)

  IF f == 'R'
   m.recid  = People.recid_lpu
   m.fil_id = 0
  ELSE 
   m.recid  = Talon.recid_lpu
   m.fil_id = talon.fil_id
  ENDIF 
  
  m.errors    = c_err
  m.e_cod     = 0
  m.e_ku      = 0
  m.e_tip     = ''
  m.RefReason = sookod.refreason
  m.et230     = 1 && Если МЭК!
  m.osn230    = sookod.osn230
  
  INSERT INTO Ctrl FROM MEMVAR 
  
 ENDSCAN 
 
 SET RELATION OFF INTO Talon
 SET RELATION OFF INTO People 
 SET RELATION OFF INTO sookod
 
 USE IN err

 pnResult = pnResult + OpenFile("&EFile", "s_err", "SHARED", 'rid')
 pnResult = pnResult + OpenFile("&EFile", "r_err", "SHARED", 'rrid', 'again')
 
 SELECT people 
 SET ORDER TO sn_pol
 SET RELATION TO recid INTO r_err
 SELECT talon 
 SET ORDER TO recid_lpu
 SET RELATION TO sn_pol INTO people
 
 SELECT ctrl
 INDEX ON recid FOR UPPER(LEFT(file,1))='S' TAG recid
 SET ORDER TO recid
 SET RELATION TO recid INTO talon
 
 SCAN 
  IF errors='PKA'
   m.n_err    = r_err.c_err
   m.n_ref    = IIF(SEEK(LEFT(m.n_err,2), 'sookod'), sookod.refreason, '')
   m.n_osn230 = IIF(SEEK(LEFT(m.n_err,2), 'sookod'), sookod.osn230, '')
   
   REPLACE errors WITH m.n_err, refreason WITH m.n_ref, osn230 WITH m.n_osn230

  ENDIF 
 ENDSCAN 
 
 SET RELATION OFF INTO talon
 SET ORDER TO 
 DELETE TAG ALL 

 IF USED('s_err')
  USE IN s_err
 ENDIF 
 IF USED('r_err')
  USE IN r_err
 ENDIF 
 USE IN People
 USE IN Talon 
 USE IN Ctrl
 IF !m.lSooKodUsed
  IF USED('sookod')
   USE IN sookod
  ENDIF 
 ENDIF 

RETURN 