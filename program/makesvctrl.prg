PROCEDURE MakeSVCtrl

 IF MESSAGEBOX('—Œ¡–¿“‹ —¬ŒƒÕ€… CTRL?',4+32,'')=7
  RETURN 
 ENDIF 

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

 fso.CopyFile(pTempl+'\Ctrl.dbf', pBase+'\'+m.gcPeriod+'\Ctrl'+m.qcod+'.dbf', .t.)
 IF OpenFile(pBase+'\'+m.gcPeriod+'\Ctrl'+m.qcod, 'sv', 'shar')>0
  IF USED('sv')
   USE IN sv
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(pBase+'\&gcPeriod\AisOms', 'AisOms', 'shar')>0
  RETURN .f. 
 ENDIF 
 

 SELECT AisOms
 SCAN
  m.mcod = mcod

  WAIT mcod WINDOW NOWAIT 

  lcPath = pBase+'\'+m.gcperiod+'\'+mcod
  IF !fso.FolderExists(lcPath)
   LOOP 
  ENDIF 
  IF !fso.FileExists(lcPath+'\people.dbf')
   LOOP 
  ENDIF 
  
  IF !fso.FileExists(lcPath+'\Ctrl'+m.qcod+'.dbf')
   LOOP 
  ENDIF 
  SELECT sv 
  APPEND FROM lcPath+'\Ctrl'+m.qcod

  SELECT aisoms 
  
  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('¬€ ’Œ“»“≈ œ–≈–¬¿“‹ Œ¡–¿¡Œ“ ”?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 
 
 ENDSCAN 

 WAIT CLEAR 

 USE
 USE IN sv 
 
 SET ESCAPE &OldEscStatus
RETURN 
 
