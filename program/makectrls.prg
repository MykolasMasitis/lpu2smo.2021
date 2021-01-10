PROCEDURE MakeCtrls

 IF MESSAGEBOX('œ≈–≈‘Œ–Ã»–Œ¬¿“‹ CTRL?',4+32,'')=7
  RETURN 
 ENDIF 

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

 IF OpenFile(pBase+'\&gcPeriod\AisOms', 'AisOms', 'shar')>0
  RETURN .f. 
 ENDIF 
 
 =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\spraboxx', 'sprabo', 'shar')
  
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

  =MakeCtrl(lcPath)
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
 USE IN sprabo
 
 SET ESCAPE &OldEscStatus
RETURN 
 
