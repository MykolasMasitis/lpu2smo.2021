PROCEDURE AllFlk
 PARAMETERS para1
 
 m.loForm = UPPER(para1.name)

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 
 
 m.totflk = &loForm..get_sum_flk.Value
 SCAN FOR !IsPr
  m.soapsts = soapsts
  IF m.loForm = UPPER('mailsoap') AND m.soapsts!='RECIEVED'
   LOOP 
  ENDIF 
  
  m.locflk = sum_flk
  m.totflk = m.totflk - m.locflk
  REPLACE sum_flk WITH 0
  &loForm..get_sum_flk.Value = m.totflk
  &loForm..get_sum_flk.Refresh
 ENDSCAN 

 m.totflk = &loForm..get_sum_flk.Value
 SCAN FOR !IsPr
  m.soapsts = soapsts
  IF m.loForm = UPPER('mailsoap') AND m.soapsts!='RECIEVED'
   LOOP 
  ENDIF 

  WAIT mcod WINDOW NOWAIT 

  &loForm..refresh
  m.mcod  = mcod
  ppath   = m.pbase+'\'+m.gcperiod+'\'+m.mcod

  IF !fso.FolderExists(ppath)
   LOOP 
  ENDIF 
  
  IF !fso.FileExists(ppath+'\people.dbf') OR ;
     !fso.FileExists(ppath+'\talon.dbf')
   LOOP 
  ENDIF 
  
  =OneFlk(ppath)
  
  SELECT AisOms

  m.locflk = sum_flk
  m.totflk = m.totflk + m.locflk
  &loForm..get_sum_flk.Value = m.totflk
  &loForm..get_sum_flk.Refresh

  IF CHRSAW(0) == .T.
   IF INKEY() == 27
    IF MESSAGEBOX('¬€ ’Œ“»“≈ œ–≈–¬¿“‹ Œ¡–¿¡Œ“ ”?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDSCAN 

 WAIT CLEAR 

 SET ESCAPE &OldEscStatus

 MESSAGEBOX('Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!', 0+64, '')

RETURN 
