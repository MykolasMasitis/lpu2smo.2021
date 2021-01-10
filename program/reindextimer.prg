PROCEDURE ReindexTimer
 IF !m.IsTimerOn
  IF MESSAGEBOX('ÂÊËÞ×ÈÒÜ ÒÀÉÌÅÐ ÏÅÐÅÈÍÄÅÊÑÀÖÈÈ?',4+32,'ÒÀÉÌÅÐ')=7
   RETURN 
  ENDIF 
 ELSE 
  IF MESSAGEBOX('ÂÛÊËÞ×ÈÒÜ ÒÀÉÌÅÐ ÏÅÐÅÈÍÄÅÊÑÀÖÈÈ?',4+32,'ÒÀÉÌÅÐ')=7
   RETURN 
  ENDIF 
 ENDIF 
 
 PUBLIC ot as Timer 
 ot = CREATEOBJECT("ReindTimer")
 
 m.n_secs = SECONDS()
 m.t_secs = m.t_start*60*60

 IF m.n_secs > m.t_secs 
  m.SecsToStart = (24*60*60 - m.n_secs) + m.t_secs
 ELSE 
  m.SecsToStart = m.t_secs - m.n_secs
 ENDIF 
 
 *m.SecsToStart = 60
 ot.Interval = m.SecsToStart * 1000
 ot.Enabled=.t. 
 
 m.IsTimerOn = !m.IsTimerOn
  
RETURN 

