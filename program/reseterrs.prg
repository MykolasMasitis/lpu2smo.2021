PROCEDURE ResetErrs(calias)

 m.et = goApp.etap
 DO FORM SelTipOfExp TO m.lResp 
 
 IF INLIST(m.et,'4','5','6','9')
  IF EMPTY(goApp.supexp)
   MESSAGEBOX('НЕ ВЫБРАН ЭКСПЕРТ ЭКМП!',0+16,'ЭКМП')
   RETURN 0 
  ENDIF 
 ELSE 
  *goApp.supexp = IIF(goApp.supexp<>"EXP009", goApp.supexp, "")
  goApp.supexp = IIF(goApp.supexp="EXP009", goApp.supexp, "")
 ENDIF 
 
 IF !m.lResp 
  RETURN 
 ENDIF 

 oal   = ALIAS()
 orecp = RECNO()

 IF fso.FileExists(pmee+'\ssacts\moves'+'.dbf')
  IF OpenFile(pmee+'\ssacts\moves', 'moves', 'shar')>0
   IF USED('moves')
    USE IN moves
   ENDIF 
   RETURN 0
  ENDIF 
 ENDIF 

 SELECT people

 m.nlocked=0
 IF calias = 'people'
  m.nlocked = viewexp.selected
 ELSE 
  COUNT FOR ISRLOCKED() TO m.nlocked
 ENDIF 
 
 SELECT &oal
 GO (orecp)

 IF m.nlocked <= 0
  oal   = ALIAS()
  orecp = RECNO()
  m.ppolis = sn_pol
  SELECT (calias)
  m.IsAnnul = .F.
  m.IsActDeleted = .f.

  SCAN FOR sn_pol = ppolis
   m.recid = recid 
   *m.vvir = PADL(m.recid,6,'0')+m.et+goApp.supexp+goApp.reason && PADL(recid,6,'0')+et+docexp+reason
   m.vvir = PADL(m.recid,6,'0')+m.et+goApp.supexp && По просьбе Ингосса - тяжело снимать ошибки
   IF SEEK(m.vvir, 'merror', 'id_et')

    IF goApp.smoexp=merror.usr OR goApp.smoexp="EXP009"&& удаляем только свои ошибки!

    m.t_akt  = merror.t_akt && тип акта SS or SV
    m.act_id = INT(VAL(SUBSTR(merror.n_akt,10)))
    m.actid = INT(VAL(SUBSTR(merror.n_akt,10)))
    IF USED('moves') AND m.IsAnnul = .F.
     m.IsAnnul = .T.
     ool = ALIAS()
     SELECT MAX(dat) as maxdat FROM moves GROUP BY actid INTO CURSOR curlst WHERE actid = m.actid
     m.maxdat = curlst.maxdat
     SELECT recid as lastid, et as lastet FROM moves INTO CURSOR curlst WHERE actid = m.actid AND dat=m.maxdat
     m.lastet = curlst.lastet
     m.lastid = curlst.lastid
     USE IN curlst
     SELECT (ool)
     IF m.lastet = '1'
      INSERT INTO moves (actid,et,usr,dat) VALUES (m.actid,'2',m.gcUser,DATETIME())
     ENDIF 
    ENDIF 
    IF USED('ssacts') AND m.t_akt='SS' AND m.IsActDeleted = .F.
     m.IsActDeleted = .t.
     DELETE FROM ssacts WHERE recid=m.act_id
    ENDIF 
    DELETE FROM merror WHERE recid=m.recid AND et=m.et AND docexp=goApp.supexp  && По просьбе Ингосса 15.06.2017!
   ENDIF 

   ENDIF && удаляем только свои ошибки!
  ENDSCAN 
  SELECT &oal
  GO (orecp)
 ELSE 
  IF MESSAGEBOX('СНЯТЬ ОШИБКИ СО ВСЕХ ОТОБРАННЫХ (ДА)?'+CHR(13)+CHR(10)+;
   'ИЛИ ТОЛЬКО С ТЕКУЩЕЙ ЗАПИСИ? (НЕТ)'+CHR(13)+CHR(10),4+32,'')=6

   SELECT people
   SCAN FOR ISRLOCKED()
    m.ppolis = sn_pol
    m.IsActDeleted = .f.

    SELECT (calias)
    m.IsAnnul = .F.
    SCAN FOR sn_pol = ppolis
     m.recid = recid 
     m.vvir = PADL(m.recid,6,'0')+m.et+goApp.supexp && По просьбе Ингосса 15.06.2017!
     IF SEEK(m.vvir, 'merror', 'id_et')
      
      IF goApp.smoexp=merror.usr OR goApp.smoexp="EXP009"&& удаляем только свои ошибки!
      
      m.t_akt  = merror.t_akt && тип акта SS or SV
      m.act_id = INT(VAL(SUBSTR(merror.n_akt,10)))
      m.actid = INT(VAL(SUBSTR(merror.n_akt,10)))
      IF USED('moves') AND m.IsAnnul = .F.
       m.IsAnnul = .T.
       ool = ALIAS()
       SELECT MAX(dat) as maxdat FROM moves GROUP BY actid INTO CURSOR curlst WHERE actid = m.actid
       m.maxdat = curlst.maxdat
       SELECT recid as lastid, et as lastet FROM moves INTO CURSOR curlst WHERE actid = m.actid AND dat=m.maxdat
       m.lastet = curlst.lastet
       m.lastid = curlst.lastid
       USE IN curlst
       SELECT (ool)
       IF m.lastet = '1'
        INSERT INTO moves (actid,et,usr,dat) VALUES (m.actid,'2',m.gcUser,DATETIME())
       ENDIF 
      ENDIF 
      IF USED('ssacts') AND m.t_akt='SS' AND m.IsActDeleted = .F.
       m.IsActDeleted = .t.
       DELETE FROM ssacts WHERE recid=m.act_id
      ENDIF 
      DELETE FROM merror WHERE recid=m.recid AND et=m.et AND docexp=goApp.supexp  && По просьбе Ингосса 15.06.2017!
     ENDIF 
     
     ENDIF 

    ENDSCAN 
    SELECT people 
   
   ENDSCAN 
   SELECT &oal
   GO (orecp)

  ELSE 
   oal   = ALIAS()
   orecp = RECNO()
   m.ppolis = sn_pol
   SELECT (calias)
   m.IsAnnul = .F.
   m.IsActDeleted = .f.
   SCAN FOR sn_pol = ppolis
    m.recid = recid 
    m.vvir = PADL(m.recid,6,'0')+m.et+goApp.supexp
    IF SEEK(m.vvir, 'merror', 'id_et')
    
     IF goApp.smoexp=merror.usr OR goApp.smoexp="EXP009"&& удаляем только свои ошибки!
     
     m.t_akt  = merror.t_akt && тип акта SS or SV
     m.act_id = INT(VAL(SUBSTR(merror.n_akt,10)))
     m.actid = INT(VAL(SUBSTR(merror.n_akt,10)))
     IF USED('moves')
      ool = ALIAS()
      SELECT MAX(dat) as maxdat FROM moves GROUP BY actid INTO CURSOR curlst WHERE actid = m.actid
      m.maxdat = curlst.maxdat
      SELECT recid as lastid, et as lastet FROM moves INTO CURSOR curlst WHERE actid = m.actid AND dat=m.maxdat
      m.lastet = curlst.lastet
      m.lastid = curlst.lastid
      USE IN curlst
      SELECT (ool)
      IF m.lastet = '1'
       INSERT INTO moves (actid,et,usr,dat) VALUES (m.actid,'2',m.gcUser,DATETIME())
      ENDIF 
     ENDIF 
     IF USED('ssacts') AND m.t_akt='SS' AND m.IsActDeleted = .F.
      m.IsActDeleted = .t.
      DELETE FROM ssacts WHERE recid=m.act_id
     ENDIF 
     DELETE FROM merror WHERE recid=m.recid AND et=m.et AND docexp=goApp.supexp
    ENDIF 
    
    ENDIF 
    
   ENDSCAN 
   SELECT &oal
   GO (orecp)
  ENDIF 
 ENDIF 

 IF USED('moves')
  USE IN moves
 ENDIF 

RETURN 