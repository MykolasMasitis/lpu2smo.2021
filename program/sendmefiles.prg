PROCEDURE SendMEFiles

 IF OpenFile(pBase+'\&gcPeriod\AisOms', 'AisOms', 'shar', 'mcod')>0
  RETURN .f. 
 ENDIF 

 *CREATE CURSOR PersMail (mcod c(7), lpuid n(4), sent t, sent_id c(75), ;
 	rcvd t, rcvd_id c(75), "flag" c(2))

 CREATE CURSOR PersMail (mcod c(7), lpuid n(4), sent t, sent_id c(75), ;
 	rcvd t, c_rcvd t, rcvd_id c(75), c_id c(75), "flag" c(2))

 mmy      = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)
 
 SELECT AisOms

 SCAN
*  IF s_pred-sum_flk <= 0
*   LOOP 
*  ENDIF 

  m.mcod = mcod
  m.lpu_id = lpuid

  WAIT m.mcod WINDOW NOWAIT 

  MEFile  = 'ME'+UPPER(m.qcod)+STR(lpuid,4)
  MEFilep = pOut+'\'+gcPeriod+'\ME'+UPPER(m.qcod)+STR(lpuid,4)
  
  IF !fso.FileExists(MEFilep+'.zip')
   IF fso.FileExists(MEFilep+'.dbf')
    ZipOpen(MEFilep+'.zip')
    ZipFile(MEFilep+'.dbf')
    ZipClose()
   ELSE 
    LOOP 
   ENDIF 
  ENDIF 

  IF !fso.FileExists(MEFilep+'.zip')
   LOOP 
  ENDIF 

  m.cTO  = 'oms@mgf.msk.oms'
  
  m.un_id    = SYS(3)
  m.bansfile = 'b_me_' + m.mcod
  m.tansfile = 't_me_' + m.mcod
  m.dfile    = 'd_me_' + m.mcod
  m.mmid     = m.un_id+'.'+m.gcUser+'@'+m.qmail
  m.csubj    = 'OMS#'+gcPeriod+'#'+STR(lpu_id,4)+'##RM'

  poi = fso.CreateTextFile(pOut+'\'+gcPeriod+'\'+m.tansfile)

  poi.WriteLine('To: '+m.cTO)
  poi.WriteLine('Message-Id: ' + m.mmid)
  poi.WriteLine('Subject: ' + m.csubj)
  poi.WriteLine('Content-Type: multipart/mixed')
  poi.WriteLine('Attachment: '+m.dfile+' '+MEFile+'.zip')
 
  poi.Close
 
  fso.CopyFile(MEFilep+'.zip', pAisOms+'\oms\output\'+m.dfile)
  fso.CopyFile(pOut+'\'+gcPeriod+'\'+m.tansfile, pAisOms+'\oms\output\'+m.bansfile)

  INSERT INTO PersMail (mcod, lpuid, sent, sent_id) VALUES (m.mcod, m.lpu_id, DATETIME(), m.mmid)

  SELECT AisOms
  
 ENDSCAN 
 WAIT CLEAR 
 USE 
 
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\me_mail.dbf')
  fso.DeleteFile(pBase+'\'+m.gcPeriod+'\me_mail.dbf')
 ENDIF 
 SELECT PersMail
 COPY TO &pBase\&gcPeriod\me_mail
 USE 

 MESSAGEBOX('ÎÒÏÐÀÂÊÀ ÇÀÊÎÍ×ÅÍÀ!',0+64,'')

 RETURN 
