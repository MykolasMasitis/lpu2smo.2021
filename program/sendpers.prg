PROCEDURE SendPers

 IF OpenFile(pBase+'\&gcPeriod\AisOms', 'AisOms', 'shar', 'mcod')>0
  RETURN .f. 
 ENDIF 

 CREATE CURSOR PersMail (mcod c(7), lpuid n(4), sent t, sent_id c(75), ;
 	rcvd t, rcvd_id c(75), "flag" c(2))

 lcPeriod = STR(tYear,4) + PADL(tMonth,2,'0')
 mmy      = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)
 
 SELECT AisOms

 SCAN
  m.mcod = mcod
  m.lpu_id = lpuid

  WAIT m.mcod WINDOW NOWAIT 

  lcPath = pBase+'\'+m.gcperiod+'\'+m.mcod

  IF !fso.FolderExists(lcPath)
   LOOP 
  ENDIF 

  ZipPath = lcPath
  MmyName = 'D'+m.qcod+STR(m.lpu_id,4)+'.'+mmy
  
  IF !fso.FileExists(lcPath+'\'+MmyName)
   LOOP 
  ENDIF 

  m.cTO  = 'oms@mgf.msk.oms'
  
  m.un_id    = SYS(3)
  m.bansfile = 'b_y_' + m.mcod
  m.tansfile = 't_y_' + m.mcod
  m.dfile    = 'd_y_' + m.mcod
  m.mmid     = m.un_id+'.'+m.gcUser+'@'+m.qmail
  m.csubj    = 'OMS#'+lcPeriod+'#'+STR(lpu_id,4)+'##1'

  poi = fso.CreateTextFile(lcPath + '\' + m.tansfile)

  poi.WriteLine('To: '+m.cTO)
  poi.WriteLine('Message-Id: ' + m.mmid)
  poi.WriteLine('Subject: ' + m.csubj)
  poi.WriteLine('Content-Type: multipart/mixed')
  poi.WriteLine('Attachment: '+m.dfile+' '+mmyname)
 
  poi.Close
 
  fso.CopyFile(lcPath+'\'+mmyname, pAisOms+'\oms\output\'+m.dfile)
  fso.CopyFile(lcPath+'\'+m.tansfile, pAisOms+'\oms\output\'+m.bansfile)
  
  INSERT INTO PersMail (mcod, lpuid, sent, sent_id) VALUES (m.mcod, m.lpu_id, DATETIME(), m.mmid)

  SELECT AisOms
  
 ENDSCAN 
 WAIT CLEAR 
 USE 
 
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\persmail.dbf')
  fso.DeleteFile(pBase+'\'+m.gcPeriod+'\persmail.dbf')
 ENDIF 
 SELECT PersMail
 COPY TO &pBase\&gcPeriod\persmail
 USE 
 
 MESSAGEBOX('�������� ���������!',0+64,'')
 
 
 RETURN 
