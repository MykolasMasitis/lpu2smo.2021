PROCEDURE Flk2Lpu

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 
 
 m.lWasUsedAisoms = .T.
 IF !USED('aisoms')
  m.lWasUsedAisoms = .F.
  IF OpenFile(pBase+'\&gcPeriod\AisOms', 'AisOms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   RETURN .f. 
  ENDIF 
 ENDIF 
  
 m.lWasUsedSprAbo = .T.
 IF !USED('sprabo')
  m.lWasUsedSprAbo = .F.
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\spraboxx', 'sprabo', 'shar')>0
   USE IN aisoms
   IF USED('sprabo')
    USE IN sprabo
   ENDIF 
   RETURN 
  ENDIF 
 ENDIF 
  
 lcMmy = SUBSTR(gcPeriod,5,2)+SUBSTR(gcPeriod,4,1)
 
 SELECT AisOms
 IF gcUser!='OMS'
  SET FILTER TO Usr == PADL(ALLTRIM(gcUser),6)
 ENDIF 

 SCAN
  m.mcod = mcod

  lcPath = pBase+'\'+m.gcperiod+'\'+m.mcod
  IF !fso.FolderExists(lcPath)
   LOOP 
  ENDIF 
  IF !fso.FileExists(lcPath+'\people.dbf') OR !fso.FileExists(lcPath+'\talon.dbf')
   LOOP 
  ENDIF 
  m.bname = bname
  IF EMPTY(m.bname)
   LOOP 
  ENDIF 
  IF erz_status < 2
   LOOP 
  ENDIF 
  
  DDDocNamec = "DD" + STR(lpuid,4)+LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
  DSDocNamec = "DS" + STR(lpuid,4)+LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
  DDDocName = lcpath + '\' + DDDocNamec
  DSDocName = lcpath + '\' + DSDocNamec
  
  *MESSAGEBOX(lcPath+'\b_flk_' + m.mcod,0+64,m.mcod)

 
  IF fso.FileExists(lcPath+'\b_flk_' + m.mcod)
   LOOP 
  ENDIF 
  *IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+mcod+'\Ctrl'+m.qcod+'.dbf')
  * LOOP 
  *ENDIF 

  WAIT m.mcod WINDOW NOWAIT 
  =MakeCtrl(lcPath)
  SELECT aisoms 

  IF !fso.FileExists(lcPath + '\Pr' + m.qcod + lcMmy + '.pdf')
   *LOOP 
   =McPrn(lcPath, .f., .f.)
  ENDIF 
  
  SELECT AisOms


  lcLpuID = lpuid
  m.cTO = 'oms@spuemias.msk.oms'
  m.un_id    = SYS(3)
  m.bansfile = 'b_flk_'  + mcod
  m.tansfile = 't_flk_'  + mcod
  m.d1file   = 'd1_flk_' + mcod
  m.d2file   = 'd2_flk_' + mcod
  m.d3file   = 'dd_' + mcod
  m.d4file   = 'ds_' + mcod
  m.d5file   = 'ud_' + mcod
  m.mmid     = m.un_id+'.'+m.usrmail+'@'+m.qmail
  m.csubj    = 'OMS#'+gcPeriod+'#'+STR(lcLpuID,4)+'##1'

  poi = fso.CreateTextFile(lcPath + '\' + m.tansfile)

  poi.WriteLine('To: '+m.cTO)
  poi.WriteLine('Message-Id: ' + m.mmid)
  poi.WriteLine('Subject: ' + m.csubj)
  poi.WriteLine('Content-Type: multipart/mixed')
  poi.WriteLine('Resent-Message-Id: '+ALLTRIM(cmessage))
  poi.WriteLine('Attachment: '+m.d1file+' Ctrl'+m.qcod+'.dbf')
  poi.WriteLine('Attachment: '+m.d2file+' Pr'+m.qcod+lcMmy+'.pdf')
  IF fso.FileExists(DDDocName+'.pdf')
   poi.WriteLine('Attachment: '+m.d3file+' '+DDDocNamec+'.pdf')
  ENDIF 
  IF fso.FileExists(DSDocName+'.pdf')
   poi.WriteLine('Attachment: '+m.d3file+' '+DSDocNamec+'.pdf')
  ENDIF 
 
  poi.Close
  
  IF fso.FileExists(lcPath+'\'+'Ctrl'+m.qcod+'.dbf') 
   fso.CopyFile(lcPath+'\'+'Ctrl'+m.qcod+'.dbf', pAisOms+'\oms\output\'+m.d1file)
  ENDIF 
  IF fso.FileExists(lcPath+'\'+'Pr'+m.qcod+lcMmy+'.pdf') 
   fso.CopyFile(lcPath+'\'+'Pr'+m.qcod+lcMmy+'.pdf', pAisOms+'\oms\output\'+m.d2file)
  ENDIF 
  IF fso.FileExists(DDDocName+'.pdf')
   fso.CopyFile(DDDocName+'.pdf', pAisOms+'\oms\output\'+m.d3file)
  ENDIF 
  IF fso.FileExists(DSDocName+'.pdf')
   fso.CopyFile(DSDocName+'.pdf', pAisOms+'\oms\output\'+m.d4file)
  ENDIF 

  fso.CopyFile(lcPath+'\'+m.tansfile, pAisOms+'\oms\output\'+m.bansfile)

  fso.CopyFile(lcPath+'\'+m.tansfile, lcPath+'\'+m.bansfile)
  fso.DeleteFile(lcPath+'\'+m.tansfile)
  
  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÏÐÅÐÂÀÒÜ ÎÁÐÀÁÎÒÊÓ?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 
 
 ENDSCAN 

 WAIT CLEAR 

 IF m.lWasUsedAisoms = .F.
  USE IN aisoms
 ENDIF 
 
 IF m.lWasUsedSprAbo = .F.
  USE IN sprabo
 ENDIF 
 
 SET ESCAPE &OldEscStatus
RETURN 
 
