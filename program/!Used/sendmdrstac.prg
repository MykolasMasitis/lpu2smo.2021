PROCEDURE SendMdrStac

 IF !fso.FolderExists(pOut+'\'+gcperiod+'\������������ �����������')
  MESSAGEBOX(CHR(13)+CHR(10)+'����� �� ������������'+CHR(13)+CHR(10)+;
   '�� �������� ������'+CHR(13)+CHR(10)+'�� �����������!',0+16,'')
 ENDIF 

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

 IF OpenFile(pcommon+'\stmdr','stmdr','shar','mcod')>0
  RETURN .f. 
 ENDIF 
 
 =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\spraboxx', 'sprabo', 'shar', 'lpu_id')
 
 SELECT stmdr

 SCAN
  m.mcod  = mcod
  m.lpuid = lpu_id
  
  WAIT m.mcod WINDOW NOWAIT 
  
  DocName   = pOut+'\'+gcperiod+'\������������ �����������\Pm'+m.mcod+'.doc'
  DocNameSh = 'Pm'+m.mcod

  IF !fso.FileExists(DocName)
   LOOP 
  ENDIF 

  m.cTO  = IIF(SEEK(m.lpuid, 'sprabo', 'lpu_id'), 'usr010@'+ALLTRIM(sprabo.abn_name), '')
  
  m.un_id    = SYS(3)
  m.bansfile = 'b_mdr_' + m.mcod
  m.tansfile = 't_mdr_' + m.mcod
  m.dfile    = 'd_mdr_' + m.mcod
  m.mmid     = m.un_id+'.USR010'+'@'+m.qmail
  m.csubj    = 'Otchet po modernizacii'

  poi = fso.CreateTextFile(pOut+'\'+gcperiod+'\������������ �����������' + '\' + m.tansfile)

  poi.WriteLine('To: '+m.cTO)
  poi.WriteLine('Message-Id: ' + m.mmid)
  poi.WriteLine('Subject: ' + m.csubj)
  poi.WriteLine('Content-Type: multipart/mixed')
  poi.WriteLine('Attachment: '+m.dfile+' '+DocNameSh+'.doc')
 
  poi.Close
 
  fso.CopyFile(pOut+'\'+gcperiod+'\������������ �����������\'+DocNameSh+'.doc', pAisOms+'\usr010\output\'+m.dfile)
  fso.CopyFile(pOut+'\'+gcperiod+'\������������ �����������\'+m.tansfile, pAisOms+'\usr010\output\'+m.bansfile)
  fso.DeleteFile(pOut+'\'+gcperiod+'\������������ �����������\'+m.tansfile)

  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('�� ������ �������� ���������?',4+32,'') == 6
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
 
