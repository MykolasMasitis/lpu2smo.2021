PROCEDURE SendMdrPolk

 IF !fso.FolderExists(pOut+'\'+gcperiod+'\������������ ����������')
  MESSAGEBOX(CHR(13)+CHR(10)+'����� �� ������������'+CHR(13)+CHR(10)+;
   '�� �������� ������'+CHR(13)+CHR(10)+'�� �����������!',0+16,'')
 ENDIF 

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

 MDRFile = 'mdr' + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),2)

 IF OpenFile(pout+'\'+gcperiod+'\'+mdrfile,'stmdr','shar')>0
  RETURN .f. 
 ENDIF 
 
 =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\spraboxx', 'sprabo', 'shar', 'lpu_id')
 =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', 'sprlpu', 'shar', 'mcod')
 
 SELECT stmdr

 m.issent=0
 SCAN
  m.mcod  = mcod

  IF SEEK(m.mcod, 'sprlpu')
   m.lpuid = sprlpu.lpu_id
  ELSE 
   MESSAGEBOX('MCOD '+m.mcod+' �� ������ � ����������� sprlpuxx!',0+48, m.mcod)
   LOOP 
  ENDIF 
  
  WAIT m.mcod WINDOW NOWAIT 
  
  DocName   = pOut+'\'+gcperiod+'\������������ ����������\Pm'+m.mcod+'.doc'
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

  poi = fso.CreateTextFile(pOut+'\'+gcperiod+'\������������ ����������' + '\' + m.tansfile)

  poi.WriteLine('To: '+m.cTO)
  poi.WriteLine('Message-Id: ' + m.mmid)
  poi.WriteLine('Subject: ' + m.csubj)
  poi.WriteLine('Content-Type: multipart/mixed')
  poi.WriteLine('Attachment: '+m.dfile+' '+DocNameSh+'.doc')
 
  poi.Close
 
  fso.CopyFile(pOut+'\'+gcperiod+'\������������ ����������\'+DocNameSh+'.doc', pAisOms+'\usr010\output\'+m.dfile)
  fso.CopyFile(pOut+'\'+gcperiod+'\������������ ����������\'+m.tansfile, pAisOms+'\usr010\output\'+m.bansfile)
  fso.DeleteFile(pOut+'\'+gcperiod+'\������������ ����������\'+m.tansfile)

  m.issent = m.issent+1

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
 
 MESSAGEBOX('���������� '+STR(m.issent,3)+' �������.',0+64,'������������ ����������')
 
RETURN 
 
