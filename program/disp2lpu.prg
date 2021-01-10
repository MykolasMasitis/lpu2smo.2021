PROCEDURE disp2Lpu
 
 lcPath = pBase+'\'+m.gcperiod+'\'+mcod
* lcExt    = IIF(m.gcFormat='DOC', 'doc', 'pdf')
 lcExt    = 'pdf'
 DDDocNamec = "DD" + STR(lpuid,4)+LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
 DSDocNamec = "DS" + STR(lpuid,4)+LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
 DDDocName = lcpath + '\' + DDDocNamec
 DSDocName = lcpath + '\' + DSDocNamec
 
 m.lIsDDPdf = .F.
 m.lIsDSPdf = .F.
 
 IF fso.FileExists(DDDocName+'.'+lcext)
  m.lIsDDPdf = .T.
 ENDIF 
 IF fso.FileExists(DSDocName+'.'+lcext)
  m.lIsDSPdf = .T.
 ENDIF 
 
 IF m.lIsDDPdf = .F. AND m.lIsDSPdf = .F.
  MESSAGEBOX(CHR(13)+CHR(10)+'������ ����������'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 

 lcLpuID = lpuid
 m.cTO = ALLTRIM(cfrom)
 m.un_id    = SYS(3)
 m.bansfile = 'b_dsp_'  + mcod
 m.tansfile = 't_dsp_'  + mcod
 m.d1file   = 'd1_dsp_' + mcod
 m.d2file   = 'd2_dsp_' + mcod
 m.mmid     = m.un_id+'.'+m.usrmail+'@'+m.qmail
 m.csubj    = '����� �� ���������������'

 poi = fso.CreateTextFile(lcPath + '\' + m.tansfile)

 poi.WriteLine('To: '+m.cTO)
 poi.WriteLine('Message-Id: ' + m.mmid)
 poi.WriteLine('Subject: ' + m.csubj)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Resent-Message-Id: '+ALLTRIM(cmessage))
 IF fso.FileExists(DDDocName+'.'+lcext)
  poi.WriteLine('Attachment: '+m.d1file+' '+DDDocNamec+'.'+lcext)
 ENDIF 
 IF fso.FileExists(DSDocName+'.'+lcext)
  poi.WriteLine('Attachment: '+m.d2file+' '+DSDocNamec+'.'+lcext)
 ENDIF 
 
 poi.Close
 
 IF fso.FileExists(DDDocName+'.'+lcext)
  fso.CopyFile(DDDocName+'.'+lcext, pAisOms+'\oms\output\'+m.d1file)
 ENDIF 
 IF fso.FileExists(DSDocName+'.'+lcext)
  fso.CopyFile(DSDocName+'.'+lcext, pAisOms+'\oms\output\'+m.d2file)
 ENDIF 

 fso.CopyFile(lcPath+'\'+m.tansfile, pAisOms+'\oms\output\'+m.bansfile)

 fso.CopyFile(lcPath+'\'+m.tansfile, lcPath+'\'+m.bansfile)
 fso.DeleteFile(lcPath+'\'+m.tansfile)

 
 MESSAGEBOX(CHR(13)+CHR(10)+'����� ���������!'+CHR(13)+CHR(10),0+64,'')

RETURN 
 
