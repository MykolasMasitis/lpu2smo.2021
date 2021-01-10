PROCEDURE ImpExpI3
 IF m.IsServer
  MESSAGEBOX('Â ÐÅÆÈÌÅ "ÑÅÐÂÅÐÀ" ÈÌÏÎÐÒ ÍÅ ÄÎÏÓÑÒÈÌ!',0+64,'')
  RETURN 
 ENDIF 
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÂÛ ÕÎÒÈÒÅ ÈÌÏÎÐÒÈÐÎÂÀÒÜ ÐÅÇÓËÜÒÀÒÛ ÝÊÑÏÅÐÒÈÇ?'+CHR(13)+CHR(10),4+32,'ÈÍÃÎÑÑÒÐÀÕ')=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pExpImp)
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß ÝÊÑÏÎÐÒÀ-ÈÌÏÎÐÒÀ'+CHR(13)+CHR(10)+pExpImp+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pMee+'\REQUESTS')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß ÇÀÏÐÎÑÎÂ'+CHR(13)+CHR(10)+pMee+'\REQUESTS'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pMee+'\RSS')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß ÐÅÅÑÒÐÎÂ'+CHR(13)+CHR(10)+pMee+'\RSS'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pMee+'\SVACTS')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß ÑÂÎÄÍÛÕ ÀÊÒÎÂ'+CHR(13)+CHR(10)+pMee+'\SVACTS'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pMee+'\SSACTS')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß ÏÅÐÑÎÍÀËÜÍÛÕ ÀÊÒÎÂ'+CHR(13)+CHR(10)+pMee+'\SSCTS'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pMee+'\REQUESTS\catalog.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÊÀÒÀËÎÃ ÇÀÏÐÎÑÎÂ'+CHR(13)+CHR(10)+pMee+'\REQUESTS\catalog.dbf'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pMee+'\RSS\rss.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÊÀÒÀËÎÃ ÐÅÅÑÒÐÎÂ'+CHR(13)+CHR(10)+pMee+'\RSS\rss.dbf'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pMee+'\SVACTS\svacts.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÊÀÒÀËÎÃ ÑÂÎÄÍÛÕ ÀÊÒÎÂ'+CHR(13)+CHR(10)+pMee+'\SVACTS\svacts.dbf'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pMee+'\SSACTS\ssacts.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÊÀÒÀËÎÃ ÏÅÐÑÎÍÀËÜÍÛÕ ÀÊÒÎÂ'+CHR(13)+CHR(10)+pMee+'\SSACTS\ssacts.dbf'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pMee+'\SVACTS\svacts', 'svacts', 'shar')>0
  IF USED('svacts')
   USE IN svacts
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pMee+'\SVACTS\moves', 'svmoves', 'shar')>0
  USE IN svacts
  IF USED('svmoves')
   USE IN svmoves
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pMee+'\SSACTS\ssacts', 'ssacts', 'shar')>0
  USE IN svacts
  USE IN svmoves
  IF USED('ssacts')
   USE IN ssacts
  ENDIF 
  RETURN 
 ENDIF 
* IF OpenFile(pMee+'\SSACTS\ssmoves', 'ssmoves', 'shar')>0
 IF OpenFile(pMee+'\SSACTS\moves', 'ssmoves', 'shar')>0
  USE IN svacts
  USE IN svmoves
  USE IN ssacts
  IF USED('ssmoves')
   USE IN ssmoves
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pMee+'\REQUESTS\catalog', 'rqst', 'shar')>0
  USE IN svacts
  USE IN ssacts
  USE IN svmoves
  USE IN ssmoves
  IF USED('rqst')
   USE IN rqst
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pMee+'\RSS\rss', 'rss', 'shar')>0
  USE IN svacts
  USE IN ssacts
  USE IN svmoves
  USE IN ssmoves
  USE IN rqst
  IF USED('rss')
   USE IN rss
  ENDIF 
  RETURN 
 ENDIF 

 SELECT MIN(recid) as stval FROM svacts  INTO CURSOR cur_sv
 m.stval = cur_sv.stval
 USE IN cur_sv

 IF m.stval < 100
  MESSAGEBOX('ÔÀÉË SVACTS ÍÅ ÍÀÑÒÐÎÅÍ ÍÀ ÌÍÎÃÎÏÎËÜÇÎÂÀÒÅËÜÑÊÈÉ ÐÅÆÈÌ!'+CHR(13)+CHR(10)+;
  	'ÍÀ×ÀËÜÍÎÅ ÇÍÀ×ÅÍÈÅ ÀÂÒÎÈÍÊÐÅÌÅÍÀ ÑÎÑÒÀÂËßÅÒ '+STR(m.stval,3),0+64,'svacts')
  USE IN svmoves
  USE IN ssmoves
  USE IN svacts
  USE IN ssacts
  USE IN rqst
  USE IN rss
  RETURN 
 ENDIF 

 SELECT MIN(recid) as stval FROM ssacts  INTO CURSOR cur_sv
 m.stval = cur_sv.stval
 USE IN cur_sv

 IF m.stval < 100
  MESSAGEBOX('ÔÀÉË SSACTS ÍÅ ÍÀÑÒÐÎÅÍ ÍÀ ÌÍÎÃÎÏÎËÜÇÎÂÀÒÅËÜÑÊÈÉ ÐÅÆÈÌ!'+CHR(13)+CHR(10)+;
  	'ÍÀ×ÀËÜÍÎÅ ÇÍÀ×ÅÍÈÅ ÀÂÒÎÈÍÊÐÅÌÅÍÀ ÑÎÑÒÀÂËßÅÒ '+STR(m.stval,3),0+64,'ssacts')
  USE IN svmoves
  USE IN ssmoves
  USE IN svacts
  USE IN ssacts
  USE IN rqst
  USE IN rss
  RETURN 
 ENDIF 
 
 SELECT MIN(recid) as stval FROM rqst  INTO CURSOR cur_sv
 m.stval = cur_sv.stval
 USE IN cur_sv

 IF m.stval < 100
  MESSAGEBOX('ÔÀÉË ÇÀÏÐÎÑÎÂ ÍÅ ÍÀÑÒÐÎÅÍ ÍÀ ÌÍÎÃÎÏÎËÜÇÎÂÀÒÅËÜÑÊÈÉ ÐÅÆÈÌ!'+CHR(13)+CHR(10)+;
  	'ÍÀ×ÀËÜÍÎÅ ÇÍÀ×ÅÍÈÅ ÀÂÒÎÈÍÊÐÅÌÅÍÀ ÑÎÑÒÀÂËßÅÒ '+STR(m.stval,3),0+64,'request')
  USE IN svmoves
  USE IN ssmoves
  USE IN svacts
  USE IN ssacts
  USE IN rqst
  USE IN rss
  RETURN 
 ENDIF 

 SELECT MIN(recid) as stval FROM rss  INTO CURSOR cur_sv
 m.stval = cur_sv.stval
 USE IN cur_sv

 IF m.stval < 100
  MESSAGEBOX('ÔÀÉË RSS ÍÅ ÍÀÑÒÐÎÅÍ ÍÀ ÌÍÎÃÎÏÎËÜÇÎÂÀÒÅËÜÑÊÈÉ ÐÅÆÈÌ!'+CHR(13)+CHR(10)+;
  	'ÍÀ×ÀËÜÍÎÅ ÇÍÀ×ÅÍÈÅ ÀÂÒÎÈÍÊÐÅÌÅÍÀ ÑÎÑÒÀÂËßÅÒ '+STR(m.stval,3),0+64,'rss')
  USE IN svmoves
  USE IN ssmoves
  USE IN svacts
  USE IN ssacts
  USE IN rqst
  USE IN rss
  RETURN 
 ENDIF 

 iii = 1
 expimpbak = 'expimp'+PADL(iii,3,'0')
 DO WHILE fso.FileExists(pexpimp+'\'+expimpbak+'.zip')
  iii = iii + 1
  expimpbak = 'expimp'+PADL(iii,3,'0')
 ENDDO 
 ZipOpen(pexpimp+'\'+expimpbak+'.zip')
 m.nFilesInZip = 0

 IF fso.FileExists(pExpImp+'\catalog.dbf')
  ZipFile(pExpImp+'\catalog.dbf')
  m.nFilesInZip = m.nFilesInZip + 1
  fso.DeleteFile(pExpImp+'\catalog.dbf')
 ENDIF 
 IF fso.FileExists(pExpImp+'\catalog.cdx')
  ZipFile(pExpImp+'\catalog.cdx')
  m.nFilesInZip = m.nFilesInZip + 1
  fso.DeleteFile(pExpImp+'\catalog.cdx')
 ENDIF 
 IF fso.FileExists(pExpImp+'\rss.dbf')
  ZipFile(pExpImp+'\rss.dbf')
  m.nFilesInZip = m.nFilesInZip + 1
  fso.DeleteFile(pExpImp+'\rss.dbf')
 ENDIF 
 IF fso.FileExists(pExpImp+'\rss.cdx')
  ZipFile(pExpImp+'\rss.cdx')
  m.nFilesInZip = m.nFilesInZip + 1
  fso.DeleteFile(pExpImp+'\rss.cdx')
 ENDIF 
 IF fso.FileExists(pExpImp+'\svacts.dbf')
  ZipFile(pExpImp+'\svacts.dbf')
  m.nFilesInZip = m.nFilesInZip + 1
  fso.DeleteFile(pExpImp+'\svacts.dbf')
 ENDIF 
 IF fso.FileExists(pExpImp+'\svacts.cdx')
  ZipFile(pExpImp+'\svacts.cdx')
  m.nFilesInZip = m.nFilesInZip + 1
  fso.DeleteFile(pExpImp+'\svacts.cdx')
 ENDIF 
 IF fso.FileExists(pExpImp+'\svacts.fpt')
  ZipFile(pExpImp+'\svacts.fpt')
  m.nFilesInZip = m.nFilesInZip + 1
  fso.DeleteFile(pExpImp+'\svacts.fpt')
 ENDIF 

 IF fso.FileExists(pExpImp+'\svmoves.dbf')
  ZipFile(pExpImp+'\svmoves.dbf')
  m.nFilesInZip = m.nFilesInZip + 1
  fso.DeleteFile(pExpImp+'\svmoves.dbf')
 ENDIF 
 IF fso.FileExists(pExpImp+'\svmoves.cdx')
  ZipFile(pExpImp+'\svmoves.cdx')
  m.nFilesInZip = m.nFilesInZip + 1
  fso.DeleteFile(pExpImp+'\svmoves.cdx')
 ENDIF 

 IF fso.FileExists(pExpImp+'\ssacts.dbf')
  ZipFile(pExpImp+'\ssacts.dbf')
  m.nFilesInZip = m.nFilesInZip + 1
  fso.DeleteFile(pExpImp+'\ssacts.dbf')
 ENDIF 
 IF fso.FileExists(pExpImp+'\ssacts.cdx')
  ZipFile(pExpImp+'\ssacts.cdx')
  m.nFilesInZip = m.nFilesInZip + 1
  fso.DeleteFile(pExpImp+'\ssacts.cdx')
 ENDIF 
 IF fso.FileExists(pExpImp+'\ssacts.fpt')
  ZipFile(pExpImp+'\ssacts.fpt')
  m.nFilesInZip = m.nFilesInZip + 1
  fso.DeleteFile(pExpImp+'\ssacts.fpt')
 ENDIF 

 IF fso.FileExists(pExpImp+'\ssmoves.dbf')
  ZipFile(pExpImp+'\ssmoves.dbf')
  m.nFilesInZip = m.nFilesInZip + 1
  fso.DeleteFile(pExpImp+'\ssmoves.dbf')
 ENDIF 
 IF fso.FileExists(pExpImp+'\ssmoves.cdx')
  ZipFile(pExpImp+'\ssmoves.cdx')
  m.nFilesInZip = m.nFilesInZip + 1
  fso.DeleteFile(pExpImp+'\ssmoves.cdx')
 ENDIF 

 SELECT svacts
 COPY TO &pexpimp\svacts WITH cdx 
 USE 
 SELECT ssacts
 COPY TO &pexpimp\ssacts WITH cdx 
 USE 
 SELECT rqst
 SCAN 
  m.rqfile = PADL(recid,6,'0')
  IF fso.FileExists(pMee+'\REQUESTS\'+m.rqfile+'.dbf') AND fso.FileExists(pMee+'\REQUESTS\'+m.rqfile+'.cdx')
   fso.CopyFile(pMee+'\REQUESTS\'+m.rqfile+'.dbf', pExpImp+'\'+m.rqfile+'.dbf')
   fso.CopyFile(pMee+'\REQUESTS\'+m.rqfile+'.cdx', pExpImp+'\'+m.rqfile+'.cdx')
  ENDIF 
 ENDSCAN 
 COPY TO &pexpimp\catalog WITH cdx
 USE 
 SELECT rss
 COPY TO &pexpimp\rss WITH cdx 
 USE 
 
 FOR nmn=0 TO 12
  curmonth  = IIF(tmonth-nmn>0, tmonth-nmn, 12+tmonth-nmn)
  curyear   = IIF(tmonth-nmn>0, tYear, tYear-1)
  curperiod = STR(curyear,4)+PADL(curmonth,2,'0')

  IF !fso.FolderExists(pbase+'\'+curperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+curperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+curperiod+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP 
  ENDIF 
  
  WAIT curperiod WINDOW NOWAIT 

  SELECT aisoms
  SCAN 
   m.mcod = mcod 
   m.IsVed = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
   IF !fso.FolderExists(pbase+'\'+curperiod+'\'+m.mcod)
    LOOP 
   ENDIF 
   mfile = pbase+'\'+curperiod+'\'+m.mcod+'\m'+m.mcod
   IF !fso.FileExists(mfile+'.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(mfile, 'merror', 'shar')>0
    IF USED('merror')
     USE IN merror
    ENDIF 
    LOOP 
   ENDIF 
   IF RECCOUNT('merror')<=0
    USE IN merror
    LOOP 
   ENDIF 
   USE IN merror
   
   m.impfile = 'i'+curperiod+m.mcod
   IF fso.FileExists(pExpImp+'\'+m.impfile+'.dbf')
    ZipFile(pExpImp+'\'+m.impfile+'.dbf')
    m.nFilesInZip = m.nFilesInZip + 1
    fso.DeleteFile(pExpImp+'\'+m.impfile+'.dbf')
   ENDIF 
   IF fso.FileExists(pExpImp+'\'+m.impfile+'.cdx')
    ZipFile(pExpImp+'\'+m.impfile+'.cdx')
    m.nFilesInZip = m.nFilesInZip + 1
    fso.DeleteFile(pExpImp+'\'+m.impfile+'.cdx')
   ENDIF 
   
   fso.CopyFile(mfile+'.dbf', pExpImp+'\'+m.impfile+'.dbf')
   fso.CopyFile(mfile+'.cdx', pExpImp+'\'+m.impfile+'.cdx')
   IF OpenFile(pExpImp+'\'+m.impfile+'.dbf', 'imp', 'excl')>0
    IF USED('imp')
     USE IN imp
    ENDIF
    SELECT aisoms  
    LOOP 
   ENDIF 
   
   SELECT imp 
   PACK 
   USE IN imp 

   SELECT aisoms 
  ENDSCAN 
  USE IN aisoms
  
  WAIT CLEAR 
 
 NEXT 

 ZipClose()
 IF m.nFilesInZip=0
  fso.DeleteFile(pexpimp+'\'+expimpbak+'.zip')
 ENDIF 

 MESSAGEBOX('OK!', 0+64,'')

RETURN 
