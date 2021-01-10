PROCEDURE LoadStopList
 IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÇÀÃÐÓÇÈÒÜ ÑÒÎÏ-ËÈÑÒ ÍÀ ÏÅÐÈÎÄ?',4+32,'')=7
  RETURN 
 ENDIF 
 
 pUpdDir = fso.GetParentFolderName(pBin)+'\UPDATE'
 IF !fso.FolderExists(pUpdDir)
  fso.CreateFolder(pUpdDir)
 ENDIF 

 m.zipname = 'stop'+SUBSTR(m.gcperiod,5,2)+SUBSTR(m.gcperiod,3,2)+'.zip'
 m.dbfname = 'stop'+SUBSTR(m.gcperiod,5,2)+SUBSTR(m.gcperiod,3,2)+'.dbf'
 IF !fso.FileExists(pUpdDir+'\'+m.zipname)
  MESSAGEBOX('ÔÀÉË '+m.zipname+' Â ÄÈÐÅÊÒÎÐÈÈ '+m.pUpdDir+' ÍÅ ÍÀÉÄÅÍ!',0+64,'')
  RETURN 
 ENDIF 

 m.ffile = fso.GetFile(pUpdDir+'\'+m.zipname)
 IF ffile.size >= 2
  m.fhandl = m.ffile.OpenAsTextStream
  m.lcHead = m.fhandl.Read(2)
  fhandl.Close
 ELSE 
  lcHead = ''
 ENDIF 
 IF EMPTY(m.lcHead)
  MESSAGEBOX('ÔÀÉË '+m.zipname+' ÍÅ ßÂËßÅÒÑß ZIP-ÀÐÕÈÂÎÌ!',0+64,m.lchead)
  RETURN 
 ENDIF 
 IF m.lcHead!='PK'
  MESSAGEBOX('ÔÀÉË '+m.zipname+' ÍÅ ßÂËßÅÒÑß ZIP-ÀÐÕÈÂÎÌ!',0+64,m.lchead)
  RETURN 
 ENDIF 

 IF !UnzipOpen(pUpdDir+'\'+m.zipname)
  MESSAGEBOX('ÔÀÉË '+m.zipname+' ÍÅ ßÂËßÅÒÑß ZIP-ÀÐÕÈÂÎÌ!',0+64,m.lchead)
  RETURN 
 ENDIF 
 IF !UnzipGotoFileByName(m.dbfname)
  UnzipClose()
  MESSAGEBOX('Â ÀÐÕÈÂÅ '+m.zipname+' ÍÅ ÎÁÍÀÐÓÆÅÍ ÔÀÉË ' +m.dbfname+'!',0+64,m.lchead)
  RETURN 
 ENDIF 

 m.UnZipDir  = m.pUpdDir+'\'
 IF fso.FileExists(pUpdDir+'\'+m.dbfname) 
 ELSE 
  WAIT "ÐÀÑÏÀÊÎÂÛÂÀÞ, ÆÄÈÒÅ..." WINDOW NOWAIT 
  UnzipFile(m.UnZipDir)
  WAIT CLEAR 
 ENDIF 
 UnzipClose()

 IF fso.FileExists(pUpdDir+'\'+m.dbfname)
*  MESSAGEBOX('ÔÀÉË '+m.dbfname+' ÐÀÑÏÀÊÎÂÀÍ!',0+64,'')
 ELSE
  MESSAGEBOX('ÎØÈÁÊÀ ÏÐÈ ÐÀÑÏÀÊÎÂÊÅ!',0+64,'')
  RETURN 
 ENDIF 
 
 WAIT "ÓÑÒÀÍÎÂÊÀ 866 ÑÒÐÀÍÈÖÛ..." WINDOW NOWAIT 
 oSettings.CodePage('&pUpdDir\&dbfname', 866, .t.)
 WAIT "ÑÒÐÀÍÈÖÀ 866 ÓÑÒÀÍÎÂËÅÍÀ..." WINDOW NOWAIT 
 WAIT CLEAR 
 
 IF OpenFile(pUpdDir+'\'+STRTRAN(LOWER(m.dbfname),'.dbf',''), 'n_stop', 'excl')>0
  IF USED('n_stop')
   USE IN n_stop 
  ENDIF 
  RETURN 
 ENDIF 

 m.zipname = 'stop'+SUBSTR(m.gcperiod,5,2)+SUBSTR(m.gcperiod,3,2)+'.zip'
 m.dbfname = 'stop'+SUBSTR(m.gcperiod,5,2)+SUBSTR(m.gcperiod,3,2)+'.dbf'
 IF !fso.FileExists(pUpdDir+'\'+m.zipname)
  MESSAGEBOX('ÔÀÉË '+m.zipname+' Â ÄÈÐÅÊÒÎÐÈÈ '+m.pUpdDir+' ÍÅ ÍÀÉÄÅÍ!',0+64,'')
  RETURN 
 ENDIF 

 m.prperiod = STR(IIF(tmonth=1, tyear-1, tyear),4) + PADL(IIF(tmonth=1, 12, tmonth-1),2,'0')

 IF !fso.FileExists(pbase+'\'+m.prperiod+'\stop.dbf')
  IF MESSAGEBOX(CHR(13)+CHR(10)+'ÍÅ ÑÔÎÐÌÈÐÎÂÀÍ ÑÒÎÏ-ËÈÑÒ ÇÀ ÏÐÅÄÛÄÓÙÈÉ ÏÅÐÈÎÄ!'+CHR(13)+CHR(10)+;
  	'ÑÔÎÐÌÈÐÎÂÀÒÜ ÑÒÀÐÒÎÂÛÉ ÔÀÉË Â ÒÅÊÓÙÅÌ ÏÅÐÈÎÄÅ?',0+16,'') = 7
   RETURN
  ELSE
  
  SELECT n_stop
  COPY STRUCTURE TO &pBase\&gcPeriod\stop 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\stop', 'stop', 'excl')>0
   IF USED('stop')
    USE IN stop 
   ENDIF 
   USE IN n_stop 
   RETURN 
  ENDIF 
  SELECT stop 
  INDEX on enp TAG enp 
  USE IN stop 
  ENDIF
 ELSE 
  fso.CopyFile(pbase+'\'+m.prperiod+'\stop.dbf', pbase+'\'+m.gcperiod+'\stop.dbf')
  fso.CopyFile(pbase+'\'+m.prperiod+'\stop.cdx', pbase+'\'+m.gcperiod+'\stop.cdx')
 ENDIF

 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\stop', 'stop', 'shar', 'enp')>0
  IF USED('stop')
   USE IN stop  
  ENDIF 
  USE IN n_stop 
  RETURN 
 ENDIF 
 
 DO CASE
  CASE m.qcod='S7'
   m.q_q = 5400
  CASE m.qcod='R2'
   m.q_q = 111
  CASE m.qcod='I3'
   m.q_q = 5398
  OTHERWISE 
   m.q_q = 0
 ENDCASE 
 
 SELECT n_stop 
 
 SCAN 
  IF q<>m.q_q
   LOOP 
  ENDIF 
  IF ist<>'d'
   LOOP 
  ENDIF 
  
  SCATTER MEMVAR 
  IF !SEEK(m.enp, 'stop')
   INSERT INTO stop FROM MEMVAR 
  ENDIF 

 ENDSCAN 
 USE 
 USE IN stop 

 MESSAGEBOX('OK!',0+64,'')
RETURN 