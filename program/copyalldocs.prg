PROCEDURE CopyAllDocs
 IF MESSAGEBOX('бш унрхре яйнохпнбюрэ днйслемрш лн'+CHR(13)+CHR(10)+'б ндмс дхпейрнпхч?',4+32,'')=7
  RETURN 
 ENDIF 

 m.docsdir = UPPER(pBase+'\'+m.gcPeriod+'\ALLDOCS')
 m.pdfsdir  = UPPER(pBase+'\'+m.gcPeriod+'\ALLPDFS')

 IF !fso.FolderExists(m.docsdir)
  fso.CreateFolder(m.docsdir)
  IF fso.FolderExists(m.docsdir)
   MESSAGEBOX('дхпейрнпхъ дкъ днйслемрнб янгдюмю'+CHR(13)+CHR(10)+m.docsdir,0+64,'')
  ENDIF 
 ENDIF 
 IF !fso.FolderExists(m.pdfsdir)
  fso.CreateFolder(m.pdfsdir)
  IF fso.FolderExists(m.pdfsdir)
   MESSAGEBOX('дхпейрнпхъ дкъ днйслемрнб янгдюмю'+CHR(13)+CHR(10)+m.pdfsdir,0+64,'')
  ENDIF 
 ENDIF 
 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  MESSAGEBOX('тюик ' + pBase+'\'+m.gcPeriod+'\aisoms.dbf'+CHR(13)+CHR(10)+'ме ясыеярбсер!',0+64,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT aisoms
 SCAN 
  m.mcod  = mcod
  m.lpuid = lpuid
  m.lpufolder = pbase+'\'+m.gcperiod+'\'+m.mcod
  IF !fso.FolderExists(m.lpufolder)
   LOOP 
  ENDIF 

  WAIT m.mcod+'...' WINDOW NOWAIT 

  m.actmek    = 'mc'+STR(m.lpuid,4)+m.qcod+PADL(m.tmonth,2,'0')+RIGHT(STR(m.tyear,4),1)+'.xls'
  m.newactmek = m.mcod+' юйр лщй.xls'
  IF fso.FileExists(m.lpufolder+'\'+m.actmek)
   fso.CopyFile(m.lpufolder+'\'+m.actmek, m.docsdir+'\'+m.newactmek, .t.)
  ENDIF 

  m.mkfile    = 'mk'+STR(m.lpuid,4)+m.qcod+PADL(m.tmonth,2,'0')+RIGHT(STR(m.tyear,4),1)+'.xls'
  m.newmkfile = m.mcod+' пееярп юйрнб лщй.xls'
  IF fso.FileExists(m.lpufolder+'\'+m.mkfile)
   fso.CopyFile(m.lpufolder+'\'+m.mkfile, m.docsdir+'\'+m.newmkfile, .t.)
  ENDIF 

  m.mtfile    = 'mt'+STR(m.lpuid,4)+m.qcod+PADL(m.tmonth,2,'0')+RIGHT(STR(m.tyear,4),1)+'.xls'
  m.newmtfile = m.mcod+' рюакхвмюъ тнплю юйрнб лщй.xls'
  IF fso.FileExists(m.lpufolder+'\'+m.mtfile)
   fso.CopyFile(m.lpufolder+'\'+m.mtfile, m.docsdir+'\'+m.newmtfile, .t.)
  ENDIF 

  m.prfile    = 'pr'+m.qcod+PADL(m.tmonth,2,'0')+RIGHT(STR(m.tyear,4),1)+'.xls'
  m.newprfile = m.mcod+' опнрнйнк опхелйх й нокюре яверю.xls'
  IF fso.FileExists(m.lpufolder+'\'+m.prfile)
   fso.CopyFile(m.lpufolder+'\'+m.prfile, m.docsdir+'\'+m.newprfile, .t.)
  ENDIF 

  m.pdffile    = 'pdf'+m.qcod+PADL(m.tmonth,2,'0')+RIGHT(STR(m.tyear,4),1)+'.xls'
  m.newpdffile = m.mcod+' юйр ондсьебни.xls'
  IF fso.FileExists(m.lpufolder+'\'+m.pdffile)
   fso.CopyFile(m.lpufolder+'\'+m.pdffile, m.pdfsdir+'\'+m.newpdffile, .t.)
  ENDIF 

  WAIT CLEAR 
 ENDSCAN 

 USE IN aisoms
 MESSAGEBOX('OK!',0+64,'')
RETURN 