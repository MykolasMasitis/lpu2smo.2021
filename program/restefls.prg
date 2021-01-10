PROCEDURE RestEFls
 IF MESSAGEBOX('�� ������ ������������ ����� ������'+CHR(13)+CHR(10)+;
  '�� ������ ����������?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 IF MESSAGEBOX('�� ������� � ����� ���������?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 m.lcpath = pbase+'\'+m.gcperiod
 IF !fso.FolderExists(m.lcpath)
  MESSAGEBOX('����������� ���������� '+m.lcpath,0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.lcpath+'\aisoms.dbf')
  MESSAGEBOX('����������� ���� '+UPPER(m.lcpath)+'\AISOMS.DBF!',0+16,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(m.lcpath+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
 ENDIF 
 
 SELECT aisoms
 SCAN 
  m.mcod = mcod
  WAIT m.mcod + '...' WINDOW NOWAIT 
  IF !fso.FolderExists(m.lcpath+'\'+m.mcod)
   LOOP 
  ENDIF 
  m.lpuid = lpuid
  m.mmy   = PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),1)
  m.dfile = 'D'+m.qcod+STR(m.lpuid,4)+'.'+m.mmy
  m.efile = 'e'+m.mcod+'.dbf'
  m.efilen = 'eo'+m.mcod+'.dbf'
  IF !fso.FileExists(m.lcpath+'\'+m.mcod+'\'+m.dfile)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.lcpath+'\'+m.mcod+'\'+m.efile)
   LOOP 
  ENDIF 

  oDFile = fso.GetFile(m.lcpath+'\'+m.mcod+'\'+m.dfile)
  IF oDFile.size >= 2
   fhandl = oDFile.OpenAsTextStream
   lcHead = fhandl.Read(2)
   fhandl.Close
  ELSE 
   lcHead = ''
  ENDIF 

  IF lcHead != 'PK' && ��� zip-����!
   LOOP 
  ENDIF 
 
  ZipDir  = m.lcpath+'\'+m.mcod+'\'
  rItem   = 'R' + m.qcod + 'Y.' + m.mmy
  sItem   = 'S' + m.qcod + 'Y.' + m.mmy

  IF fso.FileExists(m.lcpath+'\'+m.mcod+'\'+rItem)
   fso.DeleteFile(m.lcpath+'\'+m.mcod+'\'+rItem)
  ENDIF 
  IF fso.FileExists(m.lcpath+'\'+m.mcod+'\'+sItem)
   fso.DeleteFile(m.lcpath+'\'+m.mcod+'\'+sItem)
  ENDIF 

  UnzipOpen(m.lcpath+'\'+m.mcod+'\'+m.dfile)
  UnzipGotoFileByName(rItem)
  UnzipFile(ZipDir)
  UnzipGotoFileByName(sItem)
  UnzipFile(ZipDir)
  UnzipClose()
  
  IF fso.FileExists(m.lcpath+'\'+m.mcod+'\'+rItem) AND fso.FileExists(m.lcpath+'\'+m.mcod+'\'+sItem)
   IF OpenFile(m.lcpath+'\'+m.mcod+'\e'+m.mcod, 'err', 'excl')>0
    IF USED('err')
     USE IN err
    ENDIF 
    LOOP 
   ENDIF 
   IF OpenFile(m.lcpath+'\'+m.mcod+'\'+rItem, 'people', 'shar')>0
    IF USED('people')
     USE IN people
    ENDIF 
    IF USED('err')
     USE IN err
    ENDIF 
    LOOP 
   ENDIF 
   IF OpenFile(m.lcpath+'\'+m.mcod+'\'+sItem, 'talon', 'shar')>0
    IF USED('talon')
     USE IN talon
    ENDIF 
    IF USED('people')
     USE IN people
    ENDIF 
    IF USED('err')
     USE IN err
    ENDIF 
    LOOP 
   ENDIF 
   
   SELECT err
   COPY TO m.lcpath+'\'+m.mcod+'\'+m.efilen CDX 
   ZAP 
   SELECT people
   SCAN 
    IF EMPTY(er_c)
     LOOP 
    ENDIF 
    m.f = 'R'
    m.c_err = er_c
    m.rid = INT(VAL(recid))
    INSERT INTO err FROM MEMVAR 
   ENDSCAN 
   SELECT talon 
   SCAN 
    IF EMPTY(er_c)
     LOOP 
    ENDIF 
    m.f = 'S'
    m.c_err = er_c
    m.rid = INT(VAL(recid))
    INSERT INTO err FROM MEMVAR 
   ENDSCAN 
   USE IN people
   USE IN talon 
   USE IN err
   
   IF fso.FileExists(m.lcpath+'\'+m.mcod+'\'+rItem)
    fso.DeleteFile(m.lcpath+'\'+m.mcod+'\'+rItem)
   ENDIF 
   IF fso.FileExists(m.lcpath+'\'+m.mcod+'\'+sItem)
    fso.DeleteFile(m.lcpath+'\'+m.mcod+'\'+sItem)
   ENDIF 

   SELECT aisoms 
     
  ENDIF 
  
 ENDSCAN 
 WAIT CLEAR 
 USE IN aisoms
 
 MESSAGEBOX('��������� ���������!',0+64,'')
 
RETURN 