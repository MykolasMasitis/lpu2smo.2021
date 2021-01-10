PROCEDURE MakeOuts
 IF MESSAGEBOX(CHR(13)+CHR(10)+'���������� ��������?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 

 pUpdDir = fso.GetParentFolderName(pbin)+'\UPDATE'
 IF !fso.FolderExists(pUpdDir)
  fso.CreateFolder(pUpdDir)
 ENDIF 

 m.zipname = 'nomp'+SUBSTR(m.gcperiod,5,2)+SUBSTR(m.gcperiod,3,2)+'.zip'
 m.dbfname = 'nomp'+SUBSTR(m.gcperiod,5,2)+SUBSTR(m.gcperiod,3,2)+'.dbf'
 IF !fso.FileExists(pUpdDir+'\'+m.zipname)
  MESSAGEBOX('���� '+m.zipname+' � ���������� '+m.pUpdDir+' �� ������!',0+64,'')
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
  MESSAGEBOX('���� '+m.zipname+' �� �������� ZIP-�������!',0+64,m.lchead)
  RETURN 
 ENDIF 
 IF m.lcHead!='PK'
  MESSAGEBOX('���� '+m.zipname+' �� �������� ZIP-�������!',0+64,m.lchead)
  RETURN 
 ENDIF 

 IF !UnzipOpen(pUpdDir+'\'+m.zipname)
  MESSAGEBOX('���� '+m.zipname+' �� �������� ZIP-�������!',0+64,m.lchead)
  RETURN 
 ENDIF 
 IF !UnzipGotoFileByName(m.dbfname)
  UnzipClose()
  MESSAGEBOX('� ������ '+m.zipname+' �� ��������� ���� ' +m.dbfname+'!',0+64,m.lchead)
  RETURN 
 ENDIF 

 m.UnZipDir  = m.pUpdDir+'\'
 IF fso.FileExists(pUpdDir+'\'+m.dbfname) 
 ELSE 
  WAIT "������������, �����..." WINDOW NOWAIT 
  UnzipFile(m.UnZipDir)
  WAIT CLEAR 
 ENDIF 
 UnzipClose()

 IF fso.FileExists(pUpdDir+'\'+m.dbfname)
*  MESSAGEBOX('���� '+m.dbfname+' ����������!',0+64,'')
 ELSE
  MESSAGEBOX('������ ��� ����������!',0+64,'')
  RETURN 
 ENDIF 
 
 WAIT "��������� 866 ��������..." WINDOW NOWAIT 
 oSettings.CodePage('&pUpdDir\&dbfname', 866, .t.)
 WAIT "�������� 866 �����������..." WINDOW NOWAIT 
 WAIT CLEAR 

 IF OpenFile(pUpdDir+'\'+'nomp'+SUBSTR(m.gcperiod,5,2)+SUBSTR(m.gcperiod,3,2), 'nomp', 'excl')>0
  RETURN 
 ENDIF 

 SELECT nomp
 COPY STRUCTURE TO &pbase\&gcperiod\nsi\outs
 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\outs', 'outs', 'excl')>0
  IF USED('outs')
   USE IN outs
  ENDIF 
  USE IN nomp 
  RELEASE ospr 
  RETURN 
 ENDIF 
 
 SELECT outs
 INDEX ON s_card + ' ' + PADL(n_card,10,'0') TAG kms
 INDEX ON vsn tag vsn 
 INDEX ON enp TAG enp 

 WAIT "��������� ���������..." WINDOW NOWAIT 
 SELECT nomp 
 m.nRecs = 0 
 SCAN 
  IF q!=m.qcod
   *LOOP 
  ENDIF 
  IF q=m.qcod
   m.nRecs = m.nRecs + 1
  ENDIF 
  SCATTER MEMVAR 
  INSERT INTO outs FROM MEMVAR 
 ENDSCAN 
 WAIT CLEAR 

 USE IN nomp
 
 * ���� �������� ��������� � ���15
 SELECT lpu_tera as lpuid, coun(*) as paz FROM outs GROUP BY lpu_tera ;
	WHERE lpu_tera>0 AND q=m.qcod ORDER BY lpuid INTO CURSOR nomp_stat READWRITE 
 SELECT nomp_stat
 SUM paz TO m.paz_outs
 INDEX on lpuid TAG lpuid 
 SET ORDER TO lpuid 
 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\oms15.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\oms15', 'oms15', 'shar')>0
   IF USED('oms15')
    USE IN oms15
   ENDIF 
  ELSE 
   SELECT oms15
   SET RELATION TO lpuid INTO nomp_stat
   REPLACE ALL paz_outs WITH nomp_stat.paz
   
   SUM paz TO m.paz 
   SUM paz_outs TO m.paz_outs
   
   SET RELATION OFF INTO nomp_stat
   USE IN nomp_stat
   USE IN oms15
   
   IF m.paz=m.paz_outs
    MESSAGEBOX('���-�� �������������� �� ���������: '+TRANSFORM(m.paz_outs,'99999999')+CHR(13)+CHR(10)+;
    	'������������� ������ ����� ���-15 : '+TRANSFORM(m.paz,'99999999'), 0+64, '')
   ELSE 
    MESSAGEBOX('���-�� �������������� �� ���������  : '+TRANSFORM(m.paz_outs,'99999999')+CHR(13)+CHR(10)+;
    	'�� ������������� ������ ����� ���-15: '+TRANSFORM(m.paz,'9999999'), 0+64, '')
   ENDIF 
  
  ENDIF 
 ELSE 
  MESSAGEBOX('������ ����� ���-15 �� ���������.'+CHR(13)+CHR(10)+;
             '���-�� �������������� �� ���������: '+TRANSFORM(m.paz_outs,'99999999'), 0+64, '')
 ENDIF 

 USE IN outs

 MESSAGEBOX(CHR(13)+CHR(10)+'��������� ��������� ���������!'+CHR(13)+CHR(10), 0+64, '')

RETURN 