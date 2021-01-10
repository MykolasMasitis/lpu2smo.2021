PROCEDURE ChExpPeriod
 PARAMETERS ppath
 m.mcod = RIGHT(ALLTRIM(m.ppath),7)

 IF MESSAGEBOX('�� ������ �������� ������ ����������?'+CHR(13)+CHR(10),4+32,m.ppath)=7
  RETURN 
 ENDIF 
 IF MESSAGEBOX('�� ������� � ����� ���������?'+CHR(13)+CHR(10),4+32,m.ppath)=7
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(m.ppath)
  MESSAGEBOX('���������� �� �������!'+CHR(13)+CHR(10),0+16,m.ppath)
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.ppath+'\m'+m.mcod+'.dbf')
  MESSAGEBOX('����������� ���� m'+m.mcod+'!',0+16,m.ppath)
  RETURN 
 ENDIF 
 IF OpenFile(m.ppath+'\m'+m.mcod, 'mfile', 'shar')>0
  IF USED('mfile')
   USE IN mfile
  ENDIF 
  RETURN 
 ENDIF 
 
 IF RECCOUNT('mfile')<=0
  IF USED('mfile')
   USE IN mfile
  ENDIF 
  MESSAGEBOX('�� ���������� ��� ���������� �� �����������!'+CHR(13)+CHR(10),0+16,m.ppath)
  RETURN 
 ENDIF 
 
 SELECT et, coun(*) as cnt FROM mfile GROUP BY et INTO CURSOR curet 
 
 m.colexps = RECCOUNT('curet')
 USE IN curet

 IF m.colexps > 1
  USE IN mfile
  MESSAGEBOX('�� ���������� ��� ���������'+CHR(13)+CHR(10)+;
   STR(m.colexps,1)+' ����(��) ����������.'+CHR(13)+CHR(10)+;
   '������ ������� ����������!',0+16,'')
  RETURN 
 ENDIF 

 m.lnyear = YEAR(DATE())
 m.lnmonth = MONTH(DATE())
 m.lcperiod = STR(m.lnyear,4)+PADL(m.lnmonth,2,'0')
 
 DO FORM SetPPeriod
 
 m.lcperiod = STR(m.lnyear,4)+PADL(INT(m.lnmonth),2,'0')
 
 SELECT mfile
 SCAN 
  REPLACE e_period WITH m.lcperiod
 ENDSCAN  
 USE IN mfile
 
 MESSAGEBOX('������ �������!'+CHR(13)+CHR(10),0+64,'')

RETURN 
