PROCEDURE ActNSI
 IF MESSAGEBOX(CHR(13)+CHR(10)+'��������������� ���?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 

 pUpdDir = fso.GetParentFolderName(pbin)+'\UPDATE'
 IF !fso.FolderExists(pUpdDir)
  fso.CreateFolder(pUpdDir)
 ENDIF 

 SET DEFAULT TO (pUpdDir)
 csprfile = ''
 csprfile=GETFILE('dbf')
 IF EMPTY(csprfile)
  MESSAGEBOX(CHR(13)+CHR(10)+'�� ������ �� �������!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 ospr = fso.GetFile(csprfile)
 IF LOWER(LEFT(ospr.name,6)) != 'sprspr'
  MESSAGEBOX(CHR(13)+CHR(10)+'��� �� ���������� ���!'+CHR(13)+CHR(10),0+16,'sprsprxx')
  RELEASE ospr 
  RETURN 
 ENDIF 
 
 oSettings.CodePage(csprfile, 866, .t.)
 
 IF OpenFile(csprfile, 'spr', 'excl')>0
  =Exit()
  RETURN 
 ENDIF 
 
 SELECT spr
 
 nfields = FCOUNT('spr')
 IF nfields != 6
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '���������� ����� ('+STR(nfields,1)+') �� ����� 6!',0+16,'')
  =Exit()
  RETURN 
 ENDIF 
 
 IF !checkspr()
  RETURN 
 ENDIF 
 
 MESSAGEBOX(CHR(13)+CHR(10)+'���� ����������!'+CHR(13)+CHR(10),0+64,'')
 
 odir = ospr.ParentFolder.Path
 
 CREATE CURSOR nfiles (fname c(8))
 INDEX ON fname TAG fname
 SET ORDER TO fname
 
 INSERT INTO nfiles (fname) VALUES ('admokr')
 INSERT INTO nfiles (fname) VALUES ('codku_')
 INSERT INTO nfiles (fname) VALUES ('codotd')
 INSERT INTO nfiles (fname) VALUES ('codwdr')
 INSERT INTO nfiles (fname) VALUES ('hopff_')
 INSERT INTO nfiles (fname) VALUES ('isv012')
 INSERT INTO nfiles (fname) VALUES ('kdolg')
 INSERT INTO nfiles (fname) VALUES ('kpresl')
* INSERT INTO nfiles (fname) VALUES ('kspec')
 INSERT INTO nfiles (fname) VALUES ('mkb10_')
 INSERT INTO nfiles (fname) VALUES ('modpac')
 INSERT INTO nfiles (fname) VALUES ('ms_mkb')
 INSERT INTO nfiles (fname) VALUES ('nocodr')
 INSERT INTO nfiles (fname) VALUES ('osoerz')
 INSERT INTO nfiles (fname) VALUES ('osoree')
 INSERT INTO nfiles (fname) VALUES ('ososch')
 INSERT INTO nfiles (fname) VALUES ('profot')
 INSERT INTO nfiles (fname) VALUES ('profus')
 INSERT INTO nfiles (fname) VALUES ('promed	')
 INSERT INTO nfiles (fname) VALUES ('prv002')
 INSERT INTO nfiles (fname) VALUES ('reeskp')
 INSERT INTO nfiles (fname) VALUES ('reesms')
 INSERT INTO nfiles (fname) VALUES ('reesus')
 INSERT INTO nfiles (fname) VALUES ('reesvp')
 INSERT INTO nfiles (fname) VALUES ('rsv009')
 INSERT INTO nfiles (fname) VALUES ('sookod')
 INSERT INTO nfiles (fname) VALUES ('sovmno')
 INSERT INTO nfiles (fname) VALUES ('spr_ul')
 INSERT INTO nfiles (fname) VALUES ('sprabo')
 INSERT INTO nfiles (fname) VALUES ('sprlpu')
 INSERT INTO nfiles (fname) VALUES ('spv015')
 INSERT INTO nfiles (fname) VALUES ('tarimu')
 INSERT INTO nfiles (fname) VALUES ('tersmo')
 INSERT INTO nfiles (fname) VALUES ('tipgrp')
 INSERT INTO nfiles (fname) VALUES ('tipno_')
 INSERT INTO nfiles (fname) VALUES ('z_dsno')
 INSERT INTO nfiles (fname) VALUES ('usvmp')
 INSERT INTO nfiles (fname) VALUES ('vidvp_')
 INSERT INTO nfiles (fname) VALUES ('z_cod_')
 INSERT INTO nfiles (fname) VALUES ('z_dsno')
 INSERT INTO nfiles (fname) VALUES ('territ')
 INSERT INTO nfiles (fname) VALUES ('onreas')
 INSERT INTO nfiles (fname) VALUES ('onstad')
 INSERT INTO nfiles (fname) VALUES ('ontum_')
 INSERT INTO nfiles (fname) VALUES ('onnod_')
 INSERT INTO nfiles (fname) VALUES ('onmet_')
 INSERT INTO nfiles (fname) VALUES ('onlech')
 INSERT INTO nfiles (fname) VALUES ('onhir_')
 INSERT INTO nfiles (fname) VALUES ('onlekl')
 INSERT INTO nfiles (fname) VALUES ('onlekv')
 INSERT INTO nfiles (fname) VALUES ('onluch')
 INSERT INTO nfiles (fname) VALUES ('onlpsh')
 INSERT INTO nfiles (fname) VALUES ('onmrf_')
 INSERT INTO nfiles (fname) VALUES ('onmrds')
 INSERT INTO nfiles (fname) VALUES ('onigh_')
 INSERT INTO nfiles (fname) VALUES ('onigds')
 INSERT INTO nfiles (fname) VALUES ('onmrfr')
 INSERT INTO nfiles (fname) VALUES ('onigrt')
 INSERT INTO nfiles (fname) VALUES ('oncons')
 INSERT INTO nfiles (fname) VALUES ('onpcel')
 INSERT INTO nfiles (fname) VALUES ('onnapr')
 INSERT INTO nfiles (fname) VALUES ('onczab')
 INSERT INTO nfiles (fname) VALUES ('codprv')
 INSERT INTO nfiles (fname) VALUES ('tarion')
 INSERT INTO nfiles (fname) VALUES ('ondopk')
 INSERT INTO nfiles (fname) VALUES ('onopls')
 INSERT INTO nfiles (fname) VALUES ('onprot')
 INSERT INTO nfiles (fname) VALUES ('msext')
 INSERT INTO nfiles (fname) VALUES ('sprved')
 INSERT INTO nfiles (fname) VALUES ('sprnco')

 SELECT spr
 SCAN 
  m.name_eta = ALLTRIM(name_eta)
  m.fname = PADR(LOWER(LEFT(m.name_eta, LEN(m.name_eta)-2)),8)
  IF SEEK(m.fname, 'nfiles')
   m.etname    = ALLTRIM(nfiles.fname)
   m.intr_data = intr_data
   IF fso.FileExists(odir+'\'+m.name_eta+'.dbf')
    *IF m.intr_data<=tdat1
     *MESSAGEBOX(m.name_eta+'.dbf'+' -> '+m.etname+'xx.dbf' ,0+64,'')
    
     LOCAL oEx as Exception
     m.err = .f. 
     TRY 
      fso.CopyFile(odir+'\'+m.name_eta+'.dbf', pcommon+'\'+m.etname+'xx.dbf', .t.) && �.� ?
     CATCH TO oEx
      m.err = .t. 
     ENDTRY 
     IF m.err = .t. 
      MESSAGEBOX('������ ��� ����������� �����!'+CHR(13)+CHR(10)+;
      odir+'\'+m.name_eta+'.dbf'+CHR(13)+CHR(10)+;
      oEx.Message,0+64,'')
      RETURN .F.
     ENDIF 

    *ENDIF 
   ENDIF 
  ENDIF 
 ENDSCAN 
 
 =Exit() 
 MESSAGEBOX('��� ��������������!',0+64,'')
RETURN 

FUNCTION checkspr
 F1Name = FIELD(1)
 IF F1Name != 'SCOD'
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '������������ ������� ���� ('+F1Name+') ������ ���� SCOD!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F1Type = VARTYPE(&F1Name)
 IF F1Type!='C'
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '��� ������� ���� '+F1Name+' ������ ���� CHAR!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F1Size = FSIZE(F1Name)
 IF F1Size != 10
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '����������� ���� '+F1Name+' ������ ���� 10!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

 F2Name = FIELD(2)
 IF F2Name != 'CUR_VER'
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '������������ ������� ���� ('+F2Name+') ������ ���� CUR_VER!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F2Type = VARTYPE(&F2Name)
 IF F2Type!='C'
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '��� ���� '+F2Name+' ������ ���� CHAR!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F2Size = FSIZE(F2Name)
 IF F2Size != 10
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '����������� ���� '+F2Name+' ������ ���� 10!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

 F3Name = FIELD(3)
 IF F3Name != 'FULL_NAME'
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '������������ �������� ���� ('+F3Name+') ������ ���� FULL_NAME!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F3Type = VARTYPE(&F3Name)
 IF F3Type!='C'
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '��� ���� '+F3Name+' ������ ���� CHAR!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F3Size = FSIZE(F3Name)
 IF F3Size != 120
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '����������� ���� '+F3Name+' ������ ���� 120!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

 F4Name = FIELD(4)
 IF F4Name != 'INTR_DATA'
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '������������ ���������� ���� ('+F4Name+') ������ ���� INTR_DATA!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F4Type = VARTYPE(&F4Name)
 IF F4Type!='D'
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '��� ���� '+F4Name+' ������ ���� DATE!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

 F5Name = FIELD(5)
 IF F5Name != 'NAME_ETA'
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '������������ ������ ���� ('+F5Name+') ������ ���� NAME_ETA!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F5Type = VARTYPE(&F5Name)
 IF F5Type!='C'
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '��� ���� '+F5Name+' ������ ���� CHAR!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F5Size = FSIZE(F5Name)
 IF F5Size != 8
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '����������� ���� '+F5Name+' ������ ���� 8!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

 F6Name = FIELD(6)
 IF F6Name != 'CRC_ETA'
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '������������ ������� ���� ('+F6Name+') ������ ���� CRC_ETA!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F6Type = VARTYPE(&F6Name)
 IF F6Type!='C'
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '��� ���� '+F6Name+' ������ ���� CHAR!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F6Size = FSIZE(F6Name)
 IF F6Size != 10
  MESSAGEBOX(CHR(13)+CHR(10)+'�������� ��������� ����� '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   '����������� ���� '+F6Name+' ������ ���� 10!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

RETURN .t.

FUNCTION exit 
 IF USED('spr')
  USE IN spr 
 ENDIF 
 RELEASE ospr  
RETURN 
