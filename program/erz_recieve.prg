PROCEDURE erz_recieve
PARAMETERS loForm

m.UserDir = m.usrmail

IF !fso.FolderExists(pAisOms+'\'+m.UserDir)
 MESSAGEBOX('����������� ���������� ' + pAisOms + '\' + m.UserDIR + '!', 0+16, '')
 RETURN 0
ENDIF 

SELECT AisOms

tt_obr = 0

SCAN FOR !DELETED()
* IF !EMPTY(erz_id)
 m.soapsts = soapsts
 m.bname   = bname
 
 IF EMPTY(m.bname)
  LOOP 
 ENDIF 

* IF LOWER(m.loForm.name) = 'mailsoap' AND m.soapsts!='RECIEVED'
*  LOOP 
* ENDIF 

 IF !EMPTY(erz_id) AND erz_status=1 && �������� �� ������ ������ - �������� �� ����� ��� �� 2-3 �������
  lcDir = pBase + '\' + m.gcperiod + '\' + mcod
  m.seek_id = ALLTRIM(erz_id)

  Fb = ScanDir('&PAisOms\&UserDIR\InPut','b*.*','ERZ','&seek_id')

  IF !EMPTY(Fb)
   m.tt_obr = m.tt_obr+1
   poi = FOPEN('&PAisOms\&UserDIR\InPut\&Fb')
   IF poi != -1
    y = ''
    =FSEEK(poi,0)
    DO WHILE !FEOF(poi)
     y = FGETS(poi)
     IF y = 'Message-Id'
*      m.erz_m_id = ALLTRIM(SUBSTR(y, AT(':',y)+1))
     ENDIF 
     IF y = 'Attachment'
      Arg1 = ALLTRIM(SUBSTR(y, AT(':',y)+1, AT(' ',y,2) - AT(' ',y,1)))
      Arg2 = ALLTRIM(SUBSTR(y, AT(' ',y,2)+1))
     ENDIF 
    ENDDO  
    =FCLOSE(poi)
    IF !File('&PAisOms\&UserDIR\InPut\&Arg1') 
     WAIT "����������� ��� ���������� �������������� ����!" WINDOW
    ELSE 
     WAIT "��������� ������..." WINDOW NOWAIT 
     fso.CopyFile(PAisOms+'\'+m.UserDIR+'\InPut\'+Arg1, lcDir+'\Answer.Dbf', .t.)
     fso.DeleteFile(PAisOms+'\'+m.UserDIR+'\InPut\'+m.Fb)
     fso.DeleteFile(PAisOms+'\'+m.UserDIR+'\InPut\'+m.Arg1)
     m.lcpath = pBase+'\'+m.gcperiod+'\'+mcod
     =OneERZProcess(m.lcpath)
     WAIT "����� �� ������ �������!" WINDOW NOWAIT 
*    REPLACE erz_m_id WITH m.erz_m_id
    ENDIF 
   ELSE 
    WAIT "���������� ������� ����-�������!" WINDOW
   ENDIF
  ELSE && ���� ������ ���!
  ENDIF 
 ELSE && ���� ������ �� �����������!
 ENDIF 
 loForm.refresh
ENDSCAN 
WAIT CLEAR 

MESSAGEBOX("����������: "+STR(m.tt_obr,3)+" ������",0+64,"")

RETURN 


FUNCTION OneERZProcess(lcDir)
 IF !fso.FileExists(lcDir + '\People.dbf') OR ;
    !fso.FileExists(lcDir + '\e'+mcod+'.dbf') OR ;
    !fso.FileExists(lcDir + '\Answer.dbf')
  RETURN .F.
 ENDIF 
 m.t_0 = SECONDS()
 oSettings.CodePage(lcDir+'\Answer.dbf', 866)

 tn_result = 0
 tn_result = tn_result + OpenFile(lcDir+'\People', 'People', 'Share')
 tn_result = tn_result + OpenFile(lcDir+'\Talon', 'Talon', 'Share', 'sn_pol')
 tn_result = tn_result + OpenFile(lcDir+'\e'+mcod, 'sError', 'Share', 'rid')
 tn_result = tn_result + OpenFile(lcDir+'\e'+mcod, 'rError', 'Share', 'rrid', 'again')
 tn_result = tn_result + OpenFile(lcDir+'\Answer', 'Answer', 'Excl')
 tn_result = tn_result + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\osoerzxx', 'OsoERZ', 'Shar', 'ans_r')
 
 IF tn_result>0
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('talon')
   USE IN talon 
  ENDIF 
  IF USED('serror')
   USE IN serror
  ENDIF 
  IF USED('rerror')
   USE IN rerror
  ENDIF 
  IF USED('answer')
   USE IN answer
  ENDIF 
  IF USED('osoerz')
   USE IN osoerz
  ENDIF 
  RETURN .f. 
 ENDIF 
 
* IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\kms.dbf')
*   =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\kms','kms','shar','vs')
* ENDIF 
 
 SELECT Answer 
 DELETE TAG ALL 
 INDEX ON RecId TAG RecId
 SET ORDER TO RecId
 
 SELECT People
 SET RELATION TO PADL(RecId,6,'0') INTO Answer
 
 SCAN 
  m.recid  = recid
  m.sn_pol = sn_pol
  m.IsVs = IsVs(m.sn_pol)
  m.d_type = d_type
  
  IF EMPTY(Answer.RecId)
   LOOP 
  ENDIF 
  
  m.llpu = Answer.lpu_id
  m.prmcod = ''
  IF m.llpu>0
   IF SEEK(m.llpu, 'pilot')
    m.prmcod = IIF(SEEK(m.llpu, 'sprlpu'), sprlpu.mcod, '')
   ELSE 
    MESSAGEBOX('��������!'+CHR(13)+CHR(10)+;
    'LPU_ID '+STR(m.llpu,4)+' ������������'+CHR(13)+CHR(10)+;
     '���������� � ����������� PILOT.DBF!',0+48,'')
   ENDIF 
  ENDIF 
  REPLACE qq WITH Answer.Q, sv WITH Answer.ans_r, prmcod WITH IIF(m.d_type!='9', m.prmcod, '')
  m.sv = sv
  
  m.IsGood = IIF(SEEK(m.sv, 'osoerz') AND osoerz.kl == 'y', .T., .F.)
  IF m.IsGood=.f. AND m.IsVs AND USED('kms')
   m.vvs = SUBSTR(m.sn_pol,7,9)
   IF SEEK(m.vvs, 'kms')
    m.IsGood = .t.
    REPLACE qq WITH m.qcod, sv WITH '110'
   ENDIF 
  ENDIF 
  
  IF IsGood == .f.
   IF !SEEK(m.RecId, 'rError')
    INSERT INTO rError (f, c_err, rid) VALUES ('R', 'ERA', m.RecId)
    = SEEK(m.sn_pol, 'Talon')
    DO WHILE talon.sn_pol = m.sn_pol
     m.t_recid = Talon.RecId
     IF !SEEK(m.t_recid, 'sError')
      INSERT INTO sError (f, c_err, rid) VALUES ('S', 'PKA', m.t_recid)
     ENDIF 
     SKIP IN Talon
    ENDDO 
   ENDIF 
   LOOP 
  ENDIF 

  IF qq != m.qcod
   IF !SEEK(m.RecId, 'rError')
    INSERT INTO rError (f, c_err, rid) VALUES ('R', 'ECA', m.RecId)
    = SEEK(m.sn_pol, 'Talon')
    DO WHILE talon.sn_pol = m.sn_pol
     m.t_recid = Talon.RecId
     IF !SEEK(m.t_recid, 'sError')
      INSERT INTO sError (f, c_err, rid) VALUES ('S', 'PKA', m.t_recid)
     ENDIF 
     SKIP IN Talon
    ENDDO 
   ENDIF 
   LOOP 
  ENDIF 

 ENDSCAN 
 
 SET RELATION OFF INTO Answer
 USE 
 USE IN rError 
 USE IN sError
 SELECT Answer
 SET ORDER TO 
 DELETE TAG ALL 
 USE
 USE IN talon 
 USE IN OsoERZ
* IF USED('kms')
*  USE IN kms
* ENDIF 
 m.t_1 = SECONDS()
 SELECT AisOms
 REPLACE erz_status WITH 2, t_3 WITH m.t_1-m.t_0

RETURN .T. 