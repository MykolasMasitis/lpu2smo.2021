FUNCTION chkBase

IsArcExists = fso.FolderExists(pArc)
IF IsArcExists == .F.
 IF MESSAGEBOX("нрясрярбсер дхпейрнпхъ" + CHR(13) + "&pArc!" + CHR(13) + "янгдюрэ?",4+32, "") == 6
  fso.CreateFolder(pArc)
 ENDIF 
ENDIF 

IsDirExists = fso.FolderExists(pBase)
IF IsDirExists == .F.
 IF MESSAGEBOX("нрясрярбсер дхпейрнпхъ" + CHR(13) + "&pBase!" + CHR(13) + "янгдюрэ?",4+32, "") == 6
  fso.CreateFolder(pBase)
 ENDIF 
ENDIF 

IsDirExists = fso.FolderExists(pBase+'\'+gcPeriod)
IF IsDirExists == .F.
 *_screen.Picture = pbin+'\forms\lpu2smo.scx'
 IF MESSAGEBOX("нрясрярбсер дхпейрнпхъ" + CHR(13) + "&pBase\&gcPeriod!" + CHR(13) + "янгдюрэ?",4+32, "") == 6
  fso.CreateFolder(pBase+'\'+gcPeriod)
 ENDIF 
ENDIF 

*IsDirExists = fso.FolderExists(pDouble)
*IF IsDirExists == .F.
* fso.CreateFolder(pDouble)
*ENDIF 

IsDirExists = fso.FolderExists(pMee)
IF IsDirExists == .F.
 fso.CreateFolder(pMee)
ENDIF 

IsDirExists = fso.FolderExists(pMee+'\SVACTS')
IF IsDirExists == .F.
 fso.CreateFolder(pMee+'\SVACTS')
ENDIF 

IsDirExists = fso.FolderExists(pMee+'\SSACTS')
IF IsDirExists == .F.
 fso.CreateFolder(pMee+'\SSACTS')
ENDIF 

IF IsUsrDir=.T. AND UPPER(LEFT(m.gcUser,3))='EXP'
 m.usrdir = fso.GetParentFolderName(pbin) + '\'+UPPER(m.gcuser)
 IF !fso.FolderExists(m.usrdir)
  fso.CreateFolder(m.usrdir)
 ENDIF 
 IF !fso.FolderExists(m.usrdir+'\SVACTS')
  fso.CreateFolder(m.usrdir+'\SVACTS')
 ENDIF 
 IF !fso.FolderExists(m.usrdir+'\SSACTS')
  fso.CreateFolder(m.usrdir+'\SSACTS')
 ENDIF 
ENDIF 

DaemonDir = fso.GetParentFolderName(pbase) + '\DAEMON'
SpamDir   = fso.GetParentFolderName(pbase) + '\SPAM'
SoapDir   = fso.GetParentFolderName(pbin) + '\SOAP'

*IF !fso.FolderExists(DaemonDir)
* fso.CreateFolder(DaemonDir)
*ENDIF 

*IF !fso.FolderExists(SpamDir)
* fso.CreateFolder(SpamDir)
*ENDIF 

IF !fso.FolderExists(SoapDir)
 fso.CreateFolder(SoapDir)
ENDIF 

m.pSoap = SoapDir

IsDirExists = fso.FolderExists(pMee+'\REQUESTS')
IF IsDirExists == .F.
 fso.CreateFolder(pMee+'\REQUESTS')
ENDIF 

IsDirExists = fso.FolderExists(pMee+'\TUNES')
IF IsDirExists == .F.
 fso.CreateFolder(pMee+'\TUNES')
ENDIF 

IsDirExists = fso.FolderExists(pMee+'\BLANK')
IF IsDirExists == .F.
 fso.CreateFolder(pMee+'\BLANK')
ENDIF 

IsDirExists = fso.FolderExists(pOut)
IF IsDirExists == .F.
 IF MESSAGEBOX("нрясрярбсер дхпейрнпхъ" + CHR(13) + "&pOut!" + CHR(13) + "янгдюрэ?",4+32, "") == 6
  fso.CreateFolder(pOut)
 ENDIF 
ENDIF 

IsDirExists = fso.FolderExists(pTempl)
IF IsDirExists == .F.
 IF MESSAGEBOX("нрясрярбсер дхпейрнпхъ" + CHR(13) + "&pTempl!" + CHR(13) + "янгдюрэ?",4+32, "") == 6
  fso.CreateFolder(pTempl)
 ENDIF 
ENDIF 

*IsDirExists = fso.FolderExists(pTrash)
*IF IsDirExists == .F.
* fso.CreateFolder(pTrash)
*ENDIF 

IsDirExists = fso.FolderExists(pExpImp)
IF IsDirExists == .F.
 fso.CreateFolder(pExpImp)
ENDIF 

IF !fso.FileExists(pcommon+'\dspcodes.dbf')
 MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер тюик '+CHR(13)+CHR(10)+pcommon+'\DSPCODES.DBF'+CHR(13)+CHR(10),0+64,'')
ENDIF 

*IF !fso.FileExists(pcommon+'\mo_vmp_2014.dbf')
* MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер тюик '+CHR(13)+CHR(10)+pcommon+'\MO_VMP_2014.DBF'+CHR(13)+CHR(10),0+64,'')
*ENDIF 

*IF !fso.FileExists(pcommon+'\spi_lpu_dd_2014.dbf')
* MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер тюик '+CHR(13)+CHR(10)+pcommon+'\SPI_LPU_DD_2014.DBF'+CHR(13)+CHR(10),0+64,'')
*ENDIF 

*IF !fso.FileExists(pcommon+'\tpn.dbf')
* MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер тюик '+CHR(13)+CHR(10)+pcommon+'\TPN.DBF'+CHR(13)+CHR(10),0+64,'')
*ENDIF 

IsFileExists = fso.FileExists(pCommon+'\UsrLpu.dbf')
IF IsFileExists == .F.
 CREATE TABLE &pCommon\UsrLpu (mcod c(7), lpu_id n(4), cokr c(2), usr n(2)) 
 APPEND FROM &pCommon\sprlpuxx 
 REPLACE ALL usr WITH 1
* INDEX ON mcod TAG mcod 
* INDEX ON lpu_id TAG lpu_id
 USE 
ENDIF 

IsFileExists = fso.FileExists(pCommon+'\Users.dbf')
IF IsFileExists == .F.
 CREATE TABLE &pCommon\Users (RecId i AUTOINC NEXTVALUE 1 STEP 1, "name" c(6), ;
  fam c(25), im c(25), ot c(25), fio c(40), super l, usrmail c(6)) 
 INDEX ON recid TAG recid CANDIDATE 
 INDEX on name TAG name 
 
 INSERT INTO Users ("name",fam,im,ot,fio,super,usrmail) VALUES ;
  ('OMS','пЪАНБ','лХУЮХК','яРЮМХЯКЮБНБХВ','пЪАНБ л.я.',.t.,'USR010')
 USE
ENDIF 

IF fso.FileExists(pMee+'\SVACTS\svacts.dbf')
 IF OpenFile(pMee+'\SVACTS\svacts', 'svacts', 'shar')>0
  IF USED('svacts')
   USE IN svacts
  ENDIF 
 ELSE 
  SELECT svacts
  IF FIELD('flcod')!='FLCOD'
   USE IN svacts
   IF OpenFile(pMee+'\SVACTS\svacts', 'svacts', 'excl')>0
   ELSE 
    SELECT svacts
    ALTER TABLE svacts ADD COLUMN flcod c(12)
   ENDIF 
  ENDIF 
  IF USED('svacts')
   USE IN svacts
  ENDIF 
 ENDIF 
ENDIF 

IF fso.FileExists(pMee+'\ssacts\ssacts.dbf')
 IF OpenFile(pMee+'\ssacts\ssacts', 'ssacts', 'shar')>0
  IF USED('ssacts')
   USE IN ssacts
  ENDIF 
 ELSE 
  SELECT ssacts
  IF FIELD('flcod')!='FLCOD'
   USE IN ssacts
   IF OpenFile(pMee+'\ssacts\ssacts', 'ssacts', 'excl')>0
   ELSE 
    SELECT ssacts
    ALTER TABLE ssacts ADD COLUMN flcod c(12)
   ENDIF 
  ENDIF 
  IF USED('ssacts')
   USE IN ssacts
  ENDIF 
 ENDIF 
ENDIF 

IsDirExists = fso.FolderExists(pBase+'\'+gcPeriod+'\NSI')
IF IsDirExists == .F.
 pUpdDir = fso.GetParentFolderName(pbin)+'\UPDATE'
 IF !fso.FolderExists(pUpdDir)
  MESSAGEBOX('дхпейрнпхъ намнбкемхъ мях '+pUpdDir+CHR(13)+CHR(10)+'ме намюпсфемю! намнбкемхе мях мебнглфмн!',0+48,'')
 ELSE 
  oMailDir = fso.GetFolder(pUpdDir)
  MailDirName = oMailDir.Path
  oFilesInMailDir = oMailDir.Files
  nFilesInMailDir = oFilesInMailDir.Count
  m.LastNsi = ''
  m.fpath   = ''
  m.NsiVer  = 'aa'

  FOR EACH oFileInMailDir IN oFilesInMailDir
   m.BFullName = oFileInMailDir.Path
   m.bname     = oFileInMailDir.Name
   m.recieved  = oFileInMailDir.DateLastModified

   IF LOWER(oFileInMailDir.Name) = 'sprspr'
    IF SUBSTR(oFileInMailDir.Name,7,2) > m.NsiVer
     m.LastNsi = oFileInMailDir.Name
     m.fpath   = m.BFullName
    ENDIF 
   ENDIF 

  ENDFOR 
  
  IF EMPTY(m.LastNsi)
   MESSAGEBOX('тюик-яопюбнвмхй '+UPPER('sprsprxx.dbf')+' ме намюпсфем'+CHR(13)+CHR(10)+;
   	'б дхпейрнпхх '+pUpdDir+CHR(13)+CHR(10)+'намнбкемхе мях мебнглфмн!',0+48,'')
  ELSE 
   IF MESSAGEBOX('яюлши юйрсюкэмши хг намюпсфеммшу яопюбнвмхйнб:'+CHR(13)+CHR(10)+;
   	UPPER(m.LastNsi) + '. хяонкэгнбюрэ ецн?',4+32,'')=6 && хЯОНКЭГНБЮРЭ ОНЯКЕДМХИ НАМЮПСФЕММШИ!
   ELSE 
    SET DEFAULT TO (pUpdDir)
    csprfile = ''
    csprfile=GETFILE('dbf')
    IF EMPTY(csprfile)
     m.LastNsi = ''
     m.fpath   = ''
     m.NsiVer  = 'aa'
     MESSAGEBOX(CHR(13)+CHR(10)+'бш мхвецн ме бшапюкх!'+CHR(13)+CHR(10),0+16,'')
    ELSE 
     ospr = fso.GetFile(csprfile)
     IF LOWER(LEFT(ospr.name,6)) != 'sprspr'
      m.LastNsi = ''
      m.fpath   = ''
      m.NsiVer  = 'aa'
      MESSAGEBOX(CHR(13)+CHR(10)+'щрн ме яопюбнвмхй мях!'+CHR(13)+CHR(10),0+16,'sprsprxx')
      RELEASE ospr 
     ELSE 
      m.fpath = csprfile
     ENDIF 
    ENDIF 
   ENDIF 
  ENDIF 
  
  IF fso.FileExists(m.fpath)
*   MESSAGEBOX(UPPER(m.fpath),0+64,'')
  ENDIF 
  
  IF OpenFile(m.fpath, 'sprspr', 'shar')=0
   SELECT sprspr
   LOCATE FOR name_eta='sprspr'
   m.in_data={}
   IF FIELD('intr_data')=UPPER('intr_data')
    IF FOUND()
     m.in_data = intr_data
    ENDIF 
   ENDIF 
   USE IN sprspr
   
   IF !EMPTY(m.in_data)
    IF m.in_data != m.tdat1
     IF MESSAGEBOX('дюрю ббндю бшапюммни мях '+DTOC(m.in_data)+CHR(13)+CHR(10)+;
     	'ме яннрберярбсер мювюкс нрвермнцн оепхндю '+DTOC(m.tdat1)+CHR(13)+CHR(10)+;
     	'бяе пюбмн опнднкфхрэ?',4+32,'')=6
      WAIT "юйрсюкхгюжхъ мях..." WINDOW NOWAIT 
      DO ActualizeNSI WITH m.fpath && ОПНДНКФХРЭ
      WAIT CLEAR 
     ELSE 
      && МЕР!
     ENDIF 
    ELSE
	 MESSAGEBOX('дюрю ббндю бшапюммни мях '+DTOC(m.in_data)+CHR(13)+CHR(10)+;
     	'яннрберярбсер мювюкс нрвермнцн оепхндю '+DTOC(m.tdat1),0+64,'')
     WAIT "юйрсюкхгюжхъ мях..." WINDOW NOWAIT 
     DO ActualizeNSI WITH m.fpath
     WAIT CLEAR 
    ENDIF 
   ELSE 
    MESSAGEBOX('б бшапюммнл тюике ме намюпсфемю дюрю ббндю!',0+48,'')
   ENDIF 
  ENDIF 
  
 ENDIF 
 
 IF MESSAGEBOX("нрясрярбсер дхпейрнпхъ" + CHR(13) + "&pBase\&gcPeriod\NSI!" + CHR(13) + "янгдюрэ?",4+32, "") == 6
  WAIT "янгдюеряъ кнйюкэмши мях..." WINDOW NOWAIT 

 fso.CreateFolder(pBase+'\'+gcPeriod+'\NSI')

tyu = pcommon+'\admokrxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\admokrxx.dbf', pBase+'\'+gcPeriod+'\NSI\admokrxx.dbf')
  
tyu = pcommon+'\codku_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\codku_xx.dbf', pBase+'\'+gcPeriod+'\NSI\codku.dbf')

tyu = pcommon+'\codotdxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\codotdxx.dbf', pBase+'\'+gcPeriod+'\NSI\codotd.dbf')

tyu = pcommon+'\codwdrxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\codwdrxx.dbf', pBase+'\'+gcPeriod+'\NSI\codwdr.dbf')

tyu = pcommon+'\hopff_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\hopff_xx.dbf', pBase+'\'+gcPeriod+'\NSI\hopff.dbf')
IF fso.FileExists(pcommon+'\hopff_xx.cdx')
 fso.CopyFile(pcommon+'\hopff_xx.cdx', pBase+'\'+gcPeriod+'\NSI\hopff.cdx')
ENDIF 

tyu = pcommon+'\isv012xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\isv012xx.dbf', pBase+'\'+gcPeriod+'\NSI\isv012.dbf')

tyu = pcommon+'\kdolgxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\kdolgxx.dbf', pBase+'\'+gcPeriod+'\NSI\kdolgxx.dbf')

tyu = pcommon+'\kpreslxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\kpreslxx.dbf', pBase+'\'+gcPeriod+'\NSI\kpresl.dbf')

*tyu = pcommon+'\kspecxx.dbf'
*oSettings.CodePage('&tyu', 866, .t.)
*fso.CopyFile(pcommon+'\kspecxx.dbf', pBase+'\'+gcPeriod+'\NSI\kspec.dbf')

tyu = pcommon+'\mkb10_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\mkb10_xx.dbf', pBase+'\'+gcPeriod+'\NSI\mkb10.dbf')
IF fso.FileExists(pBase+'\'+gcPeriod+'\NSI\mkb10.dbf')
 IF OpenFile(pBase+'\'+gcPeriod+'\NSI\mkb10', 'mkb', 'excl')>0
  IF USED('mkb')
   USE IN mkb
  ENDIF 
 ELSE 
  SELECT mkb 
  IF FIELD('opl')='OPL'
   ALTER TABLE mkb ADD COLUMN IsOMS l 
   REPLACE ALL IsOms WITH IIF(opl='1', .t., .f.) 
  ENDIF 
  USE 
  IF fso.FileExists(pBase+'\'+gcPeriod+'\NSI\mkb10.bak')
   fso.DeleteFile(pBase+'\'+gcPeriod+'\NSI\mkb10.bak')
  ENDIF 
  USE 
 ENDIF 
ENDIF 

*tyu = pcommon+'\mo_vmp_2014.dbf'
*oSettings.CodePage('&tyu', 866, .t.)
*fso.CopyFile(pcommon+'\mo_vmp_2014.dbf', pBase+'\'+gcPeriod+'\NSI\mo_vmp.dbf')

tyu = pcommon+'\modpacxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\modpacxx.dbf', pBase+'\'+gcPeriod+'\NSI\modpac.dbf')

tyu = pcommon+'\ms_mkbxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\ms_mkbxx.dbf', pBase+'\'+gcPeriod+'\NSI\ms_mkb.dbf')

tyu = pcommon+'\nocodrxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\nocodrxx.dbf', pBase+'\'+gcPeriod+'\NSI\nocodr.dbf')

tyu = pcommon+'\osoerzxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\osoerzxx.dbf', pBase+'\'+gcPeriod+'\NSI\osoerzxx.dbf')

tyu = pcommon+'\osoreexx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\osoreexx.dbf', pBase+'\'+gcPeriod+'\NSI\osoree.dbf')

tyu = pcommon+'\ososchxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\ososchxx.dbf', pBase+'\'+gcPeriod+'\NSI\ososch.dbf')

tyu = pcommon+'\profotxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\profotxx.dbf', pBase+'\'+gcPeriod+'\NSI\profot.dbf')

tyu = pcommon+'\profusxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\profusxx.dbf', pBase+'\'+gcPeriod+'\NSI\profus.dbf')

tyu = pcommon+'\codprvxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\codprvxx.dbf', pBase+'\'+gcPeriod+'\NSI\codprv.dbf')

* нМЙНКНЦХЪ
tyu = pcommon+'\onreasxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onreasxx.dbf', pBase+'\'+gcPeriod+'\NSI\onreasxx.dbf')

tyu = pcommon+'\onstadxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onstadxx.dbf', pBase+'\'+gcPeriod+'\NSI\onstadxx.dbf')

tyu = pcommon+'\ontum_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\ontum_xx.dbf', pBase+'\'+gcPeriod+'\NSI\ontum_xx.dbf')

tyu = pcommon+'\onnod_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onnod_xx.dbf', pBase+'\'+gcPeriod+'\NSI\onnod_xx.dbf')

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\onnod_xx', "onnod", "excl") == 0
 SELECT onnod
 ALTER TABLE onnod ALTER COLUMN cod_n n(4)
 USE
ENDIF

tyu = pcommon+'\onmet_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onmet_xx.dbf', pBase+'\'+gcPeriod+'\NSI\onmet_xx.dbf')

tyu = pcommon+'\onlechxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onlechxx.dbf', pBase+'\'+gcPeriod+'\NSI\onlechxx.dbf')

tyu = pcommon+'\onhir_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onhir_xx.dbf', pBase+'\'+gcPeriod+'\NSI\onhir_xx.dbf')

tyu = pcommon+'\onleklxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onleklxx.dbf', pBase+'\'+gcPeriod+'\NSI\onleklxx.dbf')

tyu = pcommon+'\onlekvxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onlekvxx.dbf', pBase+'\'+gcPeriod+'\NSI\onlekvxx.dbf')

tyu = pcommon+'\onluchxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onluchxx.dbf', pBase+'\'+gcPeriod+'\NSI\onluchxx.dbf')

tyu = pcommon+'\onprotxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onprotxx.dbf', pBase+'\'+gcPeriod+'\NSI\onprotxx.dbf')

tyu = pcommon+'\onmrf_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onmrf_xx.dbf', pBase+'\'+gcPeriod+'\NSI\onmrf_xx.dbf')

tyu = pcommon+'\onmrdsxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onmrdsxx.dbf', pBase+'\'+gcPeriod+'\NSI\onmrdsxx.dbf')

tyu = pcommon+'\onigh_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onigh_xx.dbf', pBase+'\'+gcPeriod+'\NSI\onigh_xx.dbf')

tyu = pcommon+'\onigdsxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onigdsxx.dbf', pBase+'\'+gcPeriod+'\NSI\onigdsxx.dbf')

tyu = pcommon+'\onmrfrxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onmrfrxx.dbf', pBase+'\'+gcPeriod+'\NSI\onmrfrxx.dbf')

tyu = pcommon+'\onigrtxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onigrtxx.dbf', pBase+'\'+gcPeriod+'\NSI\onigrtxx.dbf')

tyu = pcommon+'\onconsxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onconsxx.dbf', pBase+'\'+gcPeriod+'\NSI\onconsxx.dbf')

tyu = pcommon+'\onpcelxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onpcelxx.dbf', pBase+'\'+gcPeriod+'\NSI\onpcelxx.dbf')

tyu = pcommon+'\onnaprxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onnaprxx.dbf', pBase+'\'+gcPeriod+'\NSI\onnaprxx.dbf')

tyu = pcommon+'\onczabxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onczabxx.dbf', pBase+'\'+gcPeriod+'\NSI\onczabxx.dbf')

tyu = pcommon+'\ondopkxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\ondopkxx.dbf', pBase+'\'+gcPeriod+'\NSI\ondopkxx.dbf')

tyu = pcommon+'\onoplsxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onoplsxx.dbf', pBase+'\'+gcPeriod+'\NSI\onoplsxx.dbf')

tyu = pcommon+'\onlpshxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\onlpshxx.dbf', pBase+'\'+gcPeriod+'\NSI\onlpshxx.dbf')

tyu = pcommon+'\msextxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\msextxx.dbf', pBase+'\'+gcPeriod+'\NSI\msext.dbf')

tyu = pcommon+'\msextxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\msextxx.dbf', pBase+'\'+gcPeriod+'\NSI\msext.dbf')

tyu = pcommon+'\sprncoxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\sprncoxx.dbf', pBase+'\'+gcPeriod+'\NSI\sprnco.dbf')

tyu = pcommon+'\tarionxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\tarionxx.dbf', pBase+'\'+gcPeriod+'\NSI\tarion.dbf')

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi\tarion', "tarion", "excl") == 0
 SELECT tarion
 ALTER TABLE tarion ALTER COLUMN cod c(8)
 SET FULLPATH OFF 
 WAIT "хмдейяхпнбюмхе тюикю "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX on cod TAG cod
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

* нМЙНКНЦХЪ

tyu = pcommon+'\prv002xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
IF fso.FileExists(pCommon+'\prv002xx.dbf')
 IF OpenFile(pCommon+'\prv002xx', 'profss', 'shar')>0
  IF USED('profss')
   USE IN profss
  ENDIF 
 ELSE 
  SELECT profss
  IF FIELD('isoms')!='ISOMS'
   USE IN profss
   IF OpenFile(pCommon+'\prv002xx', 'profss', 'excl')>0
   ELSE 
    SELECT profss
    ALTER table profss ADD COLUMN isoms l
    REPLACE ALL isoms WITH .t.
   ENDIF 
  ENDIF 
  IF USED('profss')
   USE IN profss
  ENDIF 
 ENDIF 
ENDIF 

tyu = pcommon+'\rsv009xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\rsv009xx.dbf', pBase+'\'+gcPeriod+'\NSI\rsv009.dbf')

tyu = pcommon+'\reeskpxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\reeskpxx.dbf', pBase+'\'+gcPeriod+'\NSI\reeskp.dbf')

tyu = pcommon+'\sookodxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\sookodxx.dbf', pBase+'\'+gcPeriod+'\NSI\sookodxx.dbf')

tyu = pcommon+'\sovmnoxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\sovmnoxx.dbf', pBase+'\'+gcPeriod+'\NSI\sovmno.dbf')

*tyu = pcommon+'\spi_lpu_dd_2014.dbf'
*oSettings.CodePage('&tyu', 866, .t.)
*fso.CopyFile(pcommon+'\spi_lpu_dd_2014.dbf', pBase+'\'+gcPeriod+'\NSI\spi_lpu_dd.dbf')

tyu = pcommon+'\spr_ulxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
IF OpenFile(pcommon+'\spr_ulxx', 'sprul', 'shar')<=0
 CREATE TABLE &pBase\&gcPeriod\NSI\street (ul n(5), street c(60), cokr c(2), mo c(4))
 INDEX ON ul TAG ul
 USE 
 =OpenFile(pBase+'\'+gcPeriod+'\NSI\street', 'street', 'excl')<=0
 
 SELECT sprul
 SCAN 
  m.priznak = priznak
  IF m.priznak != 'a'
   LOOP 
  ENDIF 
  m.ul = INT(VAL(kod_fo))
  m.street = ALLTRIM(nmstreet)
    
  INSERT INTO street FROM MEMVAR 
    
 ENDSCAN 
 USE 
 SELECT street
 SORT ON street TO &pBase\&gcPeriod\NSI\qwert
 ZAP
 APPEND FROM &pBase\&gcPeriod\NSI\qwert
 IF fso.FileExists(pBase+'\'+gcPeriod+'\NSI\qwert.dbf')
  fso.DeleteFile(pBase+'\'+gcPeriod+'\NSI\qwert.dbf')
 ENDIF 
 USE IN street 
ENDIF 

tyu = pcommon+'\sprlpuxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
  
IF OpenFile(pcommon+'\sprlpuxx', 'sprlpu', 'shar')<=0
 SELECT sprlpu
 COPY FOR lpu_id=fil_id TO pBase+'\'+gcPeriod+'\NSI\sprlpuxx' ;
  FIELDS lpu_id,fil_id,fcod,prn_kodved,tpn,tpns,vmp,mcod,name,fullname,cokr,adres
 COPY FOR name='оНКХЙКХМХЙЮ Я ром' TO pBase+'\'+gcPeriod+'\NSI\lputpn' ;
  FIELDS lpu_id,fil_id,fcod,prn_kodved,tpn,tpns,vmp,mcod,name,fullname,cokr,adres
 COPY FOR tpn='4' TO pBase+'\'+gcPeriod+'\NSI\horlpu' ;
  FIELDS lpu_id,fil_id,fcod,prn_kodved,tpn,tpns,vmp,mcod,name,fullname,cokr,adres
 COPY FOR tpns='4' TO pBase+'\'+gcPeriod+'\NSI\horlpus' ;
  FIELDS lpu_id,fil_id,fcod,prn_kodved,tpn,tpns,vmp,mcod,name,fullname,cokr,adres
 COPY FOR fil_id=lpu_id AND INLIST(tpn,'1','3') TO pBase+'\'+gcPeriod+'\NSI\pilot' ;
  FIELDS lpu_id,mcod
 COPY FOR fil_id=lpu_id AND INLIST(tpns,'1','3') TO pBase+'\'+gcPeriod+'\NSI\pilots' ;
  FIELDS lpu_id,mcod
 USE 
 IF OpenFile(pBase+'\'+gcPeriod+'\NSI\sprlpuxx', 'sprlpu', 'excl')<=0
  SET FULLPATH OFF 
  WAIT "хмдейяхпнбюмхе тюикю "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
  INDEX ON lpu_id TAG lpu_id
  INDEX ON fil_id TAG fil_id
  INDEX ON mcod TAG mcod
  INDEX ON cokr TAG cokr
  USE 
  SET FULLPATH OFF 
 ENDIF 
 IF OpenFile(pBase+'\'+gcPeriod+'\NSI\lputpn', 'lputpn', 'excl')<=0
  SET FULLPATH OFF 
  WAIT "хмдейяхпнбюмхе тюикю "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
  DELETE FOR INLIST(lpu_id,1891,4475)
  INDEX ON lpu_id TAG lpu_id
  INDEX ON fil_id TAG fil_id
  INDEX ON mcod TAG mcod
  USE 
  SET FULLPATH OFF 
 ENDIF 
ENDIF 

IF OpenFile(pCommon+'\sprvedxx', 'sprved', 'shar')>0
 IF USED('sprved')
  USE IN sprved
 ENDIF 
ELSE 
 IF OpenFile(pBase+'\'+gcPeriod+'\NSI\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  USE IN sprved
 ELSE 
  CREATE CURSOR c_ved (lpu_id n(4), mcod c(7), prn_kodved n(5))
  SELECT c_ved 
  INDEX on lpu_id TAG lpu_id
  INDEX on mcod TAG mcod 
  SET ORDER TO lpu_id
  
  SELECT sprved
  SCAN 
   m.lpu_id = lpu_id 
   m.mcod   = IIF(SEEK(m.lpu_id, 'sprlpu'), sprlpu.mcod, '')
   m.prn_kodved = IIF(SEEK(m.lpu_id, 'sprlpu'), sprlpu.prn_kodved, 0)
   IF !SEEK(m.lpu_id, 'c_ved') AND !EMPTY(m.mcod)
    INSERT INTO c_ved FROM MEMVAR 
   ENDIF 
  ENDSCAN 
  USE IN sprved
  
  SELECT c_ved 
  SET ORDER TO mcod 
  COPY TO &pBase\&gcPeriod\NSI\sprved WITH cdx 
  USE 
  USE IN sprlpu

 ENDIF 
ENDIF 

tyu = pcommon+'\spraboxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)

IF OpenFile(pcommon+'\spraboxx', 'sprabo', 'shar')<=0
 IF OpenFile(pBase+'\'+gcPeriod+'\NSI\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')<=0
  SELECT sprabo
  SET RELATION TO object_id INTO sprlpu
  COPY FOR !EMPTY(sprlpu.lpu_id) AND abn_type='0' fields object_id, abn_name, name;
  TO pBase+'\'+gcPeriod+'\NSI\spraboxx'
  SET RELATION OFF INTO sprlpu
  USE 
  USE IN sprlpu
 ELSE 
  USE IN sprlpu
 ENDIF 
ENDIF 

tyu = pcommon+'\spv015xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\spv015xx.dbf', pBase+'\'+gcPeriod+'\NSI\spv015.dbf')

tyu = pcommon+'\tipgrpxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\tipgrpxx.dbf', pBase+'\'+gcPeriod+'\NSI\tipgrp.dbf')

tyu = pcommon+'\tipno_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\tipno_xx.dbf', pBase+'\'+gcPeriod+'\NSI\tipnomes.dbf')

tyu = pcommon+'\vidvp_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\vidvp_xx.dbf', pBase+'\'+gcPeriod+'\NSI\vidvp.dbf')

tyu = pcommon+'\z_cod_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\z_cod_xx.dbf', pBase+'\'+gcPeriod+'\NSI\z_cod.dbf')

tyu = pcommon+'\z_dsnoxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\z_dsnoxx.dbf', pBase+'\'+gcPeriod+'\NSI\z_dsnoxx.dbf')

*fso.CopyFile(pcommon+'\polic_dp.dbf', pBase+'\'+gcPeriod+'\NSI\polic_dp.dbf')
*fso.CopyFile(pcommon+'\polic_h.dbf', pBase+'\'+gcPeriod+'\NSI\polic_h.dbf')
*fso.CopyFile(pcommon+'\spi_cz.dbf', pBase+'\'+gcPeriod+'\NSI\spi_cz.dbf')
*fso.CopyFile(pcommon+'\spi_cz_ch.dbf', pBase+'\'+gcPeriod+'\NSI\spi_cz_ch.dbf')
fso.CopyFile(pcommon+'\usrlpu.dbf', pBase+'\'+gcPeriod+'\NSI\usrlpu.dbf')
IF fso.FileExists(pcommon+'\usrlpu.cdx')
 fso.CopyFile(pcommon+'\usrlpu.cdx', pBase+'\'+gcPeriod+'\NSI\usrlpu.cdx')
ENDIF 

tyu = pcommon+'\errsmee.dbf'
fso.CopyFile(pcommon+'\errsmee.dbf', pBase+'\'+gcPeriod+'\NSI\errsmee.dbf')
tyu = pcommon+'\errsmee.cdx'
fso.CopyFile(pcommon+'\errsmee.cdx', pBase+'\'+gcPeriod+'\NSI\errsmee.cdx')

m.pr_period = IIF(m.tMonth>1, STR(m.tYear,4)+PADL(m.tMonth-1,2,'0'), STR(m.tYear-1,4)+'12')
IF fso.FolderExists(pBase+'\'+m.pr_period)

 IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\noth.dbf')
  fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\noth.dbf', pBase+'\'+m.gcperiod+'\NSI\noth.dbf')
  IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\noth.cdx')
   fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\noth.cdx', pBase+'\'+m.gcperiod+'\NSI\noth.cdx')
  ENDIF 
 ENDIF 

 IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\kms.dbf')
  fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\kms.dbf', pBase+'\'+m.gcperiod+'\NSI\kms.dbf')
  IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\kms.cdx')
   fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\kms.cdx', pBase+'\'+m.gcperiod+'\NSI\kms.cdx')
  ENDIF 
 ENDIF 

 IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\stpilot.dbf')
  fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\stpilot.dbf', pBase+'\'+m.gcperiod+'\NSI\stpilot.dbf')
  IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\stpilot.cdx')
   fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\stpilot.cdx', pBase+'\'+m.gcperiod+'\NSI\stpilot.cdx')
  ENDIF 
 ENDIF 

 IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\nsio.dbf')
  fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\nsio.dbf', pBase+'\'+m.gcperiod+'\NSI\nsio.dbf')
  IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\nsio.cdx')
   fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\nsio.cdx', pBase+'\'+m.gcperiod+'\NSI\nsio.cdx')
  ENDIF 
 ENDIF 

 IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\nsif.dbf')
  fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\nsif.dbf', pBase+'\'+m.gcperiod+'\NSI\nsif.dbf')
  IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\nsif.cdx')
   fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\nsif.cdx', pBase+'\'+m.gcperiod+'\NSI\nsif.cdx')
  ENDIF 
 ENDIF 

 IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\ns36.dbf')
  fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\ns36.dbf', pBase+'\'+m.gcperiod+'\NSI\ns36.dbf')
  IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\ns36.cdx')
   fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\ns36.cdx', pBase+'\'+m.gcperiod+'\NSI\ns36.cdx')
  ENDIF 
 ENDIF 

 IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\novzms.dbf')
  fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\novzms.dbf', pBase+'\'+m.gcperiod+'\NSI\novzms.dbf')
  IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\novzms.cdx')
   fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\novzms.cdx', pBase+'\'+m.gcperiod+'\NSI\novzms.cdx')
  ENDIF 
 ENDIF 

 IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\f003.dbf')
  fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\f003.dbf', pBase+'\'+m.gcperiod+'\NSI\f003.dbf')
  IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\f003.cdx')
   fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\f003.cdx', pBase+'\'+m.gcperiod+'\NSI\f003.cdx')
  ENDIF 
 ENDIF 

 *IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\outs.dbf')
 * fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\outs.dbf', pBase+'\'+m.gcperiod+'\NSI\outs.dbf')
 * IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\outs.cdx')
 *  fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\outs.cdx', pBase+'\'+m.gcperiod+'\NSI\outs.cdx')
 * ENDIF 
 *ENDIF 

 IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\exclhors.dbf')
  fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\exclhors.dbf', pBase+'\'+m.gcperiod+'\NSI\exclhors.dbf')
  IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\exclhors.cdx')
   fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\exclhors.cdx', pBase+'\'+m.gcperiod+'\NSI\exclhors.cdx')
  ENDIF 
 ENDIF 

 IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\pnorm_iskl.dbf')
  fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\pnorm_iskl.dbf', pBase+'\'+m.gcperiod+'\NSI\pnorm_iskl.dbf')
  IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\pnorm_iskl.cdx')
   fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\pnorm_iskl.cdx', pBase+'\'+m.gcperiod+'\NSI\pnorm_iskl.cdx')
  ENDIF 
 ENDIF 

 IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\polic_h.dbf')
  fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\polic_h.dbf', pBase+'\'+m.gcperiod+'\NSI\polic_h.dbf')
  IF fso.FileExists(pBase+'\'+m.pr_period+'\NSI\polic_h.cdx')
   fso.CopyFile(pBase+'\'+m.pr_period+'\NSI\polic_h.cdx', pBase+'\'+m.gcperiod+'\NSI\polic_h.cdx')
  ENDIF 
 ENDIF 

ENDIF 

IF fso.FileExists(pcommon+'\tarifn.dbf')
 IF OpenFile(pcommon+'\tarifn', 'tarif', 'excl')<=0
  SELECT tarif
  INDEX ON cod TAG cod
  SET ORDER TO cod 

  IF fso.FileExists(pcommon+'\tarimuxx.dbf')
   tyu = pcommon+'\tarimuxx.dbf'
   oSettings.CodePage('&tyu', 866, .t.)

   IF OpenFile(pcommon+'\tarimuxx', 'tarimu', 'excl')<=0
   
    SELECT tarimu 
    INDEX ON cod TAG cod
    SET ORDER TO cod 
    
    m.nIsChngTarif  = 0
    m.nIsChngTarifV = 0
    m.nIsChngStkd   = 0
    m.nIsChngStkdV  = 0
    
    SELECT tarif
    SET RELATION TO cod INTO tarimu
    SCAN 
     IF tarif!=tarimu.tarif
      REPLACE tarif WITH tarimu.tarif
      m.nIsChngTarif  = m.nIsChngTarif + 1
     ENDIF 
     IF FIELD('doplata', 'tarimu')='DOPLATA'
      IF doplata!=tarimu.doplata
       REPLACE doplata WITH tarimu.doplata
       m.nIsChngTarif  = m.nIsChngTarif + 1
      ENDIF 
     ENDIF 
     IF tarif_v!=tarimu.tarif
      REPLACE tarif_v WITH tarimu.tarif
      m.nIsChngTarifV  = m.nIsChngTarifV + 1
     ENDIF 
     IF stkd!=tarimu.stkd
      REPLACE stkd WITH tarimu.stkd
      m.nIsChngstkd  = m.nIsChngstkd + 1
     ENDIF 
     IF stkdv!=tarimu.stkd
      REPLACE stkdv WITH tarimu.stkd
      m.nIsChngstkdv  = m.nIsChngstkdv + 1
     ENDIF 
    ENDSCAN 
    SET RELATION OFF INTO tarimu
    SET ORDER TO 
    DELETE TAG ALL 
    USE 
    
    SELECT tarimu
    SET ORDER TO 
    DELETE TAG ALL 
    USE 

    IF m.nIsChngTarif > 0
*     MESSAGEBOX(CHR(13)+CHR(10)+'намнбкемн '+TRANSFORM(m.nIsChngTarif,'99999')+' гмювемхи TARIF '+CHR(13)+CHR(10)+;
      'б тюике TARIFN!'+CHR(13)+CHR(10),0+64,'')
    ENDIF 
    IF m.nIsChngTarifV > 0
*     MESSAGEBOX(CHR(13)+CHR(10)+'намнбкемн '+TRANSFORM(m.nIsChngTarifV,'99999')+' гмювемхи TARIF_V '+CHR(13)+CHR(10)+;
      'б тюике TARIFN!'+CHR(13)+CHR(10),0+64,'')
    ENDIF 
    IF m.nIsChngStkd > 0
*     MESSAGEBOX(CHR(13)+CHR(10)+'намнбкемн '+TRANSFORM(m.nIsChngStkd,'99999')+' гмювемхи STKD '+CHR(13)+CHR(10)+;
     'б тюике TARIFN!'+CHR(13)+CHR(10),0+64,'')
    ENDIF 
    IF m.nIsChngStkdV > 0
*     MESSAGEBOX(CHR(13)+CHR(10)+'намнбкемн '+TRANSFORM(m.nIsChngStkdV,'99999')+' гмювемхи STKDV '+CHR(13)+CHR(10)+;
      'б тюике TARIFN!'+CHR(13)+CHR(10),0+64,'')
    ENDIF 

   ELSE 
    IF USED('tarimu')
     USE IN tarimu
    ENDIF 
   ENDIF 
  ENDIF 
 ELSE 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
 ENDIF 
ENDIF 

WAIT "намнбкемхе..." WINDOW NOWAIT 
tyu = pcommon+'\tarimuxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
tyu = pcommon+'\reesusxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
tyu = pcommon+'\reesmsxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
tyu = pcommon+'\usvmpxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)

IF fso.FileExists(pcommon+'\tarifn.dbf')
 IF OpenFile(pcommon+'\tarifn', 'tarif', 'excl')<=0
  SELECT tarif
  INDEX on cod TAG cod
  SET ORDER TO cod 
 ENDIF 
ENDIF 
IF fso.FileExists(pcommon+'\tarimuxx.dbf')
 IF OpenFile(pcommon+'\tarimuxx', 'tarimu', 'excl')<=0
  SELECT tarimu 
  INDEX on cod TAG cod
  SET ORDER TO cod 
 ENDIF 
ENDIF 
IF fso.FileExists(pcommon+'\reesusxx.dbf')
 IF OpenFile(pcommon+'\reesusxx', 'reesus', 'excl')<=0
  SELECT reesus
  INDEX on cod TAG cod
  SET ORDER TO cod 
 ENDIF 
ENDIF 
IF fso.FileExists(pcommon+'\reesmsxx.dbf')
 IF OpenFile(pcommon+'\reesmsxx', 'reesms', 'excl')<=0
  SELECT reesms
  INDEX on cod TAG cod
  SET ORDER TO cod 
 ENDIF 
ENDIF 
IF fso.FileExists(pcommon+'\usvmpxx.dbf')
 IF OpenFile(pcommon+'\usvmpxx', 'usvmp', 'excl')<=0
  SELECT usvmp
  INDEX on cod TAG cod
  SET ORDER TO cod 
 ENDIF 
ENDIF 

IF USED('tarif') AND USED('tarimu') AND USED('reesus') AND USED('reesms') AND USED('usvmp')

SELECT reesus
SET RELATION TO cod INTO tarif
COUNT FOR EMPTY(tarif.cod) TO m.nNewInUs
IF m.nNewInUs > 0
* MESSAGEBOX(CHR(13)+CHR(10)+'б пееярпе сяксц (REESUSXX) намюпсфемн '+CHR(13)+CHR(10)+;
  TRANSFORM(m.nNewInUs,'99999')+' мнбшу гюохяеи!'+CHR(13)+CHR(10),0+64,'')
 SCAN 
  IF EMPTY(tarif.cod)
   SCATTER MEMVAR 
   m.comment = m.name 
   INSERT INTO tarif FROM MEMVAR 
  ENDIF 
 ENDSCAN 
ENDIF 
SET RELATION OFF INTO tarif

SELECT tarif
SET RELATION TO cod INTO reesus
COUNT FOR EMPTY(reesus.cod) AND isitus(cod) TO m.nDelInUs

IF m.nDelInUs>0
* MESSAGEBOX(CHR(13)+CHR(10)+'хг пееярпю сяксц (REESUSXX) хяйкчвемн '+CHR(13)+CHR(10)+;
  TRANSFORM(m.nDelInUs,'99999')+' гюохяеи!'+CHR(13)+CHR(10),0+64,'')
 SCAN 
  m.cod = cod
  IF IsItUs(m.cod) AND EMPTY(reesus.cod)
   DELETE 
  ENDIF 
 ENDSCAN 
 SET RELATION OFF INTO reesus
 PACK 
 SET RELATION TO cod INTO reesus
ENDIF 

m.nIsChngTpn  = 0
m.nIsChngUet1 = 0
m.nIsChngUet2 = 0
m.nIsChngName = 0

SCAN 
 m.cod = cod
 *IF !IsItUs(m.cod)
 * LOOP 
 *ENDIF 
 IF EMPTY(reesus.cod)
  LOOP 
 ENDIF 
 
 IF tpn != reesus.tpn
  m.nIsChngTpn  = m.nIsChngTpn + 1
  REPLACE tpn WITH reesus.tpn
 ENDIF 
 IF uet1 != reesus.uet1
  m.nIsChngUet1  = m.nIsChngUet1 + 1
  REPLACE uet1 WITH reesus.uet1
 ENDIF 
 IF uet2 != reesus.uet2
  m.nIsChngUet2  = m.nIsChngUet2 + 1
  REPLACE uet2 WITH reesus.uet2
 ENDIF 
 IF name != reesus.name
  m.nIsChngName  = m.nIsChngName + 1
  REPLACE name WITH reesus.name
  REPLACE comment WITH reesus.name
 ENDIF 
ENDSCAN 
SET RELATION OFF INTO reesus

IF m.nIsChngTpn > 0
* MESSAGEBOX(CHR(13)+CHR(10)+'намнбкемн '+TRANSFORM(m.nIsChngTpn,'99999')+' гмювемхи TPN '+CHR(13)+CHR(10)+;
  'б тюике TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 
IF m.nIsChngUet1 > 0
* MESSAGEBOX(CHR(13)+CHR(10)+'намнбкемн '+TRANSFORM(m.nIsChngUet1,'99999')+' гмювемхи UET1 '+CHR(13)+CHR(10)+;
  'б тюике TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 
IF m.nIsChngUet2 > 0
* MESSAGEBOX(CHR(13)+CHR(10)+'намнбкемн '+TRANSFORM(m.nIsChngUet2,'99999')+' гмювемхи UET2 '+CHR(13)+CHR(10)+;
  'б тюике TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 
IF m.nIsChngName > 0
* MESSAGEBOX(CHR(13)+CHR(10)+'намнбкемн '+TRANSFORM(m.nIsChngName,'99999')+' гмювемхи NAME '+CHR(13)+CHR(10)+;
  'б тюике TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 

SELECT reesms
SET RELATION TO cod INTO tarif
COUNT FOR EMPTY(tarif.cod) TO m.nNewInMs
IF m.nNewInMs > 0
* MESSAGEBOX(CHR(13)+CHR(10)+'б пееярпе лщянб (REESMSXX) намюпсфемн '+CHR(13)+CHR(10)+;
  TRANSFORM(m.nNewInMs,'99999')+' мнбшу гюохяеи!'+CHR(13)+CHR(10),0+64,'')
 SCAN 
  IF EMPTY(tarif.cod)
   SCATTER MEMVAR 
   m.name    = m.namem
   m.comment = m.namem
   INSERT INTO tarif FROM MEMVAR 
  ENDIF 
 ENDSCAN 
ENDIF 
SET RELATION OFF INTO tarif

SELECT tarif
SET RELATION TO cod INTO reesms
COUNT FOR EMPTY(reesms.cod) AND !isitus(cod) TO m.nDelInMs

IF m.nDelInMs > 0
* MESSAGEBOX(CHR(13)+CHR(10)+'хг пееярпю лщяНБ (REESMSXX) хяйкчвемн '+CHR(13)+CHR(10)+;
  TRANSFORM(m.nDelInMs,'99999')+' гюохяеи!'+CHR(13)+CHR(10),0+64,'')
 SCAN 
  m.cod = cod
  IF !IsItUs(m.cod) AND EMPTY(reesms.cod)
   DELETE 
  ENDIF 
 ENDSCAN 
 SET RELATION OFF INTO reesms
 PACK 
 SET RELATION TO cod INTO reesms
ENDIF 

m.nIsChngNKD  = 0
m.nIsChngSTKD = 0
m.nIsChngName = 0

SCAN 
 m.cod  = cod

 IF IsItUs(m.cod)
  LOOP 
 ENDIF 
 
 REPLACE tpn WITH '' 
 IF LEFT(name,200) != reesms.namem
  m.nIsChngName  = m.nIsChngName + 1
  REPLACE name WITH reesms.namem
  REPLACE comment WITH reesms.namem
 ENDIF 

ENDSCAN 
SET RELATION OFF INTO reesms

IF m.nIsChngNKD > 0
* MESSAGEBOX(CHR(13)+CHR(10)+'намнбкемн '+TRANSFORM(m.nIsChngNKD,'99999')+' гмювемхи N_KD '+CHR(13)+CHR(10)+;
  'б тюике TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 
IF m.nIsChngName > 0
* MESSAGEBOX(CHR(13)+CHR(10)+'намнбкемн '+TRANSFORM(m.nIsChngName,'99999')+' гмювемхи NAME '+CHR(13)+CHR(10)+;
  'б тюике TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 

m.nIsChngTarif  = 0
m.nIsChngTarifV = 0
m.nIsChngStkd   = 0
m.nIsChngStkdV  = 0

SET RELATION TO cod INTO tarimu
SCAN 
 m.cod     = cod
 m.tarif   = tarif
 m.tarif_v = tarif_v
 m.stkd    = stkd
 m.stkdv   = stkdv
 
 IF tarif!=tarimu.tarif
  REPLACE tarif WITH tarimu.tarif
  m.nIsChngTarif  = m.nIsChngTarif + 1
 ENDIF 
 IF tarif_v!=tarimu.tarif
  REPLACE tarif_v WITH tarimu.tarif
  m.nIsChngTarifV  = m.nIsChngTarifV + 1
 ENDIF 
 IF stkd!=tarimu.stkd
  REPLACE stkd WITH tarimu.stkd
  m.nIsChngstkd  = m.nIsChngstkd + 1
 ENDIF 
 IF stkdv!=tarimu.stkd
  REPLACE stkdv WITH tarimu.stkd
  m.nIsChngstkdv  = m.nIsChngstkdv + 1
 ENDIF 
 
 m.tarif   = tarif
 m.stkd    = stkd
 m.n_kd    = n_kd
 
 IF IsMes(m.Cod) OR IsVmp(m.Cod) OR IsKDs(m.cod) OR INLIST(m.cod,97014,97015,97016,97017,97018,197014,197015,197106)
  IF m.stkd!=0 AND ROUND(m.tarif/m.stkd,0)!=m.n_kd
   REPLACE n_kd WITH ROUND(m.tarif/m.stkd,0)
  ENDIF 
 ENDIF 

ENDSCAN 
SET RELATION OFF INTO tarimu

IF m.nIsChngTarif > 0
* MESSAGEBOX(CHR(13)+CHR(10)+'намнбкемн '+TRANSFORM(m.nIsChngTarif,'99999')+' гмювемхи TARIF '+CHR(13)+CHR(10)+;
  'б тюике TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 
IF m.nIsChngTarifV > 0
* MESSAGEBOX(CHR(13)+CHR(10)+'намнбкемн '+TRANSFORM(m.nIsChngTarifV,'99999')+' гмювемхи TARIF_V '+CHR(13)+CHR(10)+;
  'б тюике TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 
IF m.nIsChngStkd > 0
* MESSAGEBOX(CHR(13)+CHR(10)+'намнбкемн '+TRANSFORM(m.nIsChngStkd,'99999')+' гмювемхи STKD '+CHR(13)+CHR(10)+;
  'б тюике TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 
IF m.nIsChngStkdV > 0
* MESSAGEBOX(CHR(13)+CHR(10)+'намнбкемн '+TRANSFORM(m.nIsChngStkdV,'99999')+' гмювемхи STKDV '+CHR(13)+CHR(10)+;
  'б тюике TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 

ELSE 
 MESSAGEBOX(CHR(13)+CHR(10)+'нрясрябсер ндхм хг вершпеу менаундхлшу'+;
  CHR(13)+CHR(10)+'дкъ янгдюмхъ TARIFN тюикнб. TARIFN ме оепеянгдюм!',0+64,'')
ENDIF 

IF USED('tarif')
 SELECT tarif 
 COPY TO &pCommon\trf
 ZAP 
 APPEND FROM &pCommon\trf
 ERASE &pCommon\trf.dbf
 SET ORDER TO 
 DELETE TAG ALL 
 USE IN tarif
ENDIF 
IF USED('tarimu')
 SELECT tarimu
 SET ORDER TO 
 DELETE TAG ALL 
 USE IN tarimu
ENDIF 
IF USED('reesus')
 SELECT reesus
 SET ORDER TO 
 DELETE TAG ALL 
 USE IN reesus
ENDIF 
IF USED('reesms')
 SELECT reesms
 SET ORDER TO 
 DELETE TAG ALL 
 USE IN reesms
ENDIF 
IF USED('usvmp')
 SELECT usvmp
 SET ORDER TO 
 DELETE TAG ALL 
 USE IN usvmp
ENDIF 

  fso.CopyFile(pcommon+'\tarifn.dbf', pBase+'\'+gcPeriod+'\NSI\tarifn.dbf')
  WAIT CLEAR 

  DO comreind

  WAIT CLEAR 
 ENDIF 
ELSE 

IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\errsmee.dbf')
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\errsmee', 'errs', 'shar')>0
  IF USED('errs')
   USE IN errs
  ENDIF 
 ELSE 
  SELECT errs
  IF FIELD('vp')!='VP'
   USE IN errs
   IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\errsmee', 'errs', 'excl')>0
   ELSE 
    SELECT errs
    ALTER table errs ADD COLUMN vp n(1)
   ENDIF 
  ENDIF 
  IF USED('errs')
   USE IN errs
  ENDIF 
 ENDIF 
ENDIF 

 IF fso.FileExists(pBase+'\'+gcPeriod+'\NSI\sprlpuxx.dbf')
  IF !fso.FileExists(pBase+'\'+gcPeriod+'\NSI\horlpu.dbf')
   IF OpenFile(pBase+'\'+gcPeriod+'\NSI\sprlpuxx', 'sprlpu', 'shar')>0
    IF USED('sprlpu')
     USE IN sprlpu
    ENDIF 
   ELSE 
    SELECT sprlpu 
    COPY FOR tpn='4' TO pBase+'\'+gcPeriod+'\NSI\horlpu' ;
     FIELDS lpu_id,fil_id,tpn,vmp,mcod,name,fullname,cokr,adres
    IF OpenFile(pBase+'\'+gcPeriod+'\NSI\horlpu', 'horlpu', 'shar')>0
     IF USED('horlpu')
      USE IN horlpu
     ENDIF 
    ELSE 
     SELECT horlpu
     INDEX ON lpu_id TAG lpu_id
     INDEX ON fil_id TAG fil_id 
     INDEX ON mcod TAG mcod
     USE IN horlpu
    ENDIF 
    USE IN sprlpu
   ENDIF 
  ENDIF 
 ENDIF 
ENDIF 


RETURN 


FUNCTION IsItUs(m.usl)
 PRIVATE m.usl
 m.IsItUs = .F.
 IF BETWEEN(m.usl,1001,60999) OR BETWEEN(m.usl,96001,99999) OR ;
  BETWEEN(m.usl,101001,160999) OR BETWEEN(m.usl,196001,199999)
  m.IsItUs = .T.
 ENDIF 
RETURN m.IsItUs

PROCEDURE ActualizeNSI
 PARAMETERS para1
 csprfile = para1

 oSettings.CodePage(csprfile, 866, .t.)
 
 IF OpenFile(csprfile, 'spr', 'excl')>0
  =Exit()
  RETURN 
 ENDIF 
 
 SELECT spr
 
 nfields = FCOUNT('spr')
 IF nfields != 6
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'йнкхвеярбн онкеи ('+STR(nfields,1)+') ме пюбмн 6!',0+16,'')
  =Exit()
  RETURN 
 ENDIF 
 
 IF !checkspr()
  RETURN 
 ENDIF 
 
* MESSAGEBOX(CHR(13)+CHR(10)+'тюик опюбхкэмши!'+CHR(13)+CHR(10),0+64,'')
 
 ospr = fso.GetFile(csprfile)
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
    IF m.intr_data<=tdat1
*     MESSAGEBOX(m.name_eta+'.dbf'+' -> '+m.etname+'xx.dbf' ,0+64,'')
     fso.CopyFile(odir+'\'+m.name_eta+'.dbf', pcommon+'\'+m.etname+'xx.dbf', .t.)
    ENDIF 
   ENDIF 
  ENDIF 
 ENDSCAN 
 
 =Exit() 
RETURN 

FUNCTION checkspr
 F1Name = FIELD(1)
 IF F1Name != 'SCOD'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'мюхлемнбюмхе оепбнцн онкъ ('+F1Name+') днкфмн ашрэ SCOD!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F1Type = VARTYPE(&F1Name)
 IF F1Type!='C'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'рхо оепбнцн онкъ '+F1Name+' днкфем ашрэ CHAR!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F1Size = FSIZE(F1Name)
 IF F1Size != 10
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'пюглепмнярэ онкъ '+F1Name+' днкфмю ашрэ 10!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

 F2Name = FIELD(2)
 IF F2Name != 'CUR_VER'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'мюхлемнбюмхе брнпнцн онкъ ('+F2Name+') днкфмн ашрэ CUR_VER!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F2Type = VARTYPE(&F2Name)
 IF F2Type!='C'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'рхо онкъ '+F2Name+' днкфем ашрэ CHAR!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F2Size = FSIZE(F2Name)
 IF F2Size != 10
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'пюглепмнярэ онкъ '+F2Name+' днкфмю ашрэ 10!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

 F3Name = FIELD(3)
 IF F3Name != 'FULL_NAME'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'мюхлемнбюмхе рперэецн онкъ ('+F3Name+') днкфмн ашрэ FULL_NAME!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F3Type = VARTYPE(&F3Name)
 IF F3Type!='C'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'рхо онкъ '+F3Name+' днкфем ашрэ CHAR!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F3Size = FSIZE(F3Name)
 IF F3Size != 120
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'пюглепмнярэ онкъ '+F3Name+' днкфмю ашрэ 120!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

 F4Name = FIELD(4)
 IF F4Name != 'INTR_DATA'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'мюхлемнбюмхе вербепрнцн онкъ ('+F4Name+') днкфмн ашрэ INTR_DATA!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F4Type = VARTYPE(&F4Name)
 IF F4Type!='D'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'рхо онкъ '+F4Name+' днкфем ашрэ DATE!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

 F5Name = FIELD(5)
 IF F5Name != 'NAME_ETA'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'мюхлемнбюмхе оърнцн онкъ ('+F5Name+') днкфмн ашрэ NAME_ETA!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F5Type = VARTYPE(&F5Name)
 IF F5Type!='C'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'рхо онкъ '+F5Name+' днкфем ашрэ CHAR!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F5Size = FSIZE(F5Name)
 IF F5Size != 8
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'пюглепмнярэ онкъ '+F5Name+' днкфмю ашрэ 8!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

 F6Name = FIELD(6)
 IF F6Name != 'CRC_ETA'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'мюхлемнбюмхе ьеярнцн онкъ ('+F6Name+') днкфмн ашрэ CRC_ETA!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F6Type = VARTYPE(&F6Name)
 IF F6Type!='C'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'рхо онкъ '+F6Name+' днкфем ашрэ CHAR!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F6Size = FSIZE(F6Name)
 IF F6Size != 10
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'пюглепмнярэ онкъ '+F6Name+' днкфмю ашрэ 10!',0+16,'')
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
