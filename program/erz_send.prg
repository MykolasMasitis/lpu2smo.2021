PROCEDURE erz_send
PARAMETERS loForm

m.UserDir = m.usrmail

IF !fso.FolderExists(pAisOms+'\'+m.UserDir)
 MESSAGEBOX('нрясрярбсер дхпейрнпхъ ' + pAisOms + '\' + m.UserDir + '!', 0+16, '')
 RETURN 0
ENDIF 

m.nTotDirs  = 0
m.nNewDirs = 0

OldEscStatus = SET("Escape")
SET ESCAPE OFF 
CLEAR TYPEAHEAD 

SELECT AisOms
SCAN
 m.mcod    = mcod
 m.soapsts = soapsts
 m.bname   = bname
 
 IF EMPTY(m.bname)
  LOOP 
 ENDIF 

 *IF LOWER(m.loForm.name) = 'mailsoap' AND m.soapsts!='RECIEVED'
 * LOOP 
 *ENDIF 

 WAIT m.mcod WINDOW NOWAIT 

 lcDir = pBase+'\'+m.gcperiod+'\'+mcod

 m.nTotDirs  = m.nTotDirs + 1

 IF erz_status == 0
  m.t_0=SECONDS()
  ErzResult = OneErz(lcDir, .f.)
  m.t_1=SECONDS()
  IF ErzResult == .T.
   m.nNewDirs = m.nNewDirs + 1
   REPLACE erz_status WITH 1, t_2 WITH m.t_1-m.t_0
  ENDIF 
 ENDIF 

_vfp.ActiveForm.refresh 

 IF CHRSAW(0) 
  IF INKEY() == 27
   IF MESSAGEBOX('бш унрхре опепбюрэ напюанрйс?',4+32,'') == 6
    EXIT 
   ENDIF 
  ENDIF 
 ENDIF 

ENDSCAN 
WAIT CLEAR 

GO TOP 
_vfp.ActiveForm.refresh 

 SET ESCAPE &OldEscStatus

IF m.nNewDirs == 0
 MESSAGEBOX('опнялнрпемн '+STR(m.nTotDirs,3)+CHR(13)+CHR(10)+;
 'бяе гюопняш ятнплхпнбюмш пюмее!',0+64,'')
 RETURN 
ELSE 
 MESSAGEBOX('намюпсфемн '+ALLTRIM(STR(m.nNewDirs))+' мнбшу кос!',0+64,'')
ENDIF 



