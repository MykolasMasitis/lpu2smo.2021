PROCEDURE ConfirmOwn
 IF qq=m.qcod
  MESSAGEBOX("опхмюдкефмнярэ оюжхемрю"+CHR(13)+CHR(10)+"ондрбепфдемю пюмее!",0+64,"")
 ENDIF 
 
 _vfp.ActiveForm.LockScreen = .t.
 REPLACE qq WITH m.qcod, sv WITH '211'
 _vfp.ActiveForm.LockScreen = .f.
 _vfp.ActiveForm.refresh
RETURN 