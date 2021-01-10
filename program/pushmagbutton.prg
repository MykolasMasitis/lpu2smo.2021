PROCEDURE PushMagButton
 IF MESSAGEBOX('бш унрхре бшонкмхрэ онякеднбюрекэмн рпх ноепюжхх?'+CHR(13)+CHR(10)+;
 	'1. йнппейрхпнбйю ярпсйрспш пюанвху аюг'+CHR(13)+CHR(10)+;
 	'2. оепехмдейяюжхъ пюанвху аюг'+CHR(13)+CHR(10)+;
 	'3. оепехмдейяюжхъ мях',4+32,'')=7
  RETURN 
 ENDIF 
 
 WAIT "йнппейрхпнбйю ярпсйрспш ад..." WINDOW NOWAIT 
 DO CorStruct
 WAIT CLEAR 

 WAIT "оепехмдейяюжхъ пюанвху аюг..." WINDOW NOWAIT 
 DO BasReind
 WAIT CLEAR 

 WAIT "оепехмдейяюжхъ мях..." WINDOW NOWAIT 
 DO ComReind
 WAIT CLEAR 
 
 MESSAGEBOX('OK!',0+64,'')
 
RETURN 