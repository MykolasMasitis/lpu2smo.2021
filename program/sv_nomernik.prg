PROCEDURE sv_nomernik

IF MESSAGEBOX('унрхре ябепхрэ опхйпеокемхе я мнлепмхйнл?',0+64,'')=7
 RETURN 
ENDIF 

IF RECCOUNT('outs')<=100
 MESSAGEBOX('мнлепмхй осяр!',0+64,'')
 USE IN outs
 SELECT aisoms
 RETURN 
ENDIF 

SELECT AisOms
SCAN FOR !DELETED()
 m.mcod  = mcod
 m.ppath = pbase+'\'+m.gcperiod
 IF !fso.FolderExists(m.ppath)
  LOOP 
 ENDIF 
 m.ppath = pbase+'\'+m.gcperiod+'\'+m.mcod
 IF !fso.FileExists(m.ppath+'\people.dbf')
  LOOP 
 ENDIF 
 IF OpenFile(m.ppath+'\people', 'people', 'shar')>0
  IF USED('people')
   USE IN people
  ENDIF 
  SELECT aisoms
  LOOP 
 ENDIF 
 
 =sv_pr(m.mcod)

 USE IN people 
 SELECT aisoms
 MailView.refresh
ENDSCAN 
WAIT CLEAR 

*USE IN outs
SELECT aisoms

MESSAGEBOX("напюанрйю гюйнмвемю",0+64,"")

RETURN 


