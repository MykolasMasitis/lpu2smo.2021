PROCEDURE SumMeFiles
 m.me_dir        = m.pOut+'\'+m.gcperiod
 
 oMailDir        = fso.GetFolder(m.me_dir)
 MailDirName     = oMailDir.Path
 oFilesInMailDir = oMailDir.Files
 nFilesInMailDir = oFilesInMailDir.Count

 MESSAGEBOX('Œ¡Õ¿–”∆≈ÕŒ '+ALLTRIM(STR(nFilesInMailDir))+' ‘¿…ÀŒ¬!', 0+64, '')

 IF nFilesInMailDir<=0
  RETURN 
 ENDIF 

 IF fso.FileExists(m.me_dir+'\ME'+m.qcod+'.dbf')
  fso.DeleteFile(m.me_dir+'\ME'+m.qcod+'.dbf')
 ENDIF 

 *fso.CopyFile(pTempl+'\MEqqnnnn.dbf', m.me_dir+'\ME'+m.qcod+'.dbf')
 
 CREATE CURSOR me&qcod (lpu_id n(4), period_e c(6), et c(1), s_opl_e n(12,2), s_sank n(12,2))
 
 *IF OpenFile(m.me_dir+'\ME'+m.qcod, 'me', 'shar')>0
 * IF USED('me')
 *  USE IN me 
 * ENDIF 
 * RELEASE m.me_dir, oMailDir, MailDirName, oFilesInMailDir, nFilesInMailDir
 * RETURN 
 *ENDIF 
 
 SELECT me&qcod

 FOR EACH oFileInMailDir IN oFilesInMailDir
  m.BFullName = oFileInMailDir.Path
  m.bname     = oFileInMailDir.Name
  
  IF LEFT(UPPER(m.bname),2)<>'ME'
   LOOP 
  ENDIF 
  
  APPEND FROM &BFullName

 ENDFOR 
 
 COPY TO &me_dir\ME&qcod
 
 SELECT m.gcperiod as period, period_e as e_period, lpu_id as lpuid,;
 	et as et, SUM(s_opl_e) sexp, 000.00 as stpn, SUM(s_sank) as s_sank;
 	FROM me&qcod WHERE s_opl_e>0 ;
 	GROUP BY lpu_id, period_e, et INTO TABLE &me_dir\svmee
 SELECT svmee 
 INDEX on lpuid TAG lpuid 
 	
 CLOSE TABLES ALL 
 
 RELEASE m.me_dir, oMailDir, MailDirName, oFilesInMailDir, nFilesInMailDir

 MESSAGEBOX('OK!',0+64,'')
 
RETURN 