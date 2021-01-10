PROCEDURE SetEnv
 PARAMETERS para1, para2
 
 PUBLIC pCommon, pBase
 
 pCommon = para1
 pBase   = para2
 
 SET PROCEDURE TO Utils

 PUBLIC fso AS SCRIPTING.FileSystemObject, wshell AS Shell.Application

 fso = CREATEOBJECT('Scripting.FileSystemObject')

RETURN 