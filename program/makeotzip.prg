FUNCTION MakeOtZip(lcPath, cmcod, clpuid)
 
 m.mmy = PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),1)

 m.ctrl    = 'ctrl'+m.qcod+'.dbf' && МЭК
 m.MeFile  = 'Me'+m.qcod+STR(lpuid,4)+'.dbf' && МЭЭ

 m.PrFile  = 'Pr' + m.qcod + m.mmy + '.pdf'                && Протокол обработки
 m.McFile  = 'Mc' + STR(lpuid,4) + m.qcod + m.mmy + '.pdf' && Акт МЭК
 m.MkFile  = 'Mk' + STR(lpuid,4) + m.qcod + m.mmy + '.pdf' && Реестр Актов МЭК 
 m.MtFile  = 'Mt' + STR(lpuid,4) + m.qcod + m.mmy + '.pdf' && Табличная форма акта МЭК

 m.PdfFile = 'pdf' + m.qcod + m.mmy+'.pdf' && Акт об оплате расчетов по подушевому финансированию
 m.UDFile  = 'ud'+m.qcod+STR(lpuid,4)+'.dbf'
 m.UPFile  = 'up'+m.qcod+STR(lpuid,4)+'.dbf'

 IF !fso.FileExists(lcpath+'\'+m.ctrl)
  RETURN .F.
 ENDIF 
 IF !fso.FileExists(lcpath+'\'+m.MeFile)
  IF fso.FileExists(pOut+'\'+m.gcperiod+'\'+m.MeFile)
   fso.CopyFile(pOut+'\'+m.gcperiod+'\'+m.MeFile, lcpath+'\'+m.MeFile)
  ENDIF 
*  RETURN .F.
 ENDIF 
 IF !fso.FileExists(lcpath+'\'+m.PrFile)
  RETURN .F.
 ENDIF 
 IF !fso.FileExists(lcpath+'\'+m.McFile)
  RETURN .F.
 ENDIF 
 IF !fso.FileExists(lcpath+'\'+m.MkFile)
  RETURN .F.
 ENDIF 
 IF !fso.FileExists(lcpath+'\'+m.MtFile)
  RETURN .F.
 ENDIF 
 IF !fso.FileExists(lcpath+'\'+m.PdfFile)
*  RETURN .F.
 ENDIF 
 IF !fso.FileExists(lcpath+'\'+m.UDFile)
*  RETURN .F.
 ENDIF 
 IF !fso.FileExists(lcpath+'\'+m.UPFile)
*  RETURN .F.
 ENDIF 

 ZipFile = 'Ot'+clpuid+m.qcod+'.'+m.mmy
 IF fso.FileExists(lcPath+'\'+ZipFile)
  fso.DeleteFile(lcPath+'\'+ZipFile)
 ENDIF 
 
 ZipOpen(lcPath+'\'+ZipFile)

 IF fso.FileExists(lcpath+'\'+m.ctrl)
  ZipFile(lcPath+'\'+m.ctrl)
  *MESSAGEBOX(lcPath+'\'+m.ctrl,0+64,cmcod)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+m.MeFile)
  ZipFile(lcPath+'\'+m.MeFile)
  *MESSAGEBOX(lcPath+'\'+m.MeFile,0+64,cmcod)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+m.PrFile)
  ZipFile(lcPath+'\'+m.PrFile)
  *MESSAGEBOX(lcPath+'\'+m.PrFile,0+64,cmcod)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+m.McFile)
  ZipFile(lcPath+'\'+m.McFile)
  *MESSAGEBOX(lcPath+'\'+m.McFile,0+64,cmcod)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+m.MkFile)
  ZipFile(lcPath+'\'+m.MkFile)
  *MESSAGEBOX(lcPath+'\'+m.MkFile,0+64,cmcod)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+m.MtFile)
  ZipFile(lcPath+'\'+m.MtFile)
  *MESSAGEBOX(lcPath+'\'+m.MtFile,0+64,cmcod)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+m.PdfFile)
  ZipFile(lcPath+'\'+m.PdfFile)
  *MESSAGEBOX(lcPath+'\'+m.PdfFile,0+64,cmcod)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+m.UDFile)
  ZipFile(lcPath+'\'+m.UDFile)
  *MESSAGEBOX(lcPath+'\'+m.UDFile,0+64,cmcod)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+m.UPFile)
  ZipFile(lcPath+'\'+m.UPFile)
  *MESSAGEBOX(lcPath+'\'+m.UPFile,0+64,cmcod)
 ENDIF 

 ZipClose()
 
 IF !fso.FileExists(lcPath+'\'+ZipFile)
  RETURN .F.
 ENDIF 

RETURN  .T.