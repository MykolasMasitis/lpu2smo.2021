FUNCTION base64  
  * ?????????????? ???????? ?????? ? ??? BASE64  
   param tcinput  
   local i,ii,cBASE64,tlen,dobsimbol  
   b64alphabet= 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'  
   cbase64=''  
   dobsimbol=''  
   tlen=len(tcinput)  
   for i=1 to tlen step 24  
    triada=subst(tcinput,i,24)  
    maxdiap=24  
    do case  
     case len(triada)=8  
      maxdiap=12  
      triada=triada+'0000'  
      dobsimbol='=='  
     case len(triada)=16  
      maxdiap=18  
      triada=triada+'00'  
      dobsimbol='='  
    endcase   
    for ii=1 to maxdiap step 6  
     cbase64=cbase64+subst(b64alphabet,c2d(subst(triada,ii,6))+1,1)  
    endfor  
   endfor  
   cbase64=cbase64 +dobsimbol  
  return cbase64   
 ****  
  funct c2d  
  * ??????????? 6-?? ???? ??? ? ????????  
   param c2  
   tnout=0  
    for i=1 to 6   
     tnout=tnout+ iif(subst(c2,i,1)='1',2^(6-i),0)  
    endfor  
   return tnout   
  ***
  

FUNCTION EncodeBase641
LPARAMETERS cInFile  
 * cInFile - ???? ? ?????  
    
 *~* ??????????? ? ????????? Base64  
 *~* ????? ?? 6 ??? ??????? ? ???????, ??????????? 0?20 ? ???????? ??, ??? ?????  
 *~* ?.?. ?????? ?? ???????? ? ????? ???????? ?????????, ?? ???????? ? ??????.  
    
  LOCAL nFile, nFile1, nFile2, cTmpFile, cTmpFile1, nI, nVar, cStr,;  
  	n64 as Integer, n64_1 as Integer, mask as Integer  
    
 * ??????????? ?????? ??????????  
  cTmpFile = Gettmpname('tmp')  
  STRTOFILE('ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/', cTmpFile)  
  nFile1 = FOPEN(cTmpFile)  
    
 * ????????? ??????? ????  
  nFile = FOPEN(cInFile)  
    
 * ????????? ???????? ????   
  cTmpFile1 = Gettmpname('tmp')  
  nFile2 = FCREATE(cTmpFile1)  
  FPUTS(nFile2, 'Content-Type: application/octet-stream; name="'+JUSTFNAME(cInFile)+'"')  
  FPUTS(nFile2, 'Content-transfer-encoding: base64')  
  FPUTS(nFile2, 'Content-Disposition: attachment; filename="'+JUSTFNAME(cInFile)+'"')  
  FPUTS(nFile2, '')  
    
 * ???????? ????  
  nVar=0  
  DO WHILE NOT FEOF(nFile)  
  	mask = 0xFC0000  
 * ?????? ?? ????? ?? 3 ?????  
  	n64 = BITOR(ASC(FREAD(nFile,1))*65536, ASC(FREAD(nFile,1))*256, ASC(FREAD(nFile,1)))  
  	FOR nI=3 TO 0 STEP -1  
  		n64_1 = BITRSHIFT(BITAND(n64,mask),nI*6)  
  		FSEEK(nFile1,n64_1)  
  		FWRITE(nFile2, FREAD(nFile1,1))  
  		mask = BITRSHIFT(mask,6)  
  	ENDFOR  
  	nVar = nVar + 1  
  	IF nVar % 18 = 0  
  		FWRITE(nFile2, CRLF)  
  	ENDIF   
  ENDDO  
    
 * ?????????? ????? ????????? ?????  
  nVar = FSEEK(nFile,0,2)  
 * ?????????? ?????????? "??????" ? ? ????? ??????????????? ?????  
  nVar = FSEEK(nFile2,0,2)-CEILING(nVar/3*4+INT(nVar/54)*2)  
 * ???????? "??????" A ?? =  
  FSEEK(nFile2,-nVar,2)  
  FWRITE(nFile2,'===',nVar)  
    
  FCLOSE(nFile)  
  FCLOSE(nFile1)  
  FCLOSE(nFile2)  
    
  cStr = FILETOSTR(cTmpFile1)  
  DELETE FILE (cTmpFile)  
  DELETE FILE (cTmpFile1)  
    
RETURN cStr
  
FUNCTION EncodeBase64Old
    
  LPARAMETER tcString  
    
  LOCAL lcByte  
    
    
  LOCAL laLookUpTable[64]  
    
  LOCAL lcEncodedString  
    
  LOCAL lnLenString  
  LOCAL lnNewLenString  
    
  LOCAL lnCount  
  LOCAL lnPaddingCount  
  LOCAL lnBlockCount  
    
  LOCAL lnByte  
  LOCAL lnByte1  
  LOCAL lnByte2  
  LOCAL lnByte3  
    
  LOCAL lnTmpByte  
  LOCAL lnTmpByte1  
  LOCAL lnTmpByte2  
  LOCAL lnTmpByte3  
  LOCAL lnTmpByte4  
    
  lcEncodedString = ''  
    
  laLookUpTable[01] = 'A'  
  laLookUpTable[02] = 'B'  
  laLookUpTable[03] = 'C'  
  laLookUpTable[04] = 'D'  
  laLookUpTable[05] = 'E'  
  laLookUpTable[06] = 'F'  
  laLookUpTable[07] = 'G'  
  laLookUpTable[08] = 'H'  
  laLookUpTable[09] = 'I'  
  laLookUpTable[10] = 'J'  
  laLookUpTable[11] = 'K'  
  laLookUpTable[12] = 'L'  
  laLookUpTable[13] = 'M'  
  laLookUpTable[14] = 'N'  
  laLookUpTable[15] = 'O'  
  laLookUpTable[16] = 'P'  
  laLookUpTable[17] = 'Q'  
  laLookUpTable[18] = 'R'  
  laLookUpTable[19] = 'S'  
  laLookUpTable[20] = 'T'  
  laLookUpTable[21] = 'U'  
  laLookUpTable[22] = 'V'  
  laLookUpTable[23] = 'W'  
  laLookUpTable[24] = 'X'  
  laLookUpTable[25] = 'Y'  
  laLookUpTable[26] = 'Z'  
    
  laLookUpTable[27] = 'a'  
  laLookUpTable[28] = 'b'  
  laLookUpTable[29] = 'c'  
  laLookUpTable[30] = 'd'  
  laLookUpTable[31] = 'e'  
  laLookUpTable[32] = 'f'  
  laLookUpTable[33] = 'g'  
  laLookUpTable[34] = 'h'  
  laLookUpTable[35] = 'i'  
  laLookUpTable[36] = 'j'  
  laLookUpTable[37] = 'k'  
  laLookUpTable[38] = 'l'  
  laLookUpTable[39] = 'm'  
  laLookUpTable[40] = 'n'  
  laLookUpTable[41] = 'o'  
  laLookUpTable[42] = 'p'  
  laLookUpTable[43] = 'q'  
  laLookUpTable[44] = 'r'  
  laLookUpTable[45] = 's'  
  laLookUpTable[46] = 't'  
  laLookUpTable[47] = 'u'  
  laLookUpTable[48] = 'v'  
  laLookUpTable[49] = 'w'  
  laLookUpTable[50] = 'x'  
  laLookUpTable[51] = 'y'  
  laLookUpTable[52] = 'z'  
    
  laLookUpTable[53] = '0'  
  laLookUpTable[54] = '1'  
  laLookUpTable[55] = '2'  
  laLookUpTable[56] = '3'  
  laLookUpTable[57] = '4'  
  laLookUpTable[58] = '5'  
  laLookUpTable[59] = '6'  
  laLookUpTable[60] = '7'  
  laLookUpTable[61] = '8'  
  laLookUpTable[62] = '9'  
  laLookUpTable[63] = '+'  
  laLookUpTable[64] = '/'  
    
  lnLenString = LEN(tcString)  
    
  IF lnLenString > 0  
    
  	IF (lnLenString % 3) = 0  
  		lnPaddingCount = 0  
  		lnBlockCount   = lnLenString / 3  
  	ELSE  
 *--	Need to add padding  
  		lnPaddingCount = 3 - (lnLenString % 3)  
  		lnBlockCount   = (lnLenString + lnPaddingCount) / 3  
  	ENDIF  
    
  	lnNewLenString = lnLenString + lnPaddingCount  
    
  	IF 	lnNewLenString > lnLenString  
  		tcString = PADR(tcString, lnNewLenString, '0')  
  	ENDIF  
    
    
  	lcByte = ''  
    
    
  	FOR lnCount = 1 TO lnBlockCount  
    
  		lnByte1 = ASC(SUBSTR(tcString, (lnCount - 1) * 3 + 1, 1))  
  		lnByte2 = ASC(SUBSTR(tcString, (lnCount - 1) * 3 + 2, 1))		  
  		lnByte3 = ASC(SUBSTR(tcString, (lnCount - 1) * 3 + 3, 1))				  
    
 *-- First byte  
  		lnTmpByte1 = BITRSHIFT(BITAND(lnByte1, 252), 2)  
    
 *-- Second byte  
  		lnTmpByte  = BITLSHIFT(BITAND(lnByte1, 3), 4)  
  		lnTmpByte2 = BITRSHIFT(BITAND(lnByte2, 240), 4)  
  		lnTmpByte2 = lnTmpByte2 +  lnTmpByte  
    
 *-- Third byte  
  		lnTmpByte  = BITLSHIFT(BITAND(lnByte2, 15), 2)  
  		lnTmpByte3 = BITRSHIFT(BITAND(lnByte3, 192), 6)  
  		lnTmpByte3 = lnTmpByte3 +  lnTmpByte  
    
 *-- Fourth byte  
  		lnTmpByte4 = BITAND(lnByte3, 63)  
    
  		lcByte = lcByte + PADL(ALLTRIM(STR(lnTmpByte1)), 2, '0') + PADL(ALLTRIM(STR(lnTmpByte2)), 2, '0') + PADL(ALLTRIM(STR(lnTmpByte3)), 2, '0') + PADL(ALLTRIM(STR(lnTmpByte4)), 2, '0')  
    
    
  	NEXT  
    
  	FOR lnCount = 1 TO LEN(lcByte) STEP 2  
  		lnByte = VAL(SUBSTR(lcByte, lnCount, 2))  
    
    
  		IF lnByte >= 0 AND lnByte <= 63  
  			lcEncodedString = lcEncodedString + laLookUpTable[lnByte + 1]  
  		ELSE  
 *-- Should not happen  
  			lcEncodedString = lcEncodedString + ' '  
  		ENDIF  
    
  	NEXT  
    
 *-- Convert last 'A' to  '='  
  	DO CASE   
  	CASE lnPaddingCount = 0  
  	CASE lnPaddingCount = 1  
  		lcEncodedString = STUFF(lcEncodedString, LEN(lcEncodedString) - 1, 1, '=')  
  	CASE lnPaddingCount = 2		  
  		  
  		lcEncodedString = STUFF(lcEncodedString, LEN(lcEncodedString) - 1, 1, '=')		  
  		lcEncodedString = STUFF(lcEncodedString, LEN(lcEncodedString) - 2, 1, '=')				  
  		  
  	ENDCASE  
    
  ENDIF  
    
  RETURN lcEncodedString