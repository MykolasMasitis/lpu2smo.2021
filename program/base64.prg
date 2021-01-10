Procedure Base64_ini  
  if !pemstatus(_screen,"base64_ptr",5)  
  	Declare Integer GetProcessHeap in Win32API  
  	Declare Integer HeapAlloc in Win32Api Integer, Integer, Integer  
  	Declare RtlMoveMemory in Win32API Integer, String, Integer  
    
  	local m.hhnd, m.ptr, m.st2  
    
  	m.st2= ;  
  		chr(0x55)+chr(0x89)+chr(0xE5)+chr(0x57)+chr(0x56)+chr(0x50)+chr(0x53)+chr(0x51)+ ;  
  		chr(0x52)+chr(0x8B)+chr(0x75)+chr(0x08)+chr(0x8B)+chr(0x7D)+chr(0x0C)+chr(0x8B)+ ;  
  		chr(0x4D)+chr(0x10)+chr(0x8B)+chr(0x5D)+chr(0x14)+chr(0x31)+chr(0xC0)+chr(0x55)+ ;  
  		chr(0x31)+chr(0xED)+chr(0x8A)+chr(0x06)+chr(0xC0)+chr(0xE8)+chr(0x02)+chr(0x8A)+ ;  
  		chr(0x04)+chr(0x03)+chr(0x88)+chr(0x07)+chr(0x47)+chr(0x8A)+chr(0x06)+chr(0x24)+ ;  
  		chr(0x03)+chr(0xC0)+chr(0xE0)+chr(0x04)+chr(0x88)+chr(0xC2)+chr(0x46)+chr(0x8A)+ ;  
  		chr(0x06)+chr(0xC0)+chr(0xE8)+chr(0x04)+chr(0x08)+chr(0xD0)+chr(0x8A)+chr(0x04)+ ;  
  		chr(0x03)+chr(0x88)+chr(0x07)+chr(0x47)+chr(0x8A)+chr(0x06)+chr(0x24)+chr(0x0F)+ ;  
  		chr(0xC0)+chr(0xE0)+chr(0x02)+chr(0x88)+chr(0xC2)+chr(0x46)+chr(0x8A)+chr(0x06)+ ;  
  		chr(0xC0)+chr(0xE8)+chr(0x06)+chr(0x08)+chr(0xD0)+chr(0x8A)+chr(0x04)+chr(0x03)+ ;  
  		chr(0x88)+chr(0x07)+chr(0x47)+chr(0x8A)+chr(0x06)+chr(0x24)+chr(0x3F)+chr(0x8A)+ ;  
  		chr(0x04)+chr(0x03)+chr(0x88)+chr(0x07)+chr(0x46)+chr(0x47)+chr(0x45)+chr(0x83)+ ;  
  		chr(0xFD)+chr(0x13)+chr(0x75)+chr(0x0A)+chr(0xC6)+chr(0x07)+chr(0x0D)+chr(0x47)+ ;  
  		chr(0xC6)+chr(0x07)+chr(0x0A)+chr(0x47)+chr(0x31)+chr(0xED)+chr(0x49)+chr(0x75)+ ;  
  		chr(0xA9)+chr(0x5D)+chr(0x5A)+chr(0x59)+chr(0x5B)+chr(0x58)+chr(0x5E)+chr(0x5F)+ ;  
  		chr(0x89)+chr(0xEC)+chr(0x5D)+chr(0xC2)+chr(0x10)+chr(0x00)  
    
  	m.hhnd=GetProcessHeap()  
  	m.ptr=HeapAlloc(m.hhnd,0,len(m.st2)+16)  
  	RtlMoveMemory(m.ptr,m.st2,len(m.st2))  
  	_screen.addproperty("base64_ptr",m.ptr)  
  ENDIF
RETURN 

Function Base64  
  lparameter m.str  
  local m.str_len  
  m.str_len=len(m.str)  
  if m.str_len=0  
  	return ""  
  endif  
    
  local m.st1, m.l3, m.lost, m.ost1, m.ost2, m.ostr, m.str_old, m.str_new, m.ret  
  m.st1="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"  
  m.l3=int(m.str_len/3)  
  m.lost=m.str_len%3  
  do case  
  	case m.lost=0  
  		m.ostr=""  
  	case m.lost=1  
  		m.ost1=right(m.str,1)  
  		m.ostr=substr(m.st1,bitrshift(asc(m.ost1),2)+1,1)+ ;  
  			substr(m.st1,bitlshift(bitand(asc(m.ost1),3),4)+1,1)+"=="  
  	case m.lost=2  
  		m.ost1=left(right(m.str,2),1)  
  		m.ost2=right(m.str,1)  
  		m.ostr=substr(m.st1,bitrshift(asc(m.ost1),2)+1,1)+ ;  
  			substr(m.st1,bitor(bitlshift(bitand(asc(m.ost1),3),4),bitrshift(asc(m.ost2),4))+1,1)+ ;  
  			substr(m.st1,bitlshift(bitand(asc(m.ost2),15),2)+1,1)+"="  
  endcase  
    
  if m.l3>0  
  	m.str_old=left(m.str,m.l3*3)  
  	m.str_new=repl(chr(0),m.l3*4+int(m.l3/19)*2)  
    
  	Declare CallWindowProc in Win32API Integer, String, String @, Integer, String  
    
  	CallWindowProc(_screen.base64_ptr, m.str_old, @m.str_new, m.l3, m.st1)  
  else  
  	m.str_new=""  
  endif  
    
  m.ret=m.str_new+m.ostr  
  if m.lost=0 and right(m.ret,1)=chr(10)  
  	m.ret=left(m.ret,len(m.ret)-2)  
  endif  
    
  return m.ret
    
FUNCTION base64o
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
