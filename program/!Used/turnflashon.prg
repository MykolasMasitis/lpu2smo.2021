PROCEDURE TurnFlashOn
 TRY 
  StartValue = wshshell.regread("HKLM\SYSTEM\CurrentControlSet\Services\USBSTOR\Start") && 4-����������, 3-����������� �����
  MESSAGEBOX("HKLM\SYSTEM\CurrentControlSet\Services\USBSTOR\Start="+STR(StartValue,1), 0+64, "")
  IF StartValue != 3
   TRY 
    wshshell.RegWrite("HKLM\SYSTEM\CurrentControlSet\Services\USBSTOR\Start",3,"REG_DWORD")
    MESSAGEBOX('������ ��������!',0+64,'')
   CATCH 
    MESSAGEBOX('�� ������� �������� ������!!',0+64,'')
   ENDTRY 
  ENDIF 
 CATCH 
  MESSAGEBOX("�� ������� ��������� �������� �����"+CHR(10)+CHR(13)+"HKLM\SYSTEM\CurrentControlSet\Services\USBSTOR\Start", 0+64, "")
 ENDTRY 

 TRY 
  StartValue = wshshell.regread("HKLM\SYSTEM\CurrentControlSet\Control\StorageDevicePolicies\WriteProtect") && 0-����� ������,1-����� ������
  MESSAGEBOX("HKLM\SYSTEM\CurrentControlSet\Control\StorageDevicePolicies\WriteProtect="+STR(StartValue,1), 0+64, "")
  IF StartValue != 0
   TRY 
    wshshell.RegWrite("HKLM\SYSTEM\CurrentControlSet\Control\StorageDevicePolicies\WriteProtect",0,"REG_DWORD")
    MESSAGEBOX('������ ��������!',0+64,'')
   CATCH 
    MESSAGEBOX('�� ������� �������� ������!!',0+64,'')
   ENDTRY 
  ENDIF 
 CATCH 
  MESSAGEBOX("�� ������� ��������� �������� �����"+CHR(10)+CHR(13)+;
  "HKLM\SYSTEM\CurrentControlSet\Control\StorageDevicePolicies\WriteProtect", 0+64, "")
 ENDTRY 

RETURN 
