PROCEDURE HVScreen

nW = 0
nH = 0

=ScreenSize(@nW, @nH)

MESSAGEBOX('���������� ��������: '+ALLTRIM(STR(nW))+'x'+ALLTRIM(STR(nH)),0+64,'')


RETURN 