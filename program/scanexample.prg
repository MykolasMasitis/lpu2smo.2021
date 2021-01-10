* ������
lo_form=CREATEOBJECT('Form')		&& ������ �����
lo_form.addObject('lbl_strih','BarCodLabel')	&& ��������� ����� �� ��������� ��������
lo_form.show(1)

READ events

* �������� ������ �� ��������� ��������
DEFINE CLASS BarCodLabel as Label
	autosize=.t.
	fontsize=30
	visible=.t.
	enabled=.t.
	*\\\\\\\\\\\\\\\
	PROCEDURE init
		thisForm.addObject('scanner','_scanner')	&& �������� ������
		thisForm.scanner.on=.t.	&&  �������� �������
		*		thisForm.scanner.commport=1	&&  ��������  ����� COM1:
		thisForm.scanner.onRun='thisform.'+this.ParentPath(this,'form')+'.refresh()'	&& ������ �� �������� �� �������
		thisForm.scanner.activate()	&& ������
	endproc
	*\\\\\\\\\\\\\\
	PROCEDURE refresh
		this.Caption=thisForm.scanner.scancode		&& ��������� ����� �� �������
	endproc
	*\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE  ParentPath			&& ���� �� ������� �� ���������� ������
		lParameter loName,lc_StopClass
		*             ^ ������, �������� ������������� ������ (�������� 'form')
		local lc_ret,lc_tmp
		if vartype(loName)#'O'
			return ''
		endIf
		if vartype(m.lc_StopClass)#'C'
			lc_StopClass='Application'
		endIf
		m.lc_StopClass=alltrim(lower(m.lc_StopClass))

		lc_ret=loName.name
		lo_tmp=loName.parent
		do while vartype(lo_tmp)='O' .and. alltrim(lower(lo_tmp.Class))#m.lc_StopClass
			lc_ret=lo_tmp.name+'.'+m.lc_ret

			IF type('lo_tmp.parent')='O'
				lo_tmp=lo_tmp.parent
			ELSE
				exit
			endIf
		endDo
		return m.lc_ret
	endProc
ENDDEFINE
*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
DEFINE CLASS _scanner as custom
	*!!!* ������� ��������� ��� ������ ����������� �������.
	* ��������� ����������:
	on=.T.			&& ������ ������������
	commport= 7		&& COM-���� ���� �������� ������ ����� �����
	*         ^    @ ������ ������� � com port
	*^ ��������� ����������

	scancode=PADL('',12,'0')	&& ��������� �����
	onRun='this.parent.parent.onComm()'	&& ������ �� ��������� ���������
	*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE init
	ENDPROC
	*\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE activate		&& ���������
		IF !this.on
			RETURN .f.
		endIf
		if !(VARTYPE(this.scanner)='O')
			set Classlib To 'SCANCOM' additive	&& ��������� ���������� ������� �� ����� ����
			If !('SCANCOM.VCX'$Set("Classlib" ))
				wait WINDOW '!Set("Classlib" )'
				RETURN .f.
			ENDIF
			this.addObject('scanner','scancom.scanner_')
		endIf
		WITH This.scanner.olecomm
			* ��������� ���������� �� ����� ���������
			.commport=this.commport		&& COM-���� ���� �������� ������ ����� �����
			.EOFEnable=.T.
			.RTSEnable=.T.
			.RThreshold=1
			.Settings="9600,n,8,1"
			.SThreshold=0
			try
				.portOpen=.T.
			CATCH
			ENDTRY
			* ��������� ���������� �� ����� ���������
		endWith
	ENDPROC
	*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	*\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE off		&& ���������
		if (VARTYPE(this.scanner)='O')
			if	This.scanner.olecomm.portOpen=.T.
				This.scanner.olecomm.portOpen=.f.
			endIf
		endIf
	ENDPROC
	*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE destroy
		this.off()		&& ���������
	endProc
	*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE onComm
		WAIT WINDOW '����������� :'+this.scancode nowait
	endProc

ENDDEFINE

*SCANCOM.VCX
