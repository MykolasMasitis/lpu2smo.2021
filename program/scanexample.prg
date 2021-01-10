* Пример
lo_form=CREATEOBJECT('Form')		&& Создаём форму
lo_form.addObject('lbl_strih','BarCodLabel')	&& Добавляем текст со встроеным сканером
lo_form.show(1)

READ events

* Описание текста со встроеным сканером
DEFINE CLASS BarCodLabel as Label
	autosize=.t.
	fontsize=30
	visible=.t.
	enabled=.t.
	*\\\\\\\\\\\\\\\
	PROCEDURE init
		thisForm.addObject('scanner','_scanner')	&& Добавили сканер
		thisForm.scanner.on=.t.	&&  Полюбому включен
		*		thisForm.scanner.commport=1	&&  Полюбому  через COM1:
		thisForm.scanner.onRun='thisform.'+this.ParentPath(this,'form')+'.refresh()'	&& Ссылка на действия со сканера
		thisForm.scanner.activate()	&& инитим
	endproc
	*\\\\\\\\\\\\\\
	PROCEDURE refresh
		this.Caption=thisForm.scanner.scancode		&& последний штрих со сканера
	endproc
	*\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE  ParentPath			&& Путь до обьекта от выбранного класса
		lParameter loName,lc_StopClass
		*             ^ ОБЪЕКТ, название родительского класса (например 'form')
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
	*!!!* Меняйте настройку для своего усторойства ручками.
	* Настройка устройства:
	on=.T.			&& Сканер использовать
	commport= 7		&& COM-порт куда подкючен Сканер Штрих Кодов
	*         ^    @ Работа сканнер № com port
	*^ Настройка устройства

	scancode=PADL('',12,'0')	&& Последний штрих
	onRun='this.parent.parent.onComm()'	&& Ссылка на процедуру обработки
	*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE init
	ENDPROC
	*\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE activate		&& Включение
		IF !this.on
			RETURN .f.
		endIf
		if !(VARTYPE(this.scanner)='O')
			set Classlib To 'SCANCOM' additive	&& Открываем библиотеку классов по имени цели
			If !('SCANCOM.VCX'$Set("Classlib" ))
				wait WINDOW '!Set("Classlib" )'
				RETURN .f.
			ENDIF
			this.addObject('scanner','scancom.scanner_')
		endIf
		WITH This.scanner.olecomm
			* Настройка устройства во время активации
			.commport=this.commport		&& COM-порт куда подкючен Сканер Штрих Кодов
			.EOFEnable=.T.
			.RTSEnable=.T.
			.RThreshold=1
			.Settings="9600,n,8,1"
			.SThreshold=0
			try
				.portOpen=.T.
			CATCH
			ENDTRY
			* Настройка устройства во время активации
		endWith
	ENDPROC
	*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	*\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE off		&& Включение
		if (VARTYPE(this.scanner)='O')
			if	This.scanner.olecomm.portOpen=.T.
				This.scanner.olecomm.portOpen=.f.
			endIf
		endIf
	ENDPROC
	*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE destroy
		this.off()		&& Включение
	endProc
	*\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	PROCEDURE onComm
		WAIT WINDOW 'штрихнулись :'+this.scancode nowait
	endProc

ENDDEFINE

*SCANCOM.VCX
