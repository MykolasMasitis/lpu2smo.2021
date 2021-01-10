***********************************************************************************************
*����� ��� ������ � Microsoft Excel � Open Office(Libre Office), ������ 1.2 �� 09.09.2013 �.  *
*																							  *
*		������������ ��� ������ � ������� xls/xlt											  *
***********************************************************************************************
*
* ������ ������������� ������:
*	Ap=CREATEOBJECT('OOEX')   &&���������� �� ��������� (Microsoft Excel)
*	Ap=CREATEOBJECT('OOEX',1) &&Microsoft Excel
*	Ap=CREATEOBJECT('OOEX',2) &&Open Office (Libre Office)
*
*            *   ��������� � �������   *
*       =======���� ��������======
* Ap.NewDoc()								�������� ������� ���������
* Ap.LoadDoc(��������)						��������� ������������ ���� (���� � ����� ������)
* Ap.CloseDoc()								�������� ���������
* Ap.SaveDoc()								���������� ��������� ���������
* Ap.SaveAsDoc(��������)					���������� ��������� ��������� �� ��������� ���� � ����� �����
*      ======��������� �����========
* Ap.SetBasePageSetup()						������� ��������� ����� (��������� � ����, ��� �������)
*      =======�����========
* Ap.ActivateSheetByIndex(�����)			��������� ����� �� ��� ������ (�������)
* Ap.SetActiveSheetName(�����_���_�����)	������ ��� ��������� ����� � �����
* Ap.SetSheetName(�����_�����,�����_���_�����)	������ ��� ����������� ����� �� ��� ������(�������)
* Ap.UnProtected(������)					����� ������ ��������� �����
* Ap.Protect(������)						��������� ������ ��������� �����
* n=Ap.GetCountSheets()						���������� ������ � �����
* str=Ap.GetSheetName(�����)				��� ����� �� ��� ������
* str=Ap.GetActiveSheetName()				��� ��������� �����
* f=Ap.IsActiveSheetProtected() boolean		���������� ������, ���� ���� ������� �������
*     ====������=========
* Ap.MergeCells(������1,�������1,������2,�������2)   ���������� ������ (�������-�������� ���� ��������������)
* Ap.CellsWordWrap(������,�������)			�������� ������� ������ � ������ �� ������
* Ap.SetCellText(������,�������,�����)		������ ������ � ������ �� �����������
*str=Ap.GetCellText(������,�������)=�����	������ ������ �� ������ �� �����������
* Ap.SetNameFont(������,�������,���_������)	���������� ����� � ������
* Ap.SetFontSize(������,�������,������)		���������� ������ ������ � ������
* Ap.SetBold(������,�������)				������� ����� ������ � ������
* Ap.CellsHoriJustifyLeft(������,������)	��������� ����� � ������ �� ����������� �� ����� �������
* Ap.CellsHoriJustifyCenter(������,������)	��������� ����� � ������ �� ����������� �� ������
* Ap.CellsHoriJustifyRight(������,������)	��������� ����� � ������ �� ����������� �� ������ �������
* Ap.CellsVertJustifyCenter(������,�������)	��������� ����� � ������ �� ��������� �� ������
* Ap.CellsVertJustifyTop(������,�������)	��������� ����� � ������ �� ��������� �� �������� ����
*     =====�������=======
* Ap.SetColumnWidth(�������,������)			�������� ������ ������� (100=1��)
* Ap.InsertColByIndex(�������,[����������]) �������� 1(��� ���������) �������� ����� �� ���������� �������
* Ap.DeleteColByIndex(�������,[����������]) ������� 1(��� ���������) ��������
* Ap.ColumnHoriJustifyLeft(�������)			��������� ������ � ������� �� ����� �������
* Ap.ColumnHoriJustifyCenter(�������)		��������� ������ � ������� �� ������
* Ap.ColumnHoriJustifyRight(�������)		��������� ������ � ������� �� ������ �������
*     =====�����/�����=========
* Ap.CellsRightBorder(������,�������)		���������� ����� ������ �� ������ (���� ��� �����)
* Ap.CellsLeftBorder(������,�������)		���������� ����� ����� �� ������ (���� ��� �����)
* Ap.CellsTopBorder(������,�������)			���������� ����� ������ ������ (���� ��� �����)
* Ap.CellsBottomBorder(������,�������)		���������� ����� ����� ������ (���� ��� �����)
* Ap.CellsBorder(������,�������)			���������� ����� ������ ���� ������ (���� ��� �����)
*     ======������=========
* Ap.InsertRowByIndex(������,[����������])  �������� 1(��� ���������) ����� ��� ��������� �������
* Ap.DeleteRowByIndex(������,[����������])  ������� 1(��� ���������) �����
* Ap.SetRowHeight(������,������)			���������� ������ ������ (100=1��)
*             *  ��������  *
* Ap.BorderSize=x  (�����, �� 1 � �����)	��������� ������� ����� ����� ������, �� ����� 1(����� ������������� � ���������� ��������� �����)
* Ap.Visible=�     (����������, .T., .F.)	��������� ��������� ��������� (.T.-�����, .F.-�������)(����������� ����� !)
*==================================================================
Define Class OOEX As Custom

**************************************************
* ����������� �������� Open Office(Libre Office) *
**************************************************
* enum CellHoriJustify specifies how cell contents are aligned horizontally
* h_t_t_p://api.openoffice.org/docs/common/ref/com/sun/star/table/CellHoriJustify.html
	#Define ooCellHoriJustifySTANDARD   0  && default alignment is used (left for numbers, right for text)
	#Define ooCellHoriJustifyLEFT       1  && contents are aligned to the left edge of the cell
	#Define ooCellHoriJustifyCENTER     2  && contents are horizontally centered
	#Define ooCellHoriJustifyRIGHT      3  && contents are aligned to the right edge of the cell
	#Define ooCellHoriJustifyBLOCK      4  && contents are justified to the cell width
	#Define ooCellHoriJustifyREPEAT     5  && contents are repeated to fill the cell

* enum CellVertJustify specifies how cell contents are aligned vertically
* h_t_t_p://api.openoffice.org/docs/common/ref/com/sun/star/table/CellVertJustify.html
	#Define ooCellVertJustifySTANDARD   0  && default alignment is used
	#Define ooCellVertJustifyTOP        1  && contents are aligned with the upper edge of the cell
	#Define ooCellVertJustifyCENTER     2  && contents are aligned to the vertical middle of the cell
	#Define ooCellVertJustifyBOTTOM     3  && contents are aligned to the lower edge of the cell

* constants group DocumentZoomType. These constants specify how the document content is zoomed into the document view.
* h_t_t_p://api.openoffice.org/docs/common/ref/com/sun/star/view/DocumentZoomType.html
* service ViewSettings: h_t_t_p://api.openoffice.org/docs/common/ref/com/sun/star/view/ViewSettings.html#ZoomValue provides access to the settings of the controller of an office document.
	#Define ooDocumentZoomTypeOPTIMAL			0  && The page content width (excluding margins) at the current selection is fit into the view.
	#Define ooDocumentZoomTypePAGE_WIDTH		1  && The page width at the current selection is fit into the view.
	#Define ooDocumentZoomTypeENTIRE_PAGE		2  && A complete page of the document is fit into the view.
	#Define ooDocumentZoomTypeBY_VALUE			3  && The zoom is relative and is to be set via the property ViewSettings::ZoomValue.
	#Define ooDocumentZoomTypePAGE_WIDTH_EXACT	4  && The page width at the current selection is fit into the view, with zhe view ends exactly at the end of the page.

* constants group FontWeight. �������� ��� CharWeight. These values are used to specify whether a font is thin or bold. They may be expanded in future versions.
* h_t_t_p://api.openoffice.org/docs/common/ref/com/sun/star/awt/FontWeight.html
	#Define ooFontWeightDONTKNOW    0.000000  && The font weight is not specified/known.
	#Define ooFontWeightTHIN       50.000000  && specifies a 50% font weight.
	#Define ooFontWeightULTRALIGHT 60.000000  && specifies a 60% font weight.
	#Define ooFontWeightLIGHT      75.000000  && specifies a 75% font weight.
	#Define ooFontWeightSEMILIGHT  90.000000  && specifies a 90% font weight.
	#Define ooFontWeightNORMAL    100.000000  && specifies a normal font weight.
	#Define ooFontWeightSEMIBOLD  110.000000  && specifies a 110% font weight.
	#Define ooFontWeightBOLD      150.000000  && specifies a 150% font weight.
	#Define ooFontWeightULTRABOLD 175.000000  && specifies a 175% font weight.
	#Define ooFontWeightBLACK     200.000000  && specifies a 200% font weight.

*******************************
* ���������� ��������� ������ *
*******************************
	#Define TypePrgNone 0  &&��� ���������: �����������
	#Define TypePrgExcel 1 &&��� ���������: ������
	#Define TypePrgOO    2 &&��� ���������: ��
	#Define TypePrgDefault TypePrgExcel &&���������� �� ��������� (Microsoft Excel)
****************************
* ���������� ������        *
****************************
	TypeOfCalc=TypePrgNone  &&�������� ��� ���������� ����������
&&�� ����� ������������� ������
&&�� ��������� ���������� �� ������
	Programa=Null 			&&������ �� ������ ����������
	Document=Null			&&������ �� ������ �������� � ����������
	ActiveSheet=Null		&&������ �� ������� ���� � ��������� ����������
	oDeskTop=Null			&&������ �� ���������� ������ � �����
&& ��������� �����, ����� ����������������� � �������� ������
	BorderSize=1			&&������� ����� ����� ������
	Visible=.T.				&&��������� ��������� �� ������
******************************************
*      ���������� ������ ������          *
******************************************

* ����������� ���� � ������� �����
	Function FileNameToURL(FileName As String) As String
		Return 'file:///'+Chrtran(FileName,'\','/')
	Endfunc  &&FileNameToURL


* ������������� ������, � �������� �������� ��������
* ��������, ������������, �� ������ ������ ����� �� ����� ������
* ��� ������������� �������� ������� (��������, �� ���������� ����)
* ������ �� ����� ������.
	Function Init(parTypePrg As Integer)
*���� �� ������� ��� ������� �������� ��� ����������
*������� �� ���������
		If Pcount()=0
			This.TypeOfCalc=TypePrgDefault
		Else
			If Vartype(parTypePrg)<>"N"
				This.TypeOfCalc=TypePrgDefault
			Else
				If parTypePrg<>TypePrgExcel And parTypePrg<>TypePrgOO
					This.TypeOfCalc=TypePrgDefault
				Endif
			Endif
		Endif
		If This.TypeOfCalc=TypePrgNone
			This.TypeOfCalc=parTypePrg
		Endif
*������� ����� ���������� �� ������ ���������� ����
		If This.TypeOfCalc=TypePrgExcel
			Try
&&������� ������ ���������� ������
				This.Programa=Createobject('Excel.Application')
				This.Programa.DisplayAlerts=.F.
			Catch
				This.Programa=Null
			Endtry
		Endif
		If This.TypeOfCalc=TypePrgOO
			Try
&&������ ���������� �� �������? ����� ���� ���, ��� ����� �����
				This.Programa=Createobject('com.sun.star.ServiceManager')
			Catch
				This.Programa=Null
			Endtry
		Endif
		If Isnull(This.Programa)
			Return .F.
		Endif
	Endfunc  &&Init

* ����������� �������
	Procedure Destroy
*this.CloseDoc()
*this.CloseProg()
	Endproc  &&Destroy

* �������� ������������� ����������
	Function ProgLoaded() As Boolean
		Return Not(Isnull(This.Programa))
	Endfunc  &&ProgLoaded

* �������� ������������� ���������
	Function DocLoaded() As Boolean
		Return Not(Isnull(This.Document))
	Endfunc  &&DocLoaded

* �������� ����������
	Procedure CloseProg
		If This.DocLoaded()
			This.CloseDoc()
		Endif
		If This.ProgLoaded()
			Try
				If This.TypeOfCalc=TypePrgExcel
					This.Programa.Quit
					This.Programa=Null
				Endif
			Endtry
		Endif
		This.TypeOfCalc=TypePrgNone
	Endproc  &&CloseProg

* ������+������������� ��������� ��������� (��� ����������� ��� ��������� ���������� Visible ������)
	Procedure Visible_Assign
		Lparameters tAssign
		Local aFileProperties[2]
		If Vartype(tAssign)<>'L'		&&�������� ������ ���������� ���
			Return
		Endif
		If This.Visible=tAssign			&&���������� ��� ������ ���� ��������� ��������
			Return
		Endif
		This.Visible=tAssign
&& ��������� ����� ���������
		If This.TypeOfCalc=TypePrgExcel
			Try
				This.Programa.Visible=This.Visible
			Catch
			Endtry
		Endif
		If This.TypeOfCalc=TypePrgOO
			Try
				This.Document.getCurrentController().getFrame().getContainerWindow().SetVisible(This.Visible)
			Catch
			Endtry
		Endif
	Endproc  &&Visible_Assign

******************************************************************************
*  ��������� � ������ ��� ������������� � ���������                          *
******************************************************************************


****************************************
*   ���� ��������                      *
****************************************

* �������� ������� ��������� (������������ �����������)
	Procedure NewDoc()
		Local aFileProperties[2]
		If !This.ProgLoaded()	&&���� ���������� �� �������, �������
			Return
		Endif
		This.CloseDoc()			&&������� �������� ��������
		Desktop=Null
		If This.TypeOfCalc=TypePrgExcel
			This.Programa.WorkBooks.Add()
			This.Programa.Visible=This.Visible
			This.Document=This.Programa.ActiveWorkBook
			This.ActiveSheet=This.Document.ActiveSheet
		Endif
		If This.TypeOfCalc=TypePrgOO
			This.oDeskTop=This.Programa.CreateInstance('com.sun.star.frame.Desktop')
			aFileProperties[1] = This.Programa.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
			aFileProperties[1].Name = "Hidden"
			aFileProperties[1].Value = Not This.Visible
			aFileProperties[2] = This.Programa.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
			aFileProperties[2].Name = "ReadOnly"
			aFileProperties[2].Value = .F.
			This.Document=This.oDeskTop.LoadComponentFromURL('private:factory/scalc','_blank',0,@aFileProperties)
			This.ActivateSheetByIndex(1)
		Endif
	Endproc  &&NewDoc

* �������� ��������� (��� ������ ����������!)
	Procedure CloseDoc
*���� �������� ����������, �� ���������
		If This.DocLoaded()
			Try
				If This.TypeOfCalc=TypePrgExcel
					This.Document.Close
				Endif
				If This.TypeOfCalc=TypePrgOO
					This.Document.Dispose
				Endif
			Catch
			Finally
				This.Document=Null
				This.ActiveSheet=Null
			Endtry
		Endif
	Endproc  &&CloseDoc

* ���������� ��������� (��� ������� �� ���������, ��� �� ��� � ����)
*	����� � ������, ���� �������� ������� ������������ ��������,
*	������ ��������� � ��������� ���� ���������)
* � ������ ��������� ���������� ������������ ������
	Function SaveDoc() As Boolean
		m.Result=.F.
		Try
			If This.DocLoaded()
				If This.TypeOfCalc=TypePrgExcel
					This.Document.Save()
					If This.Document.Saved
						m.Result=.T.
					Endif
				Endif
				If This.TypeOfCalc=TypePrgOO
					This.Document.Store()
					m.Result=.T.
				Endif
			Endif
		Endtry
		Return m.Result
	Endfunc  &&SaveDoc

* ���������� ��������� �� ���������� ���� � �����
	Function SaveAsDoc(NameFile As String) As Boolean
		Local aFileProperties[2]
		m.xlExclusive=3
		xlOtherSessionChanges=3
		m.Result=.F.
		Try
			If This.DocLoaded()
				If This.TypeOfCalc=TypePrgExcel
					This.Document.SaveAs(NameFile,,,,,,m.xlExclusive, m.xlOtherSessionChanges,,,,0)
					If This.Document.Saved
						m.Result=.T.
					Endif
				Endif
				If This.TypeOfCalc=TypePrgOO
					aFileProperties[1] = This.Programa.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
					aFileProperties[1].Name = "Overwrite"
					aFileProperties[1].Value = .T.
					aFileProperties[2] = This.Programa.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
					aFileProperties[2].Name = "FilterName"
					aFileProperties[2].Value = "MS Excel 97"
					This.Document.StoreToUrl(This.FileNameToURL(NameFile),@aFileProperties)
					m.Result=.T.
				Endif
			Endif
		Catch
			MESSAGEBOX("�� ������� ��������� ���� '"+NameFile+"'"+CHR(10)+CHR(13)+;
			"��������� �������:"+CHR(10)+CHR(13)+;
			"1. �� ������ ������ ���� ������������ �����."+CHR(10)+CHR(13)+;
			"2. ������ �������������� ����."+CHR(10)+CHR(13)+;
			"3. � ����� ���� ��� ����� ���������� ������������ �������."+CHR(10)+CHR(13)+;
			"4. ���� ��� ���������� � ����� ������ ���������."+CHR(10)+CHR(13)+;
			"5. ���� ����������, �� � ��� ��� ���� �� ��� ����������.")
		Endtry
		Return m.Result
	Endfunc  &&SaveAsDoc

* ������� ������������ ����
	Procedure LoadDoc(FileName As String)
		Local aFileProperties[2]
		This.CloseDoc()
		If !This.ProgLoaded()
			Return
		Endif
		If Empty(FileName)
			Return
		Endif
		If !File(FileName)
			Return
		Endif
		This.oDeskTop=Null
		If This.TypeOfCalc=TypePrgExcel
			This.Programa.WorkBooks.Add(FileName)
			This.Programa.Visible=This.Visible
			This.Document=This.Programa.ActiveWorkBook
			This.ActiveSheet=This.Document.ActiveSheet
			This.ActivateSheetByIndex(1)
		Endif
		If This.TypeOfCalc=TypePrgOO
			This.oDeskTop=This.Programa.CreateInstance('com.sun.star.frame.Desktop')
			aFileProperties[1] = This.Programa.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
			aFileProperties[1].Name = "Hidden"
			aFileProperties[1].Value = Not This.Visible
			aFileProperties[2] = This.Programa.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
			aFileProperties[2].Name = "ReadOnly"
			aFileProperties[2].Value = .F.
			This.Document=This.oDeskTop.LoadComponentFromURL(This.FileNameToURL(FileName), '_blank', 0,@aFileProperties)
			This.ActivateSheetByIndex(1)
		Endif
	Endproc  &&LoadDoc

****************************************
*   ��������� �����                    *
****************************************

* ������ ����� ������� ��������� ������ �����
* ��� ������������,��������� � ���� �� �����������
	Procedure SetBasePageSetup()
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				With This.ActiveSheet.PageSetup
					.PrintTitleRows=""
					.PrintTitleColumns=""
				Endwith
				With This.ActiveSheet.PageSetup
					.LeftHeader=""
					.CenterHeader=""
					.RightHeader=""
					.LeftFooter=""
					.CenterFooter=""
					.RightFooter=""
					.PrintHeadings=.F.
					.PrintGridlines=.F.
					.Draft=.F.
					.BlackAndWhite=.F.
					.Zoom=.F.
					.FitToPagesWide=1
					.FitToPagesTall=.F.
				Endwith
			Endif
			If This.TypeOfCalc=TypePrgOO
				oPageStyles = This.Document.StyleFamilies.getByName("PageStyles")
				oPageSetup = oPageStyles.getByName(This.ActiveSheet.PageStyle)
				With oPageSetup
					.ScaleToPagesX=1
					.ScaleToPagesY=1000
					.HeaderOn=.F.
					.FooterOn=.F.
					.FooterShared=.F.
				Endwith
			Endif
		Endif
	Endproc  &&SetBasePageSetup


*******************************
*   �����                     *
*******************************

* ���������� ���������� ������ � �������� �����
	Function GetCountSheets() As Integer
		m.Result=0
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				m.Result=This.Document.Sheets.Count
			Endif
			If This.TypeOfCalc=TypePrgOO
				m.Result=This.Document.getSheets.GetCount
			Endif
		Endif
		Return m.Result
	Endfunc  &&GetCountSheets

* ���������� ��� ����� �� ��� ������
	Function GetSheetName(nIndex As Integer) As String
		m.Result=''
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				m.Result=This.Document.Sheets.Item(nIndex).Name
			Endif
			If This.TypeOfCalc=TypePrgOO
				m.Result=This.Document.getSheets.getByIndex(nIndex-1).GetName
			Endif
		Endif
		Return m.Result
	Endfunc  &&GetSheetName

* ���������� ��� ��������� �����
	Function GetActiveSheetName() As String
		m.Result=''
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				m.Result=This.ActiveSheet.Name
			Endif
			If This.TypeOfCalc=TypePrgOO
				m.Result=This.ActiveSheet.GetName
			Endif
		Endif
		Return m.Result
	Endfunc	&&GetActiveSheetName

* ������ ��� ��������� �����
	Procedure SetActiveSheetName(sName As String)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.ActiveSheet.Name=sName
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.SetName(sName)
			Endif
		Endif
	Endproc	&&SetActiveSheetName

* ������ ��� ����� �� ��� ������
	Procedure SetSheetName(nIndex As Integer,sName As String)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.Document.Sheets.Item(nIndex).Name=sName
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.Document.getSheets.getByIndex(nIndex-1).SetName(sName)
			Endif
		Endif
	Endproc  &&SetSheetName

* ����� ������ �������� �����
	Procedure UnProtected(sPassword As String)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.ActiveSheet.UnProtect(sPassword)
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.UnProtect(sPassword)
			Endif
		Endif
	Endproc	&&UnProtected

* ��������� ������ �������� �����
	Procedure Protect(sPassword As String)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.ActiveSheet.Protect(sPassword,Null,Null,Null,Null,0)
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.Protect(sPassword)
			Endif
		Endif
	Endproc	&&UnProtected

* ������� �������� ���� �� ��� ������ (�������)
	Function ActivateSheetByIndex(nIndex As Integer) As Boolean
		m.Result=.F.
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.Document.Sheets(nIndex).Activate		&&���������� ����
				This.ActiveSheet=This.Document.ActiveSheet	&&�������� �� ���� ������
				m.Result=.T.
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet=This.Document.getSheets.getByIndex(nIndex-1)			&&�������� ������ �� ����
				This.Document.getCurrentController().SetActiveSheet(This.ActiveSheet)	&&���������� ���
				m.Result=.T.
			Endif
		Endif
		Return m.Result
	Endfunc  &&ActivateSheetByIndex

* ��������� ������� �� ��������(�������) ���� �������
	Function IsActiveSheetProtected() As Boolean
		m.Result=.F.
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				m.Result=This.ActiveSheet.ProtectContents
			Endif
			If This.TypeOfCalc=TypePrgOO
				m.Result=This.ActiveSheet.IsProtected
			Endif
		Endif
		Return m.Result
	Endfunc	&&IsActiveSheetProtected

**************************************
*           ������                   *
**************************************

* ����������� ����� (���1,���1,���2,���2)
	Procedure MergeCells(row1 As Integer,col1 As Integer,row2 As Integer, col2 As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellRangeByPosition(col1-1,row1-1,col2-1,row2-1).merge(.T.)
			Endif
		Endif
	Endproc  &&MergeCells

* ������� ������� ������ �� ������ � ������ (������, �������)
	Procedure CellsWordWrap(Row As Integer,Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
*!*				This.Programa.ActiveSheet.Cells(Row,Col).Font.Bold=.T.
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).IsTextWrapped=.T.
			Endif
		Endif
	Endproc  &&CellsWordWrap

* �������� ����� � ������ (������,�������,�����)
	Procedure SetCellText(Row As Integer,Col As Integer,txt As String)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.ActiveSheet.Cells(Row,Col).Formula=txt
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).SetString(txt)
			Endif
		Endif
	Endproc  &&SetCellText

* ��������� ����� �� ������ (������,�������)
	Function GetCellText(Row As Integer,Col As Integer) As String
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				Return This.ActiveSheet.Cells(Row,Col).Text  &&Formula
			Endif
			If This.TypeOfCalc=TypePrgOO
				Return This.ActiveSheet.getCellByPosition(Col-1,Row-1).GetString  &&GetFormula
			Endif
		Endif
	Endfunc  &&GetCellText

* ������ ��� ������ � ������ (������,�������,���_������)
	Procedure SetNameFont(Row As Integer,Col As Integer,NameFont As String)
		If Empty(m.NameFont)
			Return
		Endif
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.Programa.ActiveSheet.Cells(Row,Col).Font.Name=m.NameFont
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).getText.createTextCursor.charFontName=m.NameFont
			Endif
		Endif
	Endproc  &&SetNameFont

* ���������� ������ ������ � ������ (������,�������,������)
	Procedure SetFontSize(Row As Integer,Col As Integer,oosize As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.Programa.ActiveSheet.Cells(Row,Col).Font.Size=oosize
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).getText.createTextCursor.CharHeight=oosize
			Endif
		Endif
	Endproc  &&SetFontSize

* ������� ����� ������ � ������ (������,�������)
	Procedure SetBold(Row As Integer,Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.Programa.ActiveSheet.Cells(Row,Col).Font.Bold=.T.
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).getText.createTextCursor.CharWeight=150
			Endif
		Endif
	Endproc  &&SetFontSize

* ��������� � ������ �� ����������� �� ����� ������� (������,�������)
	Procedure CellsHoriJustifyLeft(Row As Integer,Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
* ��� ��� ������
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).HoriJustify=ooCellHoriJustifyLEFT
			Endif
		Endif
	Endproc	&&CellsHoriJustifyLeft

* ��������� � ������ �� ����������� �� ������ (������,�������)
	Procedure CellsHoriJustifyCenter(Row As Integer,Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
* ��� ��� ������
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).HoriJustify=ooCellHoriJustifyCENTER
			Endif
		Endif
	Endproc	&&CellsHoriJustifyCenter

* ��������� � ������ �� ����������� �� ������ ������� (������,�������)
	Procedure CellsHoriJustifyRight(Row As Integer,Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
* ��� ��� ������
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).HoriJustify=ooCellHoriJustifyRIGHT
			Endif
		Endif
	Endproc	&&CellsHoriJustifyRight

* � ������ ��������� ����� �� ��������� �� ������
	Procedure CellsVertJustifyCenter(Row As Integer,Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
* ��� ��� ������
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).VertJustify=ooCellVertJustifyCENTER
			Endif
		Endif
	Endproc  &&CellsVertJustifyCenter

* � ������ ��������� ����� �� ��������� �� �������� ����
	Procedure CellsVertJustifyTop(Row As Integer,Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
* ��� ��� ������
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).VertJustify=ooCellVertJustifyTOP
			Endif
		Endif
	Endproc  &&CellsVertJustifyTop

*********************
*     �������       *
*********************

* �������� ������ �������
	Procedure SetColumnWidth(Col As Integer,Width As Integer)    && 1/100��
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.Programa.ActiveSheet.Cells(1,Col).ColumnWidth=Width/100/3
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,0).getColumns.getByIndex(0).Width=Width
			Endif
		Endif
	Endproc  &&SetColumnWidth

* �������� ������� (����������� ����� �������, ����� �� �������� ����� ��������, � ���-�� �������� (����� �� ���������)
	Procedure InsertColByIndex(Col As Integer,Cnt As Integer)
		Local countCol As Integer
		m.countCol=Iif(Parameters()=2,Cnt,1) &&���� �� ������� ���-�� �����, �� ��������� ����
		If m.countCol<1
			Return
		Endif
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.Columns.InsertByIndex(Col-1,m.countCol)
			Endif
		Endif
	Endproc  &&InsertColByIndex

* ������� ������� (����������� ����� ���������� �������, � ���-�� �������� (����� �� ���������))
	Procedure DeleteColByIndex(Col As Integer,Cnt As Integer)
		Local countCol As Integer
		m.countCol=Iif(Parameters()=2,Cnt,1) &&���� �� ������� ���-�� ��������, �� ������� ����
		If m.countCol<1
			Return
		Endif
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.Columns.RemoveByIndex(Col-1,m.countCol)
			Endif
		Endif
	Endproc  &&DeleteColByIndex

* � ������� ��������� �� ����������� �� ����� ������� ��� ������
	Procedure ColumnHoriJustifyLeft(Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
* ��� ��� ������
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,0).getColumns.getByIndex(0).HoriJustify=ooCellHoriJustifyLEFT
			Endif
		Endif
	Endproc  &&ColumnHoriJustifyLeft

* � ������� ��������� �� ����������� �� ������ ��� ������
	Procedure ColumnHoriJustifyCenter(Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
* ��� ��� ������
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,0).getColumns.getByIndex(0).HoriJustify=ooCellHoriJustifyCENTER
			Endif
		Endif
	Endproc  &&ColumnHoriJustifyCenter

* � ������� ��������� �� ����������� �� ������ ������� ��� ������
	Procedure ColumnHoriJustifyRIGHT(Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
* ��� ��� ������
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,0).getColumns.getByIndex(0).HoriJustify=ooCellHoriJustifyRIGHT
			Endif
		Endif
	Endproc  &&ColumnHoriJustifyRIGHT

******************************
*      �����/�����           *
******************************

* ���������� ����� ������ �� ������
	Procedure CellsRightBorder(Row As Integer,Col As Integer)
		Local oBL As Object   &&������ ����� ��� �����
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
*������ �����
				oBL=This.Programa.Bridge_GetStruct("com.sun.star.table.BorderLine")
				With oBL
					.Color = 0  &&���� ������
					.InnerLineWidth = 0  &&������� ���������� �����
					.OuterLineWidth = This.BorderSize  &&������� ������� �����
					.LineDistance = 0    &&��������� ����� �������
				Endwith
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).RightBorder=oBL  &&����� ������
			Endif
		Endif
	Endproc  &&CellsRightBorder

* ���������� ����� ����� �� ������
	Procedure CellsLeftBorder(Row As Integer,Col As Integer)
		Local oBL As Object   &&������ ����� ��� �����
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
*������ �����
				oBL=This.Programa.Bridge_GetStruct("com.sun.star.table.BorderLine")
				With oBL
					.Color = 0  &&���� ������
					.InnerLineWidth = 0  &&������� ���������� �����
					.OuterLineWidth = This.BorderSize  &&������� ������� �����
					.LineDistance = 0    &&��������� ����� �������
				Endwith
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).LeftBorder=oBL  &&����� ������
			Endif
		Endif
	Endproc  &&CellsLeftBorder

* ���������� ����� ������ ������
	Procedure CellsTopBorder(Row As Integer,Col As Integer)
		Local oBL As Object   &&������ ����� ��� �����
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
*������ �����
				oBL=This.Programa.Bridge_GetStruct("com.sun.star.table.BorderLine")
				With oBL
					.Color = 0  &&���� ������
					.InnerLineWidth = 0  &&������� ���������� �����
					.OuterLineWidth = This.BorderSize  &&������� ������� �����
					.LineDistance = 0    &&��������� ����� �������
				Endwith
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).TopBorder=oBL  &&����� ������
			Endif
		Endif
	Endproc  &&CellsTopBorder

* ���������� ����� ����� ������
	Procedure CellsBottomBorder(Row As Integer,Col As Integer)
		Local oBL As Object   &&������ ����� ��� �����
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
*������ �����
				oBL=This.Programa.Bridge_GetStruct("com.sun.star.table.BorderLine")
				With oBL
					.Color = 0  &&���� ������
					.InnerLineWidth = 0  &&������� ���������� �����
					.OuterLineWidth = This.BorderSize  &&������� ������� �����
					.LineDistance = 0    &&��������� ����� �������
				Endwith
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).BottomBorder=oBL  &&����� ������
			Endif
		Endif
	Endproc  &&CellsBottomBorder

* ���������� ����� ������ ������
	Procedure CellsBorder(Row As Integer,Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
				This.CellsLeftBorder(Row,Col)
				This.CellsTopBorder(Row,Col)
				This.CellsRightBorder(Row,Col)
				This.CellsBottomBorder(Row,Col)
			Endif
		Endif
	Endproc  &&CellsBorder

**************
*   ������   *
**************

* �������� ������ (����������� ����� ������, ��� ������� ����� ���������, � ���-�� ����� (����� �� ���������)
	Procedure InsertRowByIndex(Row As Integer,Cnt As Integer)
		Local countRow As Integer
		m.countRow=Iif(Parameters()=2,Cnt,1) &&���� �� ������� ���-�� �����, �� ��������� ����
		If m.countRow<1
			Return
		Endif
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.Rows.InsertByIndex(Row-1,m.countRow)
			Endif
		Endif
	Endproc  &&InsertRowByIndex

* ������� ������ (����������� ����� ������, � ���-�� ����� (����� �� ���������))
	Procedure DeleteRowByIndex(Row As Integer,Cnt As Integer)
		Local countRow As Integer
		m.countRow=Iif(Parameters()=2,Cnt,1) &&���� �� ������� ���-�� �����, �� ������� ����
		If m.countRow<1
			Return
		Endif
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.Rows.RemoveByIndex(Row-1,m.countRow)
			Endif
		Endif
	Endproc  &&DeleteRowByIndex

* ���������� ������ ������(��������� ����� ������ � �� ������)
	Procedure SetRowHeight(Row As Integer,Height As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.ActiveSheet.Cells(m.Row,1).RowHeight=m.Height/33
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(0,m.Row-1).getRows.getByIndex(0).Height=m.Height
			Endif
		Endif
	Endproc 	&&SetRowHeight

Enddefine
