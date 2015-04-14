Attribute VB_Name = "M01_MainProc"
Option Explicit

'参照設定→MicroSoft Scripting RunTime
'CommentStr
Public cPGList() As C_PROGRAMS

Public cBookName As String

'ModuleType
Public Enum ComponentType
    aNormalModule = 1
    aClassModule = 2
    aUserForm = 3
    aWorkSheet = 100
End Enum

'GetComponentType
Function GetComponentType(aComponentType As Integer)
    Select Case aComponentType
        Case ComponentType.aNormalModule
            GetComponentType = "標準モジュール"
        Case ComponentType.aClassModule
            GetComponentType = "クラスモジュール"
        Case ComponentType.aUserForm
            GetComponentType = "ユーザーフォーム"
        Case ComponentType.aWorkSheet
            GetComponentType = "ワークシート"
    End Select
    
End Function

'MainProc
Sub CreatePSDocument(ctr As IRibbonControl)
    Dim wDocument As Workbook
    Call AppStart
    If Not IsExistsIni Then
        Call CreateIniFile
    End If
    Call IniRead
    FrmPrgsBar.Show vbModeless
    FrmPrgsBar.MainInit "設計書作成中...", 3
    FrmPrgsBar.AddMain
    Call ReadProject
    
    'OutPut
    FrmPrgsBar.AddMain
    Set wDocument = Workbooks.Add
    Call SetPrograms(wDocument.Sheets(1))
    FrmPrgsBar.AddMain
    Call SetModuleContents(wDocument)
    Call AppEnd
    Unload FrmPrgsBar
    MsgBox "全ての処理が終了しました。"
End Sub
'各モジュールの処理です
Sub SetModuleContents(aDocument As Workbook)
    Dim wCnt        As Integer
    Dim wProcName   As String
    Dim wProcList   As Dictionary '<String,Long>
    Dim wModuleStep As Long
    Dim wKey        As Variant
    Dim wCnt2       As Integer
    Dim wShNm       As String
    Dim ws          As Worksheet
    
    For Each ws In aDocument.Worksheets
        If ws.Name = "Sheet2" Then aDocument.Sheets("Sheet2").Delete
        If ws.Name = "Sheet3" Then aDocument.Sheets("Sheet3").Delete
    Next

    FrmPrgsBar.SubInit "プロシージャ一覧作成中...", UBound(cPGList) - 1
    For wCnt = 0 To UBound(cPGList) - 1
        FrmPrgsBar.Add
        aDocument.Sheets.Add After:=aDocument.Sheets(aDocument.Sheets.Count)
        Set wProcList = New Dictionary
        For wModuleStep = cPGList(wCnt).aCodeModule.CountOfDeclarationLines + 1 To cPGList(wCnt).aCodeModule.CountOfLines
            wProcName = cPGList(wCnt).aCodeModule.ProcOfLine(wModuleStep, vbext_pk_Proc)
            On Error Resume Next
            If Not wProcList.Exists(wProcName) Then
                wProcList.Add wProcName, cPGList(wCnt).aCodeModule.ProcBodyLine(wProcName, vbext_pk_Proc)
            End If
        Next wModuleStep
        With ActiveSheet
            wShNm = StrConv(cPGList(wCnt).aNo, vbWide) & "．" & cPGList(wCnt).aProgramName
            If Len(wShNm) >= 32 Then
                wShNm = StrConv(cPGList(wCnt).aNo, vbWide) & "．"
            End If
            If wProcList.Count > 0 Then
                .Name = wShNm
                .Range("A6").Value = .Name
                .Range("B8").Value = "本モジュールのプロシージャ一覧を以下に示す"
                .Range("B11").Value = "No"
                .Range("C11").Value = "属性"
                .Range("D11").Value = "名称"
                .Range("E11").Value = "日本語名称"
                .Range("F11").Value = "備考"
                .Range("B11:F11").Interior.Color = 10092543
                .Columns(1).ColumnWidth = 2
                .Columns(2).ColumnWidth = 5
                .Columns(3).ColumnWidth = 10
                .Columns(4).ColumnWidth = 40
                .Columns(5).ColumnWidth = 60
                .Columns(6).ColumnWidth = 30
                '内容
                wCnt2 = 1
                For Each wKey In wProcList.Keys
                    .Range("B11").Offset(wCnt2, 0).Value = wCnt2
                    .Range("C11").Offset(wCnt2, 0).Value = IIf(Left(cPGList(wCnt).aCodeModule.Lines(wProcList(wKey), 1), InStr(cPGList(wCnt).aCodeModule.Lines(wProcList(wKey), 1), " ") - 1) = "Private", "Private", "Public")
                    .Range("D11").Offset(wCnt2, 0).Value = wKey
                    .Range("E11").Offset(wCnt2, 0).Value = GetProcContent(cPGList(wCnt).aCodeModule, CStr(wKey), wProcList(wKey))
                    wCnt2 = wCnt2 + 1
                Next
                '罫線
                Call BordersSet(.Range("B11").Resize(wCnt2, 5))
            Else
                .Name = wShNm
                .Range("A6").Value = .Name
                .Range("B8").Value = "本モジュールにプロシージャは存在しない。"
                .Columns(1).ColumnWidth = 2
                .Columns(2).ColumnWidth = 5
                .Columns(3).ColumnWidth = 10
                .Columns(4).ColumnWidth = 40
                .Columns(5).ColumnWidth = 60
                .Columns(6).ColumnWidth = 30
            End If
            'フォント
            .Cells.Font.Size = 10
        End With
        'ヘッダー部
        Call HeaderCreate(ActiveSheet, cPGList(wCnt).aNo, cPGList(wCnt).aProgramName)
        'ウィンドウ枠の固定
        ActiveSheet.Rows(5).Select
        ActiveWindow.FreezePanes = True
        ActiveSheet.Range("A1").Select
        'プリントの設定
        Call SetUpPrint(ActiveSheet)
    Next wCnt
End Sub

Sub ReadProject()
    Dim wVbc             As VBComponent
    Dim wNormalList      As New Collection
    Dim wClassList       As New Collection
    Dim wUserFormList    As New Collection
    Dim wWorksheetList   As New Collection
    Dim wBuff            As C_PROGRAMS
    FrmPrgsBar.SubInit "プログラム情報の取得中...", ActiveWorkbook.VBProject.VBComponents.Count + 1
    For Each wVbc In ActiveWorkbook.VBProject.VBComponents
        FrmPrgsBar.Add
        Set wBuff = New C_PROGRAMS
        With wBuff
'            .aProjectName = ActiveWorkbook.Name
            If wVbc.Type = ComponentType.aWorkSheet Then
                .aProgramName = wVbc.Name & "(" & wVbc.Properties("Name") & ")"
            Else
                .aProgramName = wVbc.Name
            End If
            .aProgramDiv = wVbc.Type
            .aProgramContent = Trim(wVbc.CodeModule.Lines(cIniKeyList.aModuleContentRow, cIniKeyList.aModuleContentRow2))
            Set .aCodeModule = wVbc.CodeModule
        End With
        Select Case wVbc.Type
            Case ComponentType.aWorkSheet
                If cIniKeyList.aSheetSelect Then
                    wWorksheetList.Add wBuff
                End If
            Case ComponentType.aNormalModule
                If cIniKeyList.aNormalSelect Then
                    wNormalList.Add wBuff
                End If
            Case ComponentType.aUserForm
                If cIniKeyList.aFrmSelect Then
                    wUserFormList.Add wBuff
                End If
            Case ComponentType.aClassModule
                If cIniKeyList.aClsSelect Then
                    wClassList.Add wBuff
                End If
        End Select
    Next
    Call Marge(wWorksheetList, wNormalList, wUserFormList, wClassList)
    FrmPrgsBar.Add
End Sub

Sub Marge(aWorksheetList As Collection, aNormalList As Collection, aUserFormList As Collection, aClassList As Collection)
    Dim wCnt As Integer
    Dim wNo  As Integer
    ReDim cPGList(0 To aWorksheetList.Count + aNormalList.Count + aUserFormList.Count + aClassList.Count)
    wNo = 1
    For wCnt = 1 To aWorksheetList.Count
        aWorksheetList.Item(wCnt).aNo = wNo
        Set cPGList(wNo - 1) = aWorksheetList.Item(wCnt)
        wNo = wNo + 1
    Next wCnt
    For wCnt = 1 To aNormalList.Count
        aNormalList.Item(wCnt).aNo = wNo
        Set cPGList(wNo - 1) = aNormalList.Item(wCnt)
        wNo = wNo + 1
    Next wCnt
    For wCnt = 1 To aUserFormList.Count
        aUserFormList.Item(wCnt).aNo = wNo
        Set cPGList(wNo - 1) = aUserFormList.Item(wCnt)
        wNo = wNo + 1
    Next wCnt
    For wCnt = 1 To aClassList.Count
        aClassList.Item(wCnt).aNo = wNo
        Set cPGList(wNo - 1) = aClassList.Item(wCnt)
        wNo = wNo + 1
    Next wCnt
'    cPGList(0).aProjectName = ActiveWorkbook.Name
cBookName = ActiveWorkbook.Name
End Sub



'罫線引く
Sub BordersSet(aRng As Range)
    With aRng
        .Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        .Borders(xlEdgeTop).ColorIndex = xlAutomatic
        .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        .Borders(xlEdgeRight).ColorIndex = xlAutomatic
        .Borders(xlInsideVertical).ColorIndex = xlAutomatic
        .Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
    End With
End Sub

'プログラム一覧を作成
Sub SetPrograms(aSh As Worksheet)
    Dim wCnt As Integer
    FrmPrgsBar.SubInit "モジュール一覧作成中...", 2
    FrmPrgsBar.Add
    'ヘッダー
    With aSh
        .Name = "０．モジュール一覧"
        .Range("A6").Value = .Name
        .Range("B8").Value = "本システムのモジュール一覧を以下に示す"
        .Range("B11").Value = "章番"
        .Range("C11").Value = "ブック名"
        .Range("D11").Value = "モジュール名"
        .Range("E11").Value = "モジュール区分"
        .Range("F11").Value = "モジュール概要"
        .Range("G11").Value = "備考"
        .Range("B11:G11").Interior.Color = 10092543
        .Columns(1).ColumnWidth = 2
        .Columns(2).ColumnWidth = 4
        .Columns(3).ColumnWidth = 20
        .Columns(4).ColumnWidth = 30
        .Columns(5).ColumnWidth = 20
        .Columns(6).ColumnWidth = 50
        .Columns(7).ColumnWidth = 20
        '内容
        .Range("C12").Value = cBookName
        For wCnt = 0 To UBound(cPGList) - 1
            .Range("B12").Offset(wCnt, 0).Value = cPGList(wCnt).aNo
'            .Range("C12").Offset(wCnt, 0).Value = cPGList(wCnt).aProjectName
            .Range("D12").Offset(wCnt, 0).Value = cPGList(wCnt).aProgramName
            .Range("E12").Offset(wCnt, 0).Value = GetComponentType(cPGList(wCnt).aProgramDiv)
            .Range("F12").Offset(wCnt, 0).Value = IIf(cIniKeyList.aModuleContentNotExist, "", GetProgramContent(cPGList(wCnt).aProgramContent))
        Next wCnt
        '罫線
        Call BordersSet(.Range("B11").Resize(wCnt + 1, 6))
        'フォント
        .Cells.Font.Size = 10
        'ヘッダー部
        Call HeaderCreate(aSh, 0, "モジュール一覧")
        'ウィンドウ枠の固定
        .Rows(5).Select
        ActiveWindow.FreezePanes = True
        .Range("A1").Select
        'プリントの設定
        Call SetUpPrint(aSh)
    End With
    FrmPrgsBar.Add
End Sub

Public Function GetProgramContent(aContentStr As String) As String
    Dim wBuff() As String
    Dim wCnt    As Integer
    Dim wResult As String
    wResult = aContentStr
    wBuff() = Split(cIniKeyList.aModuleRemComment, ",")
    For wCnt = 0 To UBound(wBuff)
        wResult = Replace(wResult, wBuff(wCnt), "")
    Next wCnt
    GetProgramContent = Trim(wResult)
End Function

'Option
Sub ViewFrmOption(ctr As IRibbonControl)
    FrmOption.Show
End Sub

'ヘッダー部を作成する
Public Sub HeaderCreate(aSh As Worksheet, aSeqNum As Integer, aSeqName As String)
    '--ブック名
    Const BOOK_NAME_OBJ_NAME          As String = "Txt_Book_Name"
    Const BOOK_NAME_OBJ_WIDTH         As Double = 185
    Const BOOK_NAME_OBJ_HEIGHT        As Double = 38
    Const BOOK_NAME_OBJ_LEFT          As Double = 0.5
    Const BOOK_NAME_OBJ_TOP           As Double = 0.5
    '--シート名
    Const SHEET_NAME_OBJ_NAME         As String = "Txt_Sheet_Name"
    Const SHEET_NAME_OBJ_WIDTH        As Double = 370
    Const SHEET_NAME_OBJ_HEIGHT       As Double = BOOK_NAME_OBJ_HEIGHT
    Const SHEET_NAME_OBJ_LEFT         As Double = BOOK_NAME_OBJ_WIDTH
    Const SHEET_NAME_OBJ_TOP          As Double = BOOK_NAME_OBJ_TOP
    '--作成者
    Const AUTHER_NAME_OBJ_NAME        As String = "Txt_Auther_Name"
    Const AUTHER_NAME_OBJ_WIDTH       As Double = 85
    Const AUTHER_NAME_OBJ_HEIGHT      As Double = BOOK_NAME_OBJ_HEIGHT
    Const AUTHER_NAME_OBJ_LEFT        As Double = BOOK_NAME_OBJ_WIDTH + SHEET_NAME_OBJ_WIDTH
    Const AUTHER_NAME_OBJ_TOP         As Double = BOOK_NAME_OBJ_TOP
    '--作成日
    Const CREATE_DATE_OBJ_NAME        As String = "Txt_Create_Date"
    Const CREATE_DATE_OBJ_WIDTH       As Double = AUTHER_NAME_OBJ_WIDTH
    Const CREATE_DATE_OBJ_HEIGHT      As Double = BOOK_NAME_OBJ_HEIGHT
    Const CREATE_DATE_OBJ_LEFT        As Double = BOOK_NAME_OBJ_WIDTH + SHEET_NAME_OBJ_WIDTH + AUTHER_NAME_OBJ_WIDTH
    Const CREATE_DATE_OBJ_TOP         As Double = BOOK_NAME_OBJ_TOP
    '--更新者
    Const UPDATER_NAME_OBJ_NAME       As String = "Txt_Updater_Name"
    Const UPDATER_NAME_OBJ_WIDTH      As Double = AUTHER_NAME_OBJ_WIDTH
    Const UPDATER_NAME_OBJ_HEIGHT     As Double = BOOK_NAME_OBJ_HEIGHT
    Const UPDATER_NAME_OBJ_LEFT       As Double = BOOK_NAME_OBJ_WIDTH + SHEET_NAME_OBJ_WIDTH + AUTHER_NAME_OBJ_WIDTH + CREATE_DATE_OBJ_WIDTH
    Const UPDATER_NAME_OBJ_TOP        As Double = BOOK_NAME_OBJ_TOP
    '--更新日
    Const UPDATE_DATE_OBJ_NAME        As String = "Txt_Update_Date"
    Const UPDATE_DATE_OBJ_WIDTH       As Double = AUTHER_NAME_OBJ_WIDTH
    Const UPDATE_DATE_OBJ_HEIGHT      As Double = BOOK_NAME_OBJ_HEIGHT
    Const UPDATE_DATE_OBJ_LEFT        As Double = BOOK_NAME_OBJ_WIDTH + SHEET_NAME_OBJ_WIDTH + AUTHER_NAME_OBJ_WIDTH + CREATE_DATE_OBJ_WIDTH + UPDATER_NAME_OBJ_WIDTH
    Const UPDATE_DATE_OBJ_TOP         As Double = BOOK_NAME_OBJ_TOP
    '--ブック名
    With aSh.Shapes.AddShape(msoShapeRectangle, BOOK_NAME_OBJ_LEFT, BOOK_NAME_OBJ_TOP, BOOK_NAME_OBJ_WIDTH, BOOK_NAME_OBJ_HEIGHT)
        .Name = BOOK_NAME_OBJ_NAME
        .TextFrame.Characters.Text = cBookName
        .TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        .TextFrame.Characters.Font.Size = 10.5
        .TextFrame.Characters.Font.Bold = False
        .Fill.Visible = -1
        .Line.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = 2
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Line.Weight = 0.75
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame.Characters.Font.Color = 0
        .TextFrame2.TextRange.Font.NameFarEast = "+mj-ea"
        .TextFrame2.TextRange.Font.Name = "+mj-ea"
        .Line.ForeColor.ObjectThemeColor = msoThemeColorText1
        .Line.ForeColor.TintAndShade = 0
        .Line.ForeColor.Brightness = 0.0500000007
        .Line.Transparency = 0
    End With
    '--シート名
    With aSh.Shapes.AddShape(msoShapeRectangle, SHEET_NAME_OBJ_LEFT, AUTHER_NAME_OBJ_TOP, SHEET_NAME_OBJ_WIDTH, SHEET_NAME_OBJ_HEIGHT)
        .Name = SHEET_NAME_OBJ_NAME
        .TextFrame.Characters.Text = "第" & aSeqNum & "章" & vbNewLine & aSeqName
        .TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        .TextFrame.Characters.Font.Size = 10.5
        .TextFrame.Characters.Font.Bold = False
        .Fill.Visible = -1
        .Line.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = 2
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Line.Weight = 0.75
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame.Characters.Font.Color = 0
        .TextFrame2.TextRange.Font.NameFarEast = "+mj-ea"
        .TextFrame2.TextRange.Font.Name = "+mj-ea"
        .Line.ForeColor.ObjectThemeColor = msoThemeColorText1
        .Line.ForeColor.TintAndShade = 0
        .Line.ForeColor.Brightness = 0.0500000007
        .Line.Transparency = 0
    End With
    '--作成者
    With aSh.Shapes.AddShape(msoShapeRectangle, AUTHER_NAME_OBJ_LEFT, SHEET_NAME_OBJ_TOP, AUTHER_NAME_OBJ_WIDTH, AUTHER_NAME_OBJ_HEIGHT)
        .Name = AUTHER_NAME_OBJ_NAME
        .TextFrame.Characters.Text = "作成者" & vbNewLine & IIf(cIniKeyList.aAcnSelect, CreateObject("WScript.NetWork").UserName, cIniKeyList.aAutName)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        .TextFrame.Characters.Font.Size = 10.5
        .TextFrame.Characters.Font.Bold = False
        .Fill.Visible = -1
        .Line.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = 2
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Line.Weight = 0.75
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame.Characters.Font.Color = 0
        .TextFrame2.TextRange.Font.NameFarEast = "+mj-ea"
        .TextFrame2.TextRange.Font.Name = "+mj-ea"
        .Line.ForeColor.ObjectThemeColor = msoThemeColorText1
        .Line.ForeColor.TintAndShade = 0
        .Line.ForeColor.Brightness = 0.0500000007
        .Line.Transparency = 0
    End With
    '--作成日
    With aSh.Shapes.AddShape(msoShapeRectangle, CREATE_DATE_OBJ_LEFT, CREATE_DATE_OBJ_TOP, CREATE_DATE_OBJ_WIDTH, CREATE_DATE_OBJ_HEIGHT)
        .Name = CREATE_DATE_OBJ_NAME
        .TextFrame.Characters.Text = "作成日" & vbNewLine & IIf(cIniKeyList.aNowSelect, Date, cIniKeyList.aCreDate)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        .TextFrame.Characters.Font.Size = 10.5
        .TextFrame.Characters.Font.Bold = False
        .Fill.Visible = -1
        .Line.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = 2
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Line.Weight = 0.75
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame.Characters.Font.Color = 0
        .TextFrame2.TextRange.Font.NameFarEast = "+mj-ea"
        .TextFrame2.TextRange.Font.Name = "+mj-ea"
        .Line.ForeColor.ObjectThemeColor = msoThemeColorText1
        .Line.ForeColor.TintAndShade = 0
        .Line.ForeColor.Brightness = 0.0500000007
        .Line.Transparency = 0
    End With
    '--更新者
    With aSh.Shapes.AddShape(msoShapeRectangle, UPDATER_NAME_OBJ_LEFT, UPDATER_NAME_OBJ_TOP, UPDATE_DATE_OBJ_WIDTH, UPDATER_NAME_OBJ_HEIGHT)
        .Name = UPDATER_NAME_OBJ_NAME
        .TextFrame.Characters.Text = "更新者"
        .TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        .TextFrame.Characters.Font.Size = 10.5
        .TextFrame.Characters.Font.Bold = False
        .Fill.Visible = -1
        .Line.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = 2
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Line.Weight = 0.75
        .TextFrame2.VerticalAnchor = msoAnchorTop
        .TextFrame.Characters.Font.Color = 0
        .TextFrame2.TextRange.Font.NameFarEast = "+mj-ea"
        .TextFrame2.TextRange.Font.Name = "+mj-ea"
        .Line.ForeColor.ObjectThemeColor = msoThemeColorText1
        .Line.ForeColor.TintAndShade = 0
        .Line.ForeColor.Brightness = 0.0500000007
        .Line.Transparency = 0
    End With
    '--更新日
    With aSh.Shapes.AddShape(msoShapeRectangle, UPDATE_DATE_OBJ_LEFT, UPDATE_DATE_OBJ_TOP, UPDATE_DATE_OBJ_WIDTH, UPDATE_DATE_OBJ_HEIGHT)
        .Name = UPDATE_DATE_OBJ_NAME
        .TextFrame.Characters.Text = "更新日"
        .TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        .TextFrame.Characters.Font.Size = 10.5
        .TextFrame.Characters.Font.Bold = False
        .Fill.Visible = -1
        .Line.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = 2
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Line.Weight = 0.75
        .TextFrame2.VerticalAnchor = msoAnchorTop
        .TextFrame.Characters.Font.Color = 0
        .TextFrame2.TextRange.Font.NameFarEast = "+mj-ea"
        .TextFrame2.TextRange.Font.Name = "+mj-ea"
        .Line.ForeColor.ObjectThemeColor = msoThemeColorText1
        .Line.ForeColor.TintAndShade = 0
        .Line.ForeColor.Brightness = 0.0500000007
        .Line.Transparency = 0
    End With
    
End Sub

Sub SetUpPrint(aSh As Worksheet)
    With aSh.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = "&P ページ"
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.708661417322835)
        .RightMargin = Application.InchesToPoints(0.708661417322835)
        .TopMargin = Application.InchesToPoints(0.748031496062992)
        .BottomMargin = Application.InchesToPoints(0.748031496062992)
        .HeaderMargin = Application.InchesToPoints(0.31496062992126)
        .FooterMargin = Application.InchesToPoints(0.31496062992126)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
End Sub

Sub AppStart()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
End Sub
Sub AppEnd()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Function GetProcContent(aCodeModule As CodeModule, aProcName As String, aProcBodyLine As Long) As String
    Dim wProcStartLine As Long
    Dim wProcCntent    As String
    Dim wNewLine       As Long
    Dim i              As Long
    Dim wBuff()        As String
    Dim wCnt           As Integer
    If cIniKeyList.aProcContentNotExist Then
        GetProcContent = ""
    Else
        wProcStartLine = aCodeModule.ProcStartLine(aProcName, vbext_pk_Proc)
        wNewLine = -1
        '--行で指定
        If cIniKeyList.aProcOptWhere = OptModuleType.aOptRow Then
            wNewLine = aProcBodyLine - cIniKeyList.aProcContentRow
            'コメントで指定
        ElseIf cIniKeyList.aProcOptWhere = OptModuleType.aOptCom Then
            For i = wProcStartLine To aProcBodyLine
                If InStr(aCodeModule.Lines(i, 1), cIniKeyList.aProcContent) > 0 Then
                    wNewLine = i
                    Exit For
                End If
            Next i
        End If
        '概要取得
        If wNewLine >= 0 Then
            wProcCntent = aCodeModule.Lines(wNewLine, cIniKeyList.aProcContentRow2)
        End If
        wBuff() = Split(cIniKeyList.aProcRemComment, ",")
        For wCnt = 0 To UBound(wBuff)
            wProcCntent = Replace(wProcCntent, wBuff(wCnt), "")
        Next wCnt
        GetProcContent = Trim(wProcCntent)
    End If
End Function
