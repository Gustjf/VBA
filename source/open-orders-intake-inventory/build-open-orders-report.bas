Attribute VB_Name = "modOpenOrdersReport"
Option Explicit

' =====================================================================================================
'  BUILD OPEN ORDERS - ORDER INTAKE - INVENTORY ANALYSIS BY MSA                              Version 2.8
'  -----------------------------------------------------------------------------------------------------
'  Reads the Snowflake exports for all MSAs from one folder
'       "Order Intake*.xlsx"        (required)  -> tab Order Intake (every city in one list)
'       "Inventory by Month*.xlsx"  (optional)  -> tab Inventory by Month (units by city, then the detail)
'       "Open Order*.xlsx"          (optional)  -> read and logged only ("Open Orders" is accepted too)
'  plus the size class workbook (Config SizeClassFile, the master list of models and size classes), and
'  writes one workbook "001. Open Orders - Order Intake - Inventory Analysis By MSA.xlsx" with the tabs
'       File Format         how the report is laid out and which rules and files built it
'       Order Intake        every order line of every MSA, size class in column A, MSA in column B, and in the
'                           title row the total ordered quantity of the rows the filters show
'       Inventory by Month  units in stock by MSA and product type for one month, then the inventory lines
'       Open Orders by MSA  ordered quantity of the open orders by MSA and product type, and a chart of it
'       Run Log             what was read, written, dropped, hidden and could not be placed
'
'  HOW TO RUN  (full instructions: output/open-orders-intake-inventory/macro-instructions.md)
'       1. Import this module into a macro-enabled workbook (.xlsm)  (VBA editor > File > Import File).
'          Not into PERSONAL.XLSB: the Config sheet would be created in a hidden workbook.
'          Keep this file with Windows (CRLF) line endings; the VBA editor's Import File expects them.
'       2. Run  ShowOpenOrdersReportConfig  once. It creates the Config sheet (the control panel: settings
'          with drop-downs, buttons Build report / Check setup / Choose input folder / Choose output folder /
'          Choose size class file, a Last run box and help). Run it again after importing a new version: it
'          upgrades an existing Config sheet in place (a version 1.x sheet too) and keeps its values.
'       3. Put the exports in one folder, set SizeClassFile, then click Build report on the Config sheet.
'          Every build first checks the whole setup (the same check as the Check setup button) and stops with
'          one list of problems, before anything is created, if the settings, folders or files are not usable.
'       The Config sheet is the only way to build: in the Macros list (Alt+F8) only ShowOpenOrdersReportConfig
'       appears. The five button macros take an unused Optional argument, which keeps them out of that list.
'       Config is protected without a password (Review > Unprotect Sheet) so only the Value cells can change.
'
'  BUSINESS RULES (the manager's decisions of 2026-10-12; plan Section 5)
'       * Every file of a kind is read and merged. The MSA of a row comes from the MSA column (Config
'         MSAColumn, else the first header equal to MSA, containing MSA, equal to CITY, or containing CITY, METRO
'         or MARKET), else from the file name. MSA text is trimmed, spaces collapsed, compared ignoring case;
'         blank = (no MSA). In the code an MSA is called a "city" (the variable names of version 2.0).
'       * Order Intake keeps PRODUCT_TYPE_ABBR CTL, MX and SSL (written trimmed and upper-cased); other types
'         are dropped and counted. CANCELLED orders and SSL rows are written and hidden by AutoFilter criteria
'         unless Config ShowCancelled / ShowSSL is TRUE.
'       * Size Class from MODEL_GROUP via the size class file (exact name, then the name's first word), the only
'         source of a class. The allowed classes (v2.8) are the three tables of its Set-up tab (CTL size class,
'         MX size class, SSL size class), in that order, top to bottom; without a Set-up tab, the 17 built in (CTL
'         by rated operating capacity, MX by nominal tonnes, SSL by frame size). A model not in the file is
'         UNMAPPED (bold, peach).
'       * Sorted by size class, order status (ORDERED, Not yet processed, INVOICED, Completely processed,
'         other, CANCELLED), ORDER_DATE newest first (date part, undated last), MSA, then export order.
'       * Visible total (v2.6): the frozen title row of Order Intake holds "Total sum of ordered Quantity of visible
'         selection:" (F1) and =SUBTOTAL(109,...) over the whole SUM(F.ORDERED_QUANTITY) column (G1, above that
'         column), so the figure follows whatever filters the reader sets; the tab's note moves to H1.
'       * Inventory keeps Compact Track Loaders, Mini Excavators and Skid-Steer Loaders. Config InventoryMonth
'         (any day in the month; blank = the latest month in the files) is the month the "Units in" table
'         shows (live SUMIFS formulas of TOTAL over the whole detail block, by MSA and product type: the old
'         CTL= and MX= figures) and the detail rows are filtered to. ORGANIZATION_NAME and
'         ORGANIZATION_CUSTOMER_NUMBER close the detail (blank if an export lacks them).
'       * Open Orders by MSA (v2.5): the ordered quantity (sum of SUM(F.ORDERED_QUANTITY)) of the Order Intake rows
'         whose ORDER_STATUS is ORDERED or Not yet processed, compared as SUMIFS compares text (the whole text,
'         ignoring case, not ignoring stray spaces), by MSA and product type: live SUMIFS formulas over the Order
'         Intake tab, then a clustered column chart of the table (one bar per MSA in each product type group; SSL
'         on the chart only with ShowSSL TRUE). The Run Log gives each MSA's open rows and warns about statuses
'         with stray spaces (plan: open-orders-by-msa-tab-implementation-plan.md).
'       * Dates stored as text are converted when they read as a date; anything else is left as it is and
'         its rows are listed on the Run Log. Blank rows are ignored: rows read = written + dropped + blank.
'       * Style: Cambria, horizontal rules only, green headers, rows banded light green and white, printed
'         landscape one page wide (STYLE SHEET constants below).
'
'  build_report_reference.py (Python) applies the same rules and writes the same Run Log lines, and
'  check_report.py checks a finished report against them. A rule changed here must be changed there too.
'  Requires Excel 2010 or later. Exports must be .xlsx / .xlsm / .xls (save CSV exports as .xlsx first:
'  CSV date parsing is locale dependent).
' =====================================================================================================

Private Const BUILDER_TEXT As String = "Open Orders report macro v2.8"
Private Const CFG_SHEET As String = "Config"
Private Const OLD_MAP_SHEET As String = "SizeClassMap"      ' version 1.x; no longer used
Private Const DEFAULT_OUTPUT_NAME As String = "001. Open Orders - Order Intake - Inventory Analysis By MSA.xlsx"
Private Const BUTTON_NAME As String = "btnBuildOpenOrdersReport"
Private Const BTN_CHECK As String = "btnCheckOpenOrdersSetup"
Private Const BTN_INPUT As String = "btnChooseOpenOrdersInput"
Private Const BTN_OUTPUT As String = "btnChooseOpenOrdersOutput"
Private Const BTN_SIZE As String = "btnChooseOpenOrdersSizeClass"
Private Const BUTTON_ROW As Long = 12           ' Config: row of the buttons
Private Const STATUS_ROW As Long = 14           ' Config: "Last run" box (title, Result, When, Last report, Cities, Details)
Private Const HELP_ROW As Long = 21             ' Config: "If something goes wrong" box
Private Const MAX_LISTED As Long = 20           ' problems / warnings listed in the Last run box
Private Const MSG_LISTED As Long = 8            ' problems / warnings listed in a message box
Private Const STATUS_PREFIX As String = "Open Orders report: "
Private Const MAX_ROWS_LISTED As Long = 10      ' row numbers listed per Run Log line
Private Const NO_CITY As String = "(no MSA)"
' File Format Table 1: the family and rule text of each block of size classes (v2.7: three families)
Private Const CTL_FAMILY As String = "Track loaders (CTL)"
Private Const MX_FAMILY As String = "Mini excavators (MX)"
Private Const SSL_FAMILY As String = "Skid-steer loaders (SSL)"
Private Const CTL_RULE_TEXT As String = "by rated operating capacity in lb; the class is read from the size class file"
Private Const MX_RULE_TEXT As String = "by nominal tonnes; the class is read from the size class file"
Private Const SSL_RULE_TEXT As String = "by frame size; the class is read from the size class file"

' File Format tab: Table 1 starts in a fixed row and has a row per size class (v2.8), so the tables under it move:
' Table 2 two rows under its UNMAPPED row (row 24 with the 17 built-in classes), Tables 3 and 4 each FF_TABLE_GAP rows
' further (build_report_reference.py does the same)
Private Const FF_CLASS_ROW As Long = 3
Private Const FF_TABLE_GAP As Long = 9
' size classes (v2.8): UNMAPPED, and anything not allowed, ranks 99, so at most 98 classes
Private Const UNMAPPED_RANK As Long = 99
Private Const MAX_CLASSES As Long = 98
Private Const FF_MONTH_COL As Long = 6          ' F: "Months in this build", the list the month cell points at

' Config sheet colours (unchanged from version 1.x)
Private Const CLR_YELLOW As Long = 65535        ' FFFF00  Config header and box titles
Private Const CLR_INPUT As Long = 13434879      ' FFFFCC  Config value cells; the report's month cell
Private Const CLR_OK As Long = 13561798         ' C6EFCE  Last run: built / ready
Private Const CLR_WARN As Long = 10284031       ' FFEB9C  Last run: built / ready, with warnings
Private Const CLR_BAD As Long = 13551615        ' FFC7CE  Last run: not built / failed
Private Const CLR_BOX As Long = 15921906        ' F2F2F2  Config: labels of the Last run and help boxes
' STYLE SHEET of the report (v2.4): set like an engineering report, in green. Cambria (ships with Office),
' horizontal black rules only (no vertical lines), bold white headers on green with a medium rule above and a
' double rule under them, a medium rule between groups, every second visible row light green (conditional
' formatting, so the bands follow the filters), UNMAPPED Size Class cells peach, the month cell yellow (CLR_INPUT);
' the Run Log in Consolas. Printed landscape, one page wide, with a page header and footer.
Private Const FONT_REPORT As String = "Cambria"
Private Const FONT_LOG As String = "Consolas"
Private Const BODY_SIZE As Long = 10
Private Const HEAD_HEIGHT As Single = 20         ' header rows
Private Const ROW_HEIGHT As Single = 16          ' table body rows
Private Const HEAD_FILL As Long = 3506772        ' 548235  header rows
Private Const HEAD_TEXT As Long = 16777215       ' FFFFFF  header text
Private Const BAND_FILL As Long = 14348258       ' E2EFDA  every second visible body row
Private Const CLR_FLAG As Long = 14083324        ' FCE4D6  UNMAPPED Size Class cells
Private Const REPORT_TITLE As String = "Open Orders - Order Intake - Inventory Analysis By MSA"   ' page header
Private Const HF_FONT As String = "&""Cambria,Regular""&9"
' Column alignment, one letter per column: L left, R right (dates, numbers), C centre (short values). L and R
' cells are indented one step, the cell padding of a printed table.
Private Const OI_ALIGN As String = "LLRRLCRCLLL"
' Order Intake title row (v2.6): the visible total's label right-aligned in F1, its figure left-aligned in G1 (above
' SUM(F.ORDERED_QUANTITY)), so the two read as one line, and the tab's note in H1 instead of D1, so the label
' does not cut it off
Private Const OI_TOTAL_LABEL As String = "Total sum of ordered Quantity of visible selection:"
Private Const OI_NOTE_COL As Long = 8
Private Const SUM_ALIGN As String = "LRRRR"
Private Const INV_ALIGN As String = "LLLLRLRLR"

Private Type SourceTable
    Ok As Boolean
    Raw As Variant                ' 2-D Variant straight from UsedRange.Value
    FirstRow As Long              ' worksheet row of Raw(1, *), for row numbers in the Run Log
    HdrRow As Long                ' row of Raw holding the column headers
    LastRow As Long               ' last non-blank row of Raw
    NRead As Long                 ' rows below the header up to the last non-blank one (blank rows included)
    DataRows As Long              ' non-blank rows below the header
    ColCount As Long
    Headers() As String           ' 1-based, upper-cased, trimmed
    SetupFound As Boolean         ' v2.8, size class file only: the workbook has a Set-up tab
    SetupRaw As Variant           ' its UsedRange.Value (2-D, 1-based)
    SetupRow As Long              ' worksheet row and column of SetupRaw(1, 1), for cell addresses in warnings
    SetupCol As Long
End Type

Private Type InputFile
    Path As String
    FileName As String
    Kind As String                ' OI | INV | OO
    FileCity As String            ' city from the file name (used when there is no city column)
    Tbl As SourceTable
    Missing As String             ' required columns not found ("" = usable)
    Idx() As Long                 ' column of each required header (0-based like the header list)
    CityCol As Long               ' the MSA column; 0 = MSA from the file name
    CityHdr As String
    OrgCol As Long                ' Inventory: ORGANIZATION_NAME and ORGANIZATION_CUSTOMER_NUMBER (0 = not in the file)
    CustCol As Long
    Written As Long
    BlankCity As Long
    CityKeys As String            ' distinct cities of this file, Chr(1)-delimited
    NCities As Long
    Conv(0 To 1) As Long          ' text dates converted / not converted, per date column
    NBad(0 To 1) As Long
    BadRows(0 To 1) As String
End Type

Private Type ReportOptions
    InputFolder As String
    OutputFolder As String
    OutputFileName As String
    SizeClassFile As String
    CityColumn As String
    InvMonth As Long              ' year * 12 + month - 1; -1 = the latest month in the files
    ShowCancelled As Boolean
    ShowSSL As Boolean
    OpenAfterBuild As Boolean
End Type

Private mLog As Collection
Private mStep As String
Private mNonIsoText As Boolean    ' a text date not in yyyy-mm-dd form was converted (read in the PC's date order)
Private mUpgradeNote As String    ' set when ShowOpenOrdersReportConfig or a button upgraded a version 1.x Config
' the input files, in alphabetical order of file name
Private mFiles() As InputFile
Private mNFiles As Long
Private mReady As Boolean         ' the files were read and the rows prepared
' the size class file
Private mScKeys() As String, mScFirst() As String, mScClass() As String, mScCount As Long
Private mScName As String, mScWarnings As Collection
' the allowed size classes in rank order (v2.8): name as written, family (1 CTL, 2 MX, 3 SSL), and the Run Log line
Private mClsName() As String, mClsFam() As Long, mClsCount As Long, mClsLine As String
' cities: registry in the order first seen (key = upper case, name = first spelling), then alphabetical order
Private mCityKeys() As String, mCityNames() As String, mNCities As Long, mLastCity As Long
Private mCityOrder() As Long, mCityRank() As Long
Private mUnmapped() As String, mUnmappedAll As String
' rows of another product type dropped: file, type label (as exported) and count, in the order first seen
Private mDropFile() As Long, mDropLabel() As String, mDropCount() As Long, mNDrop As Long
' Order Intake rows (A..K; column B holds the city's registry index until written) and their sort data
Private mOI() As Variant, mNOI As Long, mOICity() As Long, mOIStatus() As Long, mOIDay() As Double, mOIOrd() As Long
' Inventory rows (PRODUCT_TYPE, MODEL_GROUP, MODEL_NUMBER, INVENTORY_DATE, TOTAL) and their sort data
Private mINV() As Variant, mNINV As Long, mINVCity() As Long, mINVType() As Long, mINVMonth() As Long
Private mINVDay() As Double, mINVOrd() As Long
' months: the one shown (-1 = none) and every month in the build, ascending
Private mMonth As Long, mMonths() As Long, mNMonths As Long
' open orders (ORDER_STATUS ORDERED or Not yet processed) over all MSAs: rows, and units by product type (CTL, MX, SSL)
Private mOpenRows As Long, mOpenUnits(0 To 2) As Double

' =====================================================================================================
'  PUBLIC ENTRY POINTS
'  ShowOpenOrdersReportConfig is the only macro in the Macros list (Alt+F8): it sets up the Config sheet.
'  The other five are the Config buttons. Their unused Optional argument keeps them out of the Macros list
'  (Excel lists only procedures without arguments), so the report is built from the Config sheet only; the
'  buttons, and code such as Application.Run, still call them without an argument.
' =====================================================================================================

' Creates (or upgrades) the Config sheet and lays out the control panel: checked Value cells with drop-downs,
' the buttons, the Last run box, the help box, and sheet protection. Run it once after importing, and again
' after importing a new version (values already set are kept; a version 1.x sheet is converted).
Public Sub ShowOpenOrdersReportConfig()
    mUpgradeNote = ""
    EnsureConfigSheet
    SetUpConfigPanel
    If Len(mUpgradeNote) > 0 Then WriteLastRun mUpgradeNote, CLR_WARN, "-", mUpgradeNote
    If ThisWorkbook.Windows.Count > 0 Then
        If ThisWorkbook.Windows(1).Visible Then
            ThisWorkbook.Activate
            ThisWorkbook.Worksheets(CFG_SHEET).Activate
        End If
    End If
    MsgBox "The Config sheet is ready. Fill in the yellow Value cells (Choose input folder and Choose size class file " & _
           "set the paths), click Check setup, then Build report." & IIf(Len(mUpgradeNote) > 0, vbCrLf & vbCrLf & mUpgradeNote, ""), _
           vbInformation, "Open Orders report"
End Sub

' Button "Build report": checks the setup, then builds the report workbook from the files in InputFolder.
Public Sub BuildOpenOrdersReport(Optional ByVal notInMacroList As Variant)
    Dim opt As ReportOptions
    Dim wbOut As Workbook, wsFmt As Worksheet, wsOI As Worksheet, wsInv As Worksheet, wsOO As Worksheet, wsLog As Worksheet
    Dim outPath As String, t0 As Single, builtAt As Date, i As Long
    Dim problems As Collection, warnings As Collection, toCheck As Collection, cancelled As Boolean
    Dim result As String, details As String, nItems As Long

    Dim prevScreen As Boolean, prevEvents As Boolean, prevAlerts As Boolean, prevCalc As XlCalculation
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    prevCalc = Application.Calculation

    On Error GoTo Fail
    Application.EnableCancelKey = xlErrorHandler     ' Ctrl+Break also restores the settings below
    t0 = Timer
    Set mLog = New Collection
    Set problems = New Collection
    Set warnings = New Collection

    mStep = "checking the setup"
    ShowStatus "checking the setup..."
    mUpgradeNote = ""
    EnsureConfigSheet
    EnsureConfigPanel
    CheckSetup opt, problems, warnings, True, cancelled
    If cancelled Then
        Application.StatusBar = False
        Exit Sub
    End If
    If problems.Count > 0 Or Not mReady Then
        Application.StatusBar = False
        WriteLastRun "Not built: " & CountText(problems.Count, "problem") & " to fix", CLR_BAD, CityNames(), _
                     ListText(problems, "PROBLEM: ", MAX_LISTED, 0) & ListText(warnings, "WARNING: ", MAX_LISTED, 0)
        MsgBox ClipMessage("The report was not built. Fix this on the Config sheet, then click Build report again:" & _
                           vbLf & vbLf & ListText(problems, "- ", MSG_LISTED, 150)), vbExclamation, "Open Orders report"
        Exit Sub
    End If

    mStep = "sorting the rows"
    ShowStatus "sorting..."
    SortData
    BuildLogLines opt

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    mStep = "creating the output workbook"
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Application.Calculation = xlCalculationManual
    Set wsFmt = wbOut.Worksheets(1)
    wsFmt.Name = "File Format"
    Set wsOI = wbOut.Worksheets.Add(After:=wsFmt)
    wsOI.Name = "Order Intake"
    Set wsInv = wbOut.Worksheets.Add(After:=wsOI)
    wsInv.Name = "Inventory by Month"
    Set wsOO = wbOut.Worksheets.Add(After:=wsInv)
    wsOO.Name = "Open Orders by MSA"
    Set wsLog = wbOut.Worksheets.Add(After:=wsOO)
    wsLog.Name = "Run Log"
    builtAt = Now

    mStep = "writing the File Format tab"
    WriteFileFormat wsFmt, opt, builtAt
    mStep = "writing the Order Intake tab"
    ShowStatus "writing Order Intake (" & mNOI & " rows)..."
    WriteOrderIntake wsOI, opt, builtAt
    mStep = "writing the Inventory by Month tab"
    ShowStatus "writing Inventory by Month (" & mNINV & " rows)..."
    WriteInventory wsInv, opt, builtAt
    mStep = "writing the Open Orders by MSA tab"
    ShowStatus "writing Open Orders by MSA..."
    WriteOpenOrders wsOO, opt, builtAt
    mStep = "setting up the pages for printing"
    SetPrintCommunication False
    SetUpPage wsFmt, "", 1, builtAt
    SetUpPage wsOI, "$1:$2", 0, builtAt
    SetUpPage wsInv, "$" & (7 + mNCities) & ":$" & (7 + mNCities), 0, builtAt
    SetUpPage wsOO, "", 1, builtAt
    SetUpPage wsLog, "", 0, builtAt
    SetPrintCommunication True
    mStep = "writing the Run Log"
    WriteRunLog wsLog, opt, builtAt, Timer - t0
    HideGridlines wsFmt
    HideGridlines wsLog

    mStep = "saving the output"
    ShowStatus "saving..."
    wbOut.Activate
    wsFmt.Activate
    wsFmt.Range("A1").Select
    outPath = opt.OutputFolder & Application.PathSeparator & opt.OutputFileName
    outPath = ConfirmOutputPath(outPath, wbOut)
    Application.Calculation = xlCalculationAutomatic     ' evaluate the summary and store the file in automatic mode
    Application.DisplayAlerts = False
    wbOut.SaveAs Filename:=outPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = prevAlerts

    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Application.StatusBar = False

    If Not opt.OpenAfterBuild Then wbOut.Close SaveChanges:=False

    Set toCheck = New Collection                        ' Run Log lines a person should look at
    For i = 1 To mLog.Count
        If RunLogLineToCheck(CStr(mLog.Item(i))) And Not InCollection(warnings, CStr(mLog.Item(i))) Then toCheck.Add mLog.Item(i)
    Next i
    nItems = warnings.Count + toCheck.Count
    details = ListText(warnings, "WARNING: ", MAX_LISTED, 0) & ListText(toCheck, "RUN LOG: ", MAX_LISTED, 0)
    If nItems = 0 Then
        WriteLastRun "Report built", CLR_OK, CityNames(), "", outPath
    Else
        WriteLastRun "Report built - " & CountText(nItems, "item") & " to look at", CLR_WARN, CityNames(), details, outPath
    End If
    result = "Report built in " & Format$(Timer - t0, "0.0") & " s" & vbCrLf & vbCrLf & outPath & vbCrLf & vbCrLf & _
             "MSAs: " & CityNames() & vbCrLf & "Inventory month: " & IIf(mMonth >= 0, MonthText(mMonth), "(none)") & _
             vbCrLf & vbCrLf & "See the Run Log tab for row counts and any UNMAPPED model groups."
    If nItems > 0 Then result = result & vbCrLf & vbCrLf & CountText(nItems, "item") & " to look at: see Last run on the Config sheet."
    MsgBox ClipMessage(result), vbInformation, "Open Orders report"
    Exit Sub

Fail:
    Dim msg As String
    msg = "Failed while " & mStep & vbCrLf & vbCrLf & "Error " & Err.Number & ": " & Err.Description
    If Not wbOut Is Nothing Then msg = msg & vbCrLf & vbCrLf & "The partly built report is left open and unsaved so you can see how far it got."
    Application.DisplayAlerts = prevAlerts
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Application.StatusBar = False
    WriteLastRun "Failed while " & mStep, CLR_BAD, "", Replace(msg, vbCrLf, vbLf)
    MsgBox msg, vbCritical, "Open Orders report"
End Sub

' Button "Check setup": the checks every build starts with, without building anything. Problems (which
' would stop a build) and warnings go to a message and to the Last run box.
Public Sub CheckOpenOrdersReportSetup(Optional ByVal notInMacroList As Variant)
    Dim opt As ReportOptions, cancelled As Boolean
    Dim problems As Collection, warnings As Collection, msg As String, names As String, reportText As String
    On Error GoTo Fail
    Set mLog = New Collection
    Set problems = New Collection
    Set warnings = New Collection
    mUpgradeNote = ""
    EnsureConfigSheet
    EnsureConfigPanel
    CheckSetup opt, problems, warnings, False, cancelled
    Application.StatusBar = False
    names = CityNames()
    If Len(opt.OutputFolder) > 0 Then reportText = opt.OutputFolder & Application.PathSeparator & opt.OutputFileName Else reportText = "(folder not set yet)"
    If problems.Count = 0 Then
        If warnings.Count = 0 Then
            WriteLastRun "Setup check: ready to build", CLR_OK, names, ""
        Else
            WriteLastRun "Setup check: ready to build (" & CountText(warnings.Count, "warning") & ")", CLR_WARN, names, _
                         ListText(warnings, "WARNING: ", MAX_LISTED, 0)
        End If
        msg = "Ready to build."
    Else
        WriteLastRun "Setup check: " & CountText(problems.Count, "problem") & " to fix", CLR_BAD, names, _
                     ListText(problems, "PROBLEM: ", MAX_LISTED, 0) & ListText(warnings, "WARNING: ", MAX_LISTED, 0)
        msg = "Not ready. Fix this on the Config sheet, then click Check setup again:" & vbLf & vbLf & ListText(problems, "- ", MSG_LISTED, 150)
    End If
    msg = msg & vbLf & vbLf & "MSAs: " & names & vbLf & "Inventory months: " & MonthsText() & vbLf & "Report: " & reportText
    If warnings.Count > 0 Then msg = msg & vbLf & vbLf & "Warnings (a build still runs):" & vbLf & ListText(warnings, "- ", MSG_LISTED, 150)
    If problems.Count = 0 Then
        MsgBox ClipMessage(msg), vbInformation, "Open Orders report"
    Else
        MsgBox ClipMessage(msg), vbExclamation, "Open Orders report"
    End If
    Exit Sub
Fail:
    Application.StatusBar = False
    MsgBox "Setup check failed. Error " & Err.Number & ": " & Err.Description, vbCritical, "Open Orders report"
End Sub

' Button "Choose input folder...": a folder picker that writes the chosen folder into Config InputFolder.
Public Sub ChooseOpenOrdersInputFolder(Optional ByVal notInMacroList As Variant)
    ChooseFolderFor "InputFolder", "Select the folder holding the Order Intake / Inventory by Month / Open Order exports"
End Sub

' Button "Choose output folder...": a folder picker that writes the chosen folder into Config OutputFolder.
Public Sub ChooseOpenOrdersOutputFolder(Optional ByVal notInMacroList As Variant)
    ChooseFolderFor "OutputFolder", "Select the folder to save the report in"
End Sub

' Button "Choose size class file...": a file picker that writes the chosen workbook into Config SizeClassFile.
Public Sub ChooseOpenOrdersSizeClassFile(Optional ByVal notInMacroList As Variant)
    Dim c As Range, p As String
    EnsureConfigReady
    Set c = ConfigCell("SizeClassFile")
    If c Is Nothing Then
        MsgBox "The Config sheet has no SizeClassFile row. Run ShowOpenOrdersReportConfig (Alt+F8).", vbExclamation, "Open Orders report"
        Exit Sub
    End If
    p = PickFile("Select the size class workbook (the master list of models and size classes)", GetConfig("SizeClassFile"))
    If Len(p) > 0 Then SetConfigValue c, p
End Sub

' =====================================================================================================
'  SETUP CHECK  (run before every build, and by the Check setup button)
' =====================================================================================================

' Reads Config into opt and checks everything a build needs before anything is created: the settings, the
' folders (the output folder must take a test file), the size class file, the exports (read in full, so the
' cities and months are known) and the report file. Problems stop a build; warnings do not. With forBuild,
' a blank InputFolder opens the folder picker (cancelled = True when it is closed); without, it is a
' warning. Sets mReady when the rows were prepared.
Private Sub CheckSetup(opt As ReportOptions, problems As Collection, warnings As Collection, _
                       ByVal forBuild As Boolean, ByRef cancelled As Boolean)
    Dim keys As Variant, i As Long, k As Long, inputOk As Boolean, cityOk As Boolean, errText As String
    Dim kinds As Variant, labels As Variant, n As Long, names As String, wbx As Workbook

    mReady = False
    mNFiles = 0
    mNCities = 0
    mNMonths = 0
    mMonth = -1
    mNonIsoText = False
    Set mScWarnings = New Collection
    mScCount = 0
    mClsCount = 0                                        ' v2.8: no classes until the size class file is read
    mClsLine = ""
    If Len(mUpgradeNote) > 0 Then warnings.Add mUpgradeNote
    keys = ConfigKeys()
    For i = LBound(keys) To UBound(keys)
        If ConfigRow(CStr(keys(i))) = 0 Then problems.Add "The Config sheet has no " & keys(i) & " row. Run " & _
            "ShowOpenOrdersReportConfig (Alt+F8) to repair it."
    Next i
    opt = ReadOptions(problems)

    If Len(opt.InputFolder) = 0 Then
        If Not forBuild Then
            warnings.Add "InputFolder is blank: Build report will ask for the folder, so the export files were not checked. " & _
                         "Click Choose input folder to set it."
        ElseIf problems.Count = 0 Then
            opt.InputFolder = PickFolder("Select the folder holding the Order Intake / Inventory by Month / Open Order exports", "")
            If Len(opt.InputFolder) = 0 Then
                cancelled = True
                Exit Sub
            End If
        End If
    End If
    If Len(opt.InputFolder) > 0 Then
        opt.InputFolder = StripTrailingSeparator(opt.InputFolder)
        If FolderExists(opt.InputFolder) Then
            inputOk = True
        Else
            problems.Add "InputFolder not found: " & opt.InputFolder & ". Check the path, or click Choose input folder."
        End If
    End If

    If Len(opt.OutputFolder) = 0 Then
        opt.OutputFolder = opt.InputFolder
    Else
        opt.OutputFolder = StripTrailingSeparator(opt.OutputFolder)
        If Not FolderExists(opt.OutputFolder) Then problems.Add "OutputFolder not found: " & opt.OutputFolder & _
            ". Check the path, click Choose output folder, or leave it blank to save next to the exports."
    End If
    If Len(opt.OutputFolder) > 0 Then
        If FolderExists(opt.OutputFolder) Then
            If Not CanWriteFolder(opt.OutputFolder) Then problems.Add "The report cannot be saved in " & opt.OutputFolder & _
                " (read-only, or no permission). Choose another OutputFolder."
        End If
    End If
    For Each wbx In Application.Workbooks               ' Excel cannot save over a workbook open with unsaved changes
        If StrComp(wbx.Name, opt.OutputFileName, vbTextCompare) = 0 And Not wbx.Saved Then problems.Add opt.OutputFileName & _
            " is open in Excel with unsaved changes. Save or close it first."
    Next wbx

    ' the size class file (the master list)
    If Len(opt.SizeClassFile) = 0 Then
        problems.Add "SizeClassFile is blank. Click Choose size class file... and pick the size class workbook " & _
                     "(the master list of models and size classes)."
    ElseIf LCase$(Left$(opt.SizeClassFile, 4)) = "http" Then
        problems.Add "SizeClassFile is a web address (" & opt.SizeClassFile & "). Use the path of the file on this PC " & _
                     "(a synced SharePoint or OneDrive folder works): click Choose size class file."
    ElseIf Not FileExists(opt.SizeClassFile) Then
        problems.Add "SizeClassFile not found: " & opt.SizeClassFile & ". Click Choose size class file."
    Else
        ShowStatus "reading the size class file..."
        LoadSizeClassFile opt.SizeClassFile, problems
        For i = 1 To mScWarnings.Count
            warnings.Add mScWarnings.Item(i)
        Next i
    End If

    ' the exports: found, read in full, city column resolved, rows prepared
    If inputOk Then
        DiscoverFiles opt.InputFolder
        kinds = Array("OI", "INV", "OO")
        labels = Array("Order Intake", "Inventory by Month", "Open Order")
        For k = 0 To 2
            n = 0
            names = ""
            For i = 1 To mNFiles
                If mFiles(i).Kind = kinds(k) Then
                    n = n + 1
                    names = names & IIf(n > 1, ", ", "") & mFiles(i).FileName
                End If
            Next i
            If k = 0 And n = 0 Then problems.Add "No Order Intake files in " & opt.InputFolder & ". Name them like " & _
                """Order Intake 2026.xlsx"" (optional: ""Inventory by Month 2026.xlsx"", ""Open Orders 2026.xlsx""). " & _
                "Save CSV exports as .xlsx first."
            If n > 1 Then warnings.Add n & " " & labels(k) & " files found, all are read: " & names
        Next k
        cityOk = True
        For i = 1 To mNFiles
            mStep = "reading " & mFiles(i).FileName
            ShowStatus "reading " & mFiles(i).FileName & " (" & i & " of " & mNFiles & ")..."
            If Not TryReadSource(mFiles(i).Path, mFiles(i).Tbl, errText) Then
                problems.Add mFiles(i).FileName & " could not be read: " & errText
                cityOk = False
            ElseIf Not mFiles(i).Tbl.Ok Then
                warnings.Add FileSummary(i)
            Else
                ResolveColumns i, opt, problems, cityOk
                If Len(mFiles(i).Missing) > 0 Then warnings.Add FileSummary(i)
            End If
        Next i
        If cityOk And mNFiles > 0 Then
            mStep = "preparing the rows"
            ShowStatus "preparing the rows..."
            PrepareData opt
            mReady = True
            For i = 1 To mNFiles
                If mFiles(i).BlankCity > 0 And mFiles(i).Kind <> "OO" Then warnings.Add mFiles(i).FileName & ": " & _
                    CountText(mFiles(i).BlankCity, "row") & " with a blank MSA - shown as " & NO_CITY
            Next i
            If Not forBuild And mScCount > 0 And Len(mUnmappedAll) > 0 Then warnings.Add _
                "Models in Order Intake that are not in the size class file (UNMAPPED until added): " & Replace(mUnmappedAll, "|", ", ")
        End If
    End If

    If SheetExists(ThisWorkbook, OLD_MAP_SHEET) Then warnings.Add OLD_MAP_SHEET & " is no longer used; the master list is " & _
        "SizeClassFile. You can delete the sheet."
    If Len(ThisWorkbook.Path) = 0 Then
        warnings.Add "This workbook has never been saved. Save it as an Excel Macro-Enabled Workbook (.xlsm), " & _
                     "or the macro is lost when Excel closes."
    ElseIf LCase$(Right$(ThisWorkbook.Name, 5)) = ".xlsx" Then
        warnings.Add "This workbook is an .xlsx, which cannot keep macros. Save it as an Excel Macro-Enabled Workbook (.xlsm)."
    End If
End Sub

' True when a file can be created and deleted in folder p.
Private Function CanWriteFolder(ByVal p As String) As Boolean
    Dim f As Integer, testPath As String
    testPath = p & Application.PathSeparator & "~open-orders-write-test.tmp"
    On Error GoTo Done
    f = FreeFile
    Open testPath For Output As #f
    Close #f
    Kill testPath
    CanWriteFolder = True
Done:
End Function

' Run Log lines worth a person's look after a build (they also stay on the Run Log sheet).
Private Function RunLogLineToCheck(ByVal s As String) As Boolean
    RunLogLineToCheck = InStr(s, "UNMAPPED") > 0 Or InStr(s, "WARNING") > 0 Or InStr(s, "missing column") > 0 Or _
                        InStr(s, "left as is (") > 0 Or InStr(s, "is empty") > 0 Or InStr(s, "no valid") > 0 Or _
                        InStr(s, "left blank") > 0 Or InStr(s, "dropped") > 0 Or _
                        InStr(s, NO_CITY) > 0 Or InStr(s, "all are read") > 0 Or InStr(s, "shows zeros") > 0
End Function

Private Function InCollection(items As Collection, ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To items.Count
        If CStr(items.Item(i)) = s Then
            InCollection = True
            Exit Function
        End If
    Next i
End Function

' =====================================================================================================
'  CONFIG SHEET  (the control panel)
' =====================================================================================================

Private Function ConfigKeys() As Variant
    ConfigKeys = Array("InputFolder", "OutputFolder", "OutputFileName", "SizeClassFile", "MSAColumn", "InventoryMonth", _
                       "ShowCancelled", "ShowSSL", "OpenAfterBuild")
End Function

Private Function ConfigHelp() As Variant
    ConfigHelp = Array( _
        "Folder holding the Snowflake exports (Order Intake, Inventory by Month, Open Order) for all MSAs. Blank = a folder picker opens when you build.", _
        "Where the report is saved. Blank = same as InputFolder.", _
        "File name of the report workbook (.xlsx is added if missing). If it already exists you are asked before it is overwritten.", _
        "Full path of the size class workbook, the master list of models and size classes (a synced SharePoint or OneDrive folder works). Required. Click Choose size class file.", _
        "Header of the MSA column in the exports (MSA in today's exports). Blank = found automatically (a header MSA, then one containing MSA, then CITY, METRO or MARKET); without one the MSA comes from the file name.", _
        "Any day in the inventory month to show, e.g. 8/1/2026. Blank = the latest month in the Inventory by Month file.", _
        "FALSE hides cancelled orders behind the filter on ORDER_STATUS; they stay in the report. TRUE shows them.", _
        "FALSE hides skid-steer loaders behind the filters on both tabs and off the Open Orders by MSA chart; they stay in the report. TRUE shows them.", _
        "TRUE leaves the finished report open in Excel; FALSE closes it after saving.")
End Function

' Creates the Config sheet, or converts a version 1.x sheet in place (mUpgradeNote says so).
Private Sub EnsureConfigSheet()
    Dim ws As Worksheet, vals(0 To 8) As Variant
    If SheetExists(ThisWorkbook, CFG_SHEET) Then
        If ConfigRow("MSAColumn") = 0 Then UpgradeConfigSheet
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = CFG_SHEET
    ws.Range("A1:C1").Value = Array("Setting", "Value", "What it does")
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C1").Interior.Color = CLR_YELLOW
    vals(2) = DEFAULT_OUTPUT_NAME
    vals(6) = "FALSE"
    vals(7) = "FALSE"
    vals(8) = "TRUE"
    WriteConfigRows ws, vals
End Sub

' For the Choose buttons: the Config sheet exists and is version 2.0; one just upgraded from 1.x gets the
' panel laid out and the upgrade note in the Last run box (Build report and Check setup do the same in their way).
Private Sub EnsureConfigReady()
    mUpgradeNote = ""
    EnsureConfigSheet
    If Len(mUpgradeNote) > 0 Then
        SetUpConfigPanel
        WriteLastRun mUpgradeNote, CLR_WARN, "-", mUpgradeNote
    End If
End Sub

' Rows 2 to 10: the nine settings with their values and help text.
Private Sub WriteConfigRows(ws As Worksheet, vals() As Variant)
    Dim keys As Variant, help As Variant, i As Long
    keys = ConfigKeys()
    help = ConfigHelp()
    For i = 0 To 8
        ws.Cells(2 + i, 1).Value = keys(i)
        ws.Cells(2 + i, 2).Value = vals(i)
        ws.Cells(2 + i, 3).Value = help(i)
    Next i
    ws.Cells(7, 2).NumberFormat = "m/d/yyyy"
    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 60
    ws.Columns("C").ColumnWidth = 110
    ws.Range("B2:B10").Interior.Color = CLR_INPUT
End Sub

' An older Config sheet (no MSAColumn row): the values are read by name, rows 2 to 10 are rewritten with the
' version 2.1 settings and the values carried over. From 1.x: ShowCancelled = NOT DropCancelled, ShowSSL =
' IncludeSSL, InventoryMonth = InventorySnapshotDate; HeaderCountMode and InventoryHeaderRow are dropped.
' From 2.0: CityColumn becomes MSAColumn; everything else is kept.
Private Sub UpgradeConfigSheet()
    Dim ws As Worksheet, vals(0 To 8) As Variant, fromV1 As Boolean
    Set ws = ThisWorkbook.Worksheets(CFG_SHEET)
    fromV1 = (ConfigRow("DropCancelled") > 0 Or ConfigRow("IncludeSSL") > 0 Or ConfigRow("HeaderCountMode") > 0 Or _
              ConfigRow("InventorySnapshotDate") > 0)
    vals(0) = GetConfigValue("InputFolder")
    vals(1) = GetConfigValue("OutputFolder")
    vals(2) = GetConfigValue("OutputFileName")
    If Len(Trim$(SafeStr(vals(2)))) = 0 Then vals(2) = DEFAULT_OUTPUT_NAME
    vals(3) = GetConfigValue("SizeClassFile")
    vals(4) = GetConfigValue("CityColumn")
    If ConfigRow("InventoryMonth") > 0 Then vals(5) = GetConfigValue("InventoryMonth") Else vals(5) = GetConfigValue("InventorySnapshotDate")
    If ConfigRow("ShowCancelled") > 0 Then
        vals(6) = IIf(BoolText(GetConfigValue("ShowCancelled"), False), "TRUE", "FALSE")
    Else
        vals(6) = IIf(BoolText(GetConfigValue("DropCancelled"), True), "FALSE", "TRUE")
    End If
    If ConfigRow("ShowSSL") > 0 Then
        vals(7) = IIf(BoolText(GetConfigValue("ShowSSL"), False), "TRUE", "FALSE")
    Else
        vals(7) = IIf(BoolText(GetConfigValue("IncludeSSL"), False), "TRUE", "FALSE")
    End If
    vals(8) = IIf(BoolText(GetConfigValue("OpenAfterBuild"), True), "TRUE", "FALSE")
    ws.Unprotect
    ws.Range("A2:C10").Validation.Delete
    ws.Range("A2:C10").ClearContents
    WriteConfigRows ws, vals
    If fromV1 Then
        mUpgradeNote = "Upgraded to v2.8: HeaderCountMode and InventoryHeaderRow are no longer used; set SizeClassFile."
    Else
        mUpgradeNote = "Upgraded to v2.8: CityColumn is now MSAColumn (the MSA column of the exports); the values are kept."
    End If
End Sub

' A TRUE/FALSE value as the setting reads it; blank or unreadable = dflt.
Private Function BoolText(ByVal v As Variant, ByVal dflt As Boolean) As Boolean
    BoolText = dflt
    If VarType(v) = vbBoolean Then
        BoolText = v
        Exit Function
    End If
    Select Case UCase$(Trim$(SafeStr(v)))
        Case "TRUE", "YES", "Y", "1": BoolText = True
        Case "FALSE", "NO", "N", "0": BoolText = False
    End Select
End Function

' Lays out the Config control panel and protects Config.
Private Sub SetUpConfigPanel()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CFG_SHEET)
    ws.Unprotect
    LayoutConfigSheet ws
    EnsureButtons ws
    ProtectConfig ws
End Sub

' A Config sheet without the panel (new, made by version 1.1, or just upgraded) gets it before a build or
' check uses it, so the Last run box never lands on a sheet whose Value cells are still locked.
Private Sub EnsureConfigPanel()
    If SafeStr(ThisWorkbook.Worksheets(CFG_SHEET).Cells(STATUS_ROW, 1).Value) <> "Last run" Or Len(mUpgradeNote) > 0 Then SetUpConfigPanel
End Sub

' The control panel around the settings table: Value cells unlocked, shaded and checked as they are typed
' (a pasted value skips these checks, which is why every build checks the settings again), help text
' wrapped, the button row, the Last run box and the help box. Safe to run again: values are not touched.
Private Sub LayoutConfigSheet(ws As Worksheet)
    Dim keys As Variant, i As Long, r As Long, helpRows As Variant, parts As Variant
    ws.Cells.Locked = True
    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 60
    ws.Columns("C").ColumnWidth = 110
    keys = ConfigKeys()
    For i = LBound(keys) To UBound(keys)
        r = ConfigRow(CStr(keys(i)))
        If r > 0 Then
            ws.Cells(r, 2).Locked = False
            ws.Cells(r, 2).Interior.Color = CLR_INPUT
            ws.Cells(r, 3).WrapText = True
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).VerticalAlignment = xlTop
            ws.Rows(r).AutoFit
        End If
    Next i
    If Not ConfigCell("InventoryMonth") Is Nothing Then ConfigCell("InventoryMonth").NumberFormat = "m/d/yyyy"
    AddCheck ConfigCell("InputFolder"), xlValidateInputOnly, "", "", "InputFolder", _
        "Click Choose input folder, or paste the folder path. Blank = asked when you build.", ""
    AddCheck ConfigCell("OutputFolder"), xlValidateInputOnly, "", "", "OutputFolder", _
        "Click Choose output folder, or paste the folder path. Blank = save next to the exports.", ""
    AddCheck ConfigCell("OutputFileName"), xlValidateInputOnly, "", "", "OutputFileName", _
        "Name of the report file. .xlsx is added if missing. Not allowed: \ / : * ? "" < > |", ""
    AddCheck ConfigCell("SizeClassFile"), xlValidateInputOnly, "", "", "SizeClassFile", _
        "Click Choose size class file, or paste the full path of the size class workbook. Required.", ""
    AddCheck ConfigCell("MSAColumn"), xlValidateInputOnly, "", "", "MSAColumn", _
        "Header of the MSA column in the exports (MSA today). Blank = found automatically.", ""
    AddCheck ConfigCell("InventoryMonth"), xlValidateDate, "36526", "73415", "InventoryMonth", _
        "Any day in the month to show, such as 8/1/2026, or blank for the latest month in the file.", _
        "Type a date such as 8/1/2026, or leave the cell blank."
    AddCheck ConfigCell("ShowCancelled"), xlValidateList, "TRUE,FALSE", "", "ShowCancelled", _
        "FALSE hides cancelled orders behind the filter; TRUE shows them.", "Pick TRUE or FALSE from the list."
    AddCheck ConfigCell("ShowSSL"), xlValidateList, "TRUE,FALSE", "", "ShowSSL", _
        "FALSE hides skid-steer loaders behind the filters; TRUE shows them.", "Pick TRUE or FALSE from the list."
    AddCheck ConfigCell("OpenAfterBuild"), xlValidateList, "TRUE,FALSE", "", "OpenAfterBuild", _
        "TRUE leaves the report open after it is saved.", "Pick TRUE or FALSE from the list."

    ws.Rows(BUTTON_ROW).RowHeight = 34

    BoxTitle ws, STATUS_ROW, "Last run", "Written by Build report and Check setup. Problems stop a build; warnings and RUN LOG items do not."
    parts = Array("Result", "When", "Last report", "MSAs", "Details")
    For i = 0 To 4
        BoxRow ws, STATUS_ROW + 1 + i, CStr(parts(i))
    Next i
    ws.Cells(STATUS_ROW + 2, 2).NumberFormat = "yyyy-mm-dd hh:mm"
    ws.Cells(STATUS_ROW + 2, 2).HorizontalAlignment = xlLeft

    BoxTitle ws, HELP_ROW, "If something goes wrong", ""
    helpRows = Array( _
        "Buttons do nothing|Excel has blocked the macros. Close this workbook, right-click the .xlsm file > Properties > tick Unblock > OK, " & _
            "open it again and click Enable Content. Or keep the file in a Trusted Location (File > Options > Trust Center).", _
        "Build stops with a list|Each line names the setting or file to fix. Fix them, click Check setup until it says ready, then Build report.", _
        "Size Class is UNMAPPED|Add the model (MODEL_GROUP) and its size class to the size class file (SizeClassFile; a class from its Set-up tab, as " & _
            "listed on the File Format tab of the report), save it and build again.", _
        "An MSA is missing or wrong|The MSA comes from the MSA column of the exports, found automatically (a header MSA first, then " & _
            "CITY, METRO, MARKET) unless MSAColumn names it; the Run Log says which column each file used. Type the right header into " & _
            "MSAColumn. Rows with a blank MSA show as (no MSA).", _
        "Changing this sheet|Protected without a password so only the yellow cells change: Review > Unprotect Sheet to edit, then run " & _
            "ShowOpenOrdersReportConfig (Alt+F8) to tidy and protect it again.")
    For i = 0 To UBound(helpRows)
        parts = Split(helpRows(i), "|")
        BoxRow ws, HELP_ROW + 1 + i, CStr(parts(0))
        ws.Cells(HELP_ROW + 1 + i, 2).Value = parts(1)
        ws.Rows(HELP_ROW + 1 + i).RowHeight = TextHeight(CStr(parts(1)))
    Next i
End Sub

' A box title row: yellow across A:C, title in A, a note in B.
Private Sub BoxTitle(ws As Worksheet, ByVal r As Long, ByVal title As String, ByVal note As String)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).UnMerge
    ws.Cells(r, 1).Value = title
    ws.Cells(r, 2).Value = note
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Interior.Color = CLR_YELLOW
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Bold = True
    ws.Cells(r, 2).Font.Bold = False
End Sub

' A box row: grey label in A, B:C merged and wrapped for the text.
Private Sub BoxRow(ws As Worksheet, ByVal r As Long, ByVal label As String)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 1).Interior.Color = CLR_BOX
    ws.Cells(r, 1).Font.Bold = True
    If Not ws.Cells(r, 2).MergeCells Then ws.Cells(r, 3).ClearContents     ' so merging asks nothing
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Merge
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).WrapText = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).VerticalAlignment = xlTop
End Sub

' Data validation on one Value cell: a drop-down or range check with a hint shown when the cell is selected.
' Formulas are serial numbers and plain lists so they read the same in every regional setting.
Private Sub AddCheck(c As Range, ByVal checkType As Long, ByVal f1 As String, ByVal f2 As String, _
                     ByVal title As String, ByVal hint As String, ByVal errorText As String)
    If c Is Nothing Then Exit Sub
    With c.Validation
        .Delete
        If checkType = xlValidateInputOnly Then
            .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
        ElseIf Len(f2) > 0 Then
            .Add Type:=checkType, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=f1, Formula2:=f2
        Else
            .Add Type:=checkType, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=f1
        End If
        .IgnoreBlank = True
        If checkType = xlValidateList Then .InCellDropdown = True
        .InputTitle = Left$(title, 32)
        .InputMessage = Left$(hint, 255)
        .ShowInput = True
        If Len(errorText) > 0 Then
            .ErrorTitle = Left$(title, 32)
            .ErrorMessage = Left$(errorText, 255)
            .ShowError = True
        End If
    End With
End Sub

' The five buttons (Forms controls) on the button row, created once and put back in place on every run.
Private Sub EnsureButtons(ws As Worksheet)
    Dim x As Double, y As Double
    x = ws.Cells(BUTTON_ROW, 1).Left + 2
    y = ws.Cells(BUTTON_ROW, 1).Top + 4
    PlaceButton ws, BUTTON_NAME, "Build report", "BuildOpenOrdersReport", x, y, 130
    PlaceButton ws, BTN_CHECK, "Check setup", "CheckOpenOrdersReportSetup", x + 140, y, 110
    PlaceButton ws, BTN_INPUT, "Choose input folder...", "ChooseOpenOrdersInputFolder", x + 260, y, 150
    PlaceButton ws, BTN_OUTPUT, "Choose output folder...", "ChooseOpenOrdersOutputFolder", x + 420, y, 160
    PlaceButton ws, BTN_SIZE, "Choose size class file...", "ChooseOpenOrdersSizeClassFile", x + 590, y, 170
End Sub

Private Sub PlaceButton(ws As Worksheet, ByVal shapeName As String, ByVal caption As String, ByVal macroName As String, _
                        ByVal x As Double, ByVal y As Double, ByVal w As Double)
    Dim shp As Shape, s As Shape
    For Each s In ws.Shapes
        If s.Name = shapeName Then Set shp = s
    Next s
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, x, y, w, 26)
        shp.Name = shapeName
    End If
    shp.Left = x
    shp.Top = y
    shp.Width = w
    shp.Height = 26
    shp.Placement = xlFreeFloating
    shp.OnAction = macroName
    shp.TextFrame.Characters.Text = caption
End Sub

' No password: protection only stops accidental edits (Review > Unprotect Sheet lifts it).
Private Sub ProtectConfig(ws As Worksheet)
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
End Sub

' Fills the Last run box. A convenience only: any error here is ignored, so it never stops or fails a run.
' reportPath (shown as a link) is written only after a successful build; otherwise the last link stays.
' The sheet is protected again only if it was protected (someone may have unprotected it to edit it).
Private Sub WriteLastRun(ByVal result As String, ByVal resultColor As Long, ByVal citiesText As String, _
                         ByVal details As String, Optional ByVal reportPath As String = "")
    Dim ws As Worksheet, wasProtected As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CFG_SHEET)
    If ws Is Nothing Then Exit Sub
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect
    Do While Right$(details, 1) = vbLf
        details = Left$(details, Len(details) - 1)
    Loop
    If Len(details) = 0 Then details = "(nothing to look at)"
    If Len(citiesText) = 0 Then citiesText = "-"
    ws.Cells(STATUS_ROW + 1, 2).Value = SafeCellText(result)
    ws.Range(ws.Cells(STATUS_ROW + 1, 2), ws.Cells(STATUS_ROW + 1, 3)).Interior.Color = resultColor
    ws.Cells(STATUS_ROW + 2, 2).Value = Now
    If Len(reportPath) > 0 Then
        ws.Cells(STATUS_ROW + 3, 2).Hyperlinks.Delete
        ws.Cells(STATUS_ROW + 3, 2).Value = SafeCellText(reportPath)
        ws.Hyperlinks.Add Anchor:=ws.Cells(STATUS_ROW + 3, 2), Address:=reportPath, TextToDisplay:=reportPath
    End If
    ws.Cells(STATUS_ROW + 4, 2).Value = SafeCellText(citiesText)
    ws.Rows(STATUS_ROW + 4).RowHeight = TextHeight(citiesText)
    ws.Cells(STATUS_ROW + 5, 2).Value = SafeCellText(Left$(details, 30000))
    ws.Rows(STATUS_ROW + 5).RowHeight = TextHeight(details)
    If wasProtected Then ProtectConfig ws
End Sub

' Row height for wrapped text in merged B:C (about 160 characters and 15 points a line, at most 409 points).
Private Function TextHeight(ByVal s As String) As Double
    Dim parts As Variant, i As Long, nLines As Long, h As Double
    parts = Split(s, vbLf)
    For i = LBound(parts) To UBound(parts)
        nLines = nLines + 1 + Len(parts(i)) \ 160
    Next i
    If nLines < 1 Then nLines = 1
    h = nLines * 15
    If h > 409 Then h = 409
    TextHeight = h
End Function

Private Sub ChooseFolderFor(ByVal settingKey As String, ByVal title As String)
    Dim c As Range, p As String
    EnsureConfigReady
    Set c = ConfigCell(settingKey)
    If c Is Nothing Then
        MsgBox "The Config sheet has no " & settingKey & " row. Run ShowOpenOrdersReportConfig.", vbExclamation, "Open Orders report"
        Exit Sub
    End If
    p = PickFolder(title, GetConfig(settingKey))
    If Len(p) > 0 Then SetConfigValue c, StripTrailingSeparator(p)
End Sub

' Writes a picked path into a Value cell, also on a Config sheet laid out by an older version (locked cells).
Private Sub SetConfigValue(c As Range, ByVal p As String)
    If c.Locked And c.Worksheet.ProtectContents Then
        c.Worksheet.Unprotect
        c.Value = p
        ProtectConfig c.Worksheet
    Else
        c.Value = p
    End If
End Sub

' Row of settingKey in column A of Config (0 if missing).
Private Function ConfigRow(ByVal settingKey As String) As Long
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(CFG_SHEET)
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If StrComp(Trim$(SafeStr(ws.Cells(r, 1).Value)), settingKey, vbTextCompare) = 0 Then
            ConfigRow = r
            Exit Function
        End If
    Next r
End Function

' The Value cell of settingKey (Nothing if the row is missing).
Private Function ConfigCell(ByVal settingKey As String) As Range
    Dim r As Long
    r = ConfigRow(settingKey)
    If r > 0 Then Set ConfigCell = ThisWorkbook.Worksheets(CFG_SHEET).Cells(r, 2)
End Function

' The Config value for settingKey as stored in the cell (Empty if the key is missing).
Private Function GetConfigValue(ByVal settingKey As String) As Variant
    Dim r As Long
    r = ConfigRow(settingKey)
    If r > 0 Then GetConfigValue = ThisWorkbook.Worksheets(CFG_SHEET).Cells(r, 2).Value
End Function

Private Function GetConfig(ByVal settingKey As String) As String
    GetConfig = Trim$(SafeStr(GetConfigValue(settingKey)))
End Function

' TRUE / FALSE (also YES / NO, Y / N, 1 / 0); blank = dflt; anything else is a problem.
Private Function ReadBool(ByVal settingKey As String, ByVal dflt As Boolean, problems As Collection) As Boolean
    Dim v As Variant
    v = GetConfigValue(settingKey)
    If VarType(v) = vbBoolean Then
        ReadBool = v
        Exit Function
    End If
    Select Case UCase$(Trim$(SafeStr(v)))
        Case "": ReadBool = dflt
        Case "TRUE", "YES", "Y", "1": ReadBool = True
        Case "FALSE", "NO", "N", "0": ReadBool = False
        Case Else
            ReadBool = dflt
            problems.Add settingKey & " must be TRUE or FALSE, not """ & SafeStr(v) & """."
    End Select
End Function

' The settings, with every value that cannot be used added to problems (no setting is silently replaced).
Private Function ReadOptions(problems As Collection) As ReportOptions
    Dim o As ReportOptions, v As Variant, d As Date
    o.InputFolder = GetConfig("InputFolder")
    o.OutputFolder = GetConfig("OutputFolder")
    o.OutputFileName = GetConfig("OutputFileName")
    If Len(o.OutputFileName) = 0 Then o.OutputFileName = DEFAULT_OUTPUT_NAME
    If LCase$(Right$(o.OutputFileName, 5)) <> ".xlsx" Then o.OutputFileName = o.OutputFileName & ".xlsx"
    If o.OutputFileName Like "*[\/:*?""<>|]*" Then problems.Add "OutputFileName contains a character not allowed in " & _
        "file names (\ / : * ? "" < > |): " & o.OutputFileName
    o.SizeClassFile = GetConfig("SizeClassFile")
    If Len(o.SizeClassFile) >= 2 And Left$(o.SizeClassFile, 1) = """" And Right$(o.SizeClassFile, 1) = """" Then
        o.SizeClassFile = Trim$(Mid$(o.SizeClassFile, 2, Len(o.SizeClassFile) - 2))     ' Explorer's "Copy as path" adds quotes
    End If
    o.CityColumn = GetConfig("MSAColumn")
    o.InvMonth = -1
    v = GetConfigValue("InventoryMonth")
    If VarType(v) = vbDate Then
        o.InvMonth = MonthKey(CDate(v))
    ElseIf Len(Trim$(SafeStr(v))) > 0 Then
        If TextToDate(SafeStr(v), d) Then
            o.InvMonth = MonthKey(d)
        Else
            problems.Add "InventoryMonth is not a date: """ & SafeStr(v) & """. Type any day in the month, like 8/1/2026, " & _
                         "or leave it blank for the latest month in the file."
        End If
    End If
    o.ShowCancelled = ReadBool("ShowCancelled", False, problems)
    o.ShowSSL = ReadBool("ShowSSL", False, problems)
    o.OpenAfterBuild = ReadBool("OpenAfterBuild", True, problems)
    ReadOptions = o
End Function

' Up to maxItems entries, one per line, each with prefix and cut to maxLen characters (0 = no cut).
Private Function ListText(items As Collection, ByVal prefix As String, ByVal maxItems As Long, ByVal maxLen As Long) As String
    Dim i As Long, s As String, t As String
    For i = 1 To items.Count
        If i > maxItems Then
            s = s & prefix & "... and " & (items.Count - maxItems) & " more" & vbLf
            Exit For
        End If
        t = CStr(items.Item(i))
        If maxLen > 0 And Len(t) > maxLen Then t = Left$(t, maxLen - 3) & "..."
        s = s & prefix & t & vbLf
    Next i
    ListText = s
End Function

' A message box shows about 1,024 characters; longer text is cut with a pointer to the Last run box.
Private Function ClipMessage(ByVal s As String) As String
    If Len(s) > 1000 Then s = Left$(s, 940) & vbLf & "... (the full list is in Last run on the Config sheet)"
    ClipMessage = s
End Function

Private Function CountText(ByVal n As Long, ByVal word As String) As String
    CountText = n & " " & word & IIf(n = 1, "", "s")
End Function

Private Function CityCountText(ByVal n As Long) As String
    CityCountText = n & IIf(n = 1, " MSA", " MSAs")
End Function

' "4: (no MSA), Amarillo, Brownsville, Dallas" - the MSAs of the build, alphabetical.
Private Function CityNames() As String
    Dim i As Long, s As String
    If Not mReady Or mNCities = 0 Then
        CityNames = "none found"
        Exit Function
    End If
    For i = 1 To mNCities
        s = s & IIf(i > 1, ", ", "") & mCityNames(SortedCity(i))
    Next i
    CityNames = mNCities & ": " & s
End Function

' "July 2026, August 2026 (the report shows August 2026)"
Private Function MonthsText() As String
    Dim i As Long, s As String
    If Not mReady Or mNMonths = 0 Then
        MonthsText = "none found"
        Exit Function
    End If
    For i = 1 To mNMonths
        s = s & IIf(i > 1, ", ", "") & MonthText(mMonths(i))
    Next i
    MonthsText = s & " (the report shows " & MonthText(mMonth) & ")"
End Function

' =====================================================================================================
'  THE SIZE CLASS FILE  (the master list: Name of machine -> Size Class)
' =====================================================================================================

' Reads the size class workbook: columns found by name (with aliases, ignoring case, spaces and underscores),
' usable rows kept (first row of a name wins), one warning per row skipped (mScWarnings). The class is taken as the
' file gives it (v2.7: the only source of a class) and must be one of the allowed classes (v2.8: its Set-up tab, see
' LoadClasses); a rated operating capacity column, if any, is not read.
Private Sub LoadSizeClassFile(ByVal path As String, problems As Collection)
    Dim t As SourceTable, errText As String, cName As Long, cCls As Long, r As Long, j As Long
    Dim nm As String, rawCls As String, cls As String, k As String, where As String, missing As String
    Dim firstRow() As Long, rowNum As Long
    mScName = FileNameOf(path)
    mScCount = 0
    mClsCount = 0
    mClsLine = ""
    mStep = "reading the size class file"
    If Not TryReadSource(path, t, errText, True) Then
        problems.Add "The size class file could not be read: " & errText
        Exit Sub
    End If
    If Not t.Ok Then
        problems.Add "The size class file " & mScName & " has no header row (a row with at least three filled cells)."
        Exit Sub
    End If
    cName = AliasColumn(t, "|NAMEOFMACHINE|MACHINE|MODEL|MODELGROUP|")
    cCls = AliasColumn(t, "|SIZECLASS|CLASS|")
    If cName = 0 Then missing = "Name of machine"
    If cCls = 0 Then missing = missing & IIf(Len(missing) > 0, " or ", "") & "Size Class"
    If Len(missing) > 0 Then
        problems.Add "The size class file " & mScName & " has no " & missing & " column (headers: " & HeaderList(t) & ")."
        Exit Sub
    End If
    If Not LoadClasses(t, problems) Then Exit Sub
    ReDim mScKeys(1 To t.DataRows + 1)
    ReDim mScFirst(1 To t.DataRows + 1)
    ReDim mScClass(1 To t.DataRows + 1)
    ReDim firstRow(1 To t.DataRows + 1)
    For r = t.HdrRow + 1 To t.LastRow
        If RowIsBlank(t, r) Then GoTo NextRow
        rowNum = t.FirstRow + r - 1
        nm = Trim$(SafeStr(t.Raw(r, cName)))
        rawCls = Trim$(SafeStr(t.Raw(r, cCls)))
        where = "Size class file row " & rowNum & IIf(Len(nm) > 0, " " & nm, "")
        If Len(nm) = 0 Then
            mScWarnings.Add where & ": no machine name - row skipped"
            GoTo NextRow
        End If
        If Len(rawCls) = 0 Then
            mScWarnings.Add where & ": no size class - row skipped"
            GoTo NextRow
        End If
        cls = CanonClass(rawCls)
        If Len(cls) = 0 Then
            mScWarnings.Add where & ": """ & rawCls & """ is not a size class - row skipped"
            GoTo NextRow
        End If
        k = UCase$(nm)
        For j = 1 To mScCount
            If mScKeys(j) = k Then
                mScWarnings.Add where & ": listed again (first in row " & firstRow(j) & ") - row " & firstRow(j) & " is used"
                GoTo NextRow
            End If
        Next j
        mScCount = mScCount + 1
        mScKeys(mScCount) = k
        If InStr(k, " ") > 0 Then mScFirst(mScCount) = Left$(k, InStr(k, " ") - 1) Else mScFirst(mScCount) = k
        mScClass(mScCount) = cls
        firstRow(mScCount) = rowNum
NextRow:
    Next r
    If mScCount = 0 Then problems.Add "The size class file " & mScName & " has no usable rows (a Name of machine and " & _
        "an allowed size class)."
End Sub

' The allowed size classes (v2.8): the size class file's Set-up tab, three tables found by their headers (CTL size
' class, MX size class, SSL size class), each read top to bottom below its header, CTL first, then MX, then SSL (their
' rank order); blank cells skipped, a class seen before (ignoring case) or the word UNMAPPED warned and ignored. A file
' without a Set-up tab gets the 17 built-in classes. False after a problem (added to problems).
Private Function LoadClasses(t As SourceTable, problems As Collection) As Boolean
    Dim heads As Variant, arr As Variant, f As Long, r As Long, c As Long, j As Long, v As String, where As String
    Dim hr(1 To 3) As Long, hc(1 To 3) As Long, nFam(1 To 3) As Long, firstAt() As String, missing As String
    mClsCount = 0
    If Not t.SetupFound Then
        arr = SizeClasses()
        ReDim mClsName(1 To UBound(arr) + 1)
        ReDim mClsFam(1 To UBound(arr) + 1)
        For j = 0 To UBound(arr)
            mClsCount = mClsCount + 1
            mClsName(mClsCount) = arr(j)
            If j < 6 Then
                mClsFam(mClsCount) = 1
            ElseIf j < 13 Then
                mClsFam(mClsCount) = 2
            Else
                mClsFam(mClsCount) = 3
            End If
        Next j
        mClsLine = "Size classes: the " & mClsCount & " built into the macro (" & mScName & " has no Set-up tab)"
        LoadClasses = True
        Exit Function
    End If
    heads = Array("CTLSIZECLASS", "MXSIZECLASS", "SSLSIZECLASS")
    For r = 1 To UBound(t.SetupRaw, 1)
        For c = 1 To UBound(t.SetupRaw, 2)
            v = UCase$(Replace(Replace(Trim$(SafeStr(t.SetupRaw(r, c))), " ", ""), "_", ""))
            For f = 1 To 3
                If hr(f) = 0 And v = heads(f - 1) Then
                    hr(f) = r
                    hc(f) = c
                End If
            Next f
        Next c
    Next r
    For f = 1 To 3
        If hr(f) = 0 Then missing = missing & IIf(Len(missing) > 0, ", ", "") & FamilyAbbr(f) & " size class"
    Next f
    If Len(missing) > 0 Then
        problems.Add "The Set-up tab of " & mScName & " has no " & missing & " header (it needs one table of size " & _
            "classes each for CTL, MX and SSL)."
        Exit Function
    End If
    ReDim mClsName(1 To MAX_CLASSES)
    ReDim mClsFam(1 To MAX_CLASSES)
    ReDim firstAt(1 To MAX_CLASSES)
    For f = 1 To 3
        c = hc(f)
        For r = hr(f) + 1 To UBound(t.SetupRaw, 1)
            v = Trim$(SafeStr(t.SetupRaw(r, c)))
            If Len(v) > 0 Then
                where = ColLetters(t.SetupCol + c - 1) & (t.SetupRow + r - 1)
                For j = 1 To mClsCount
                    If StrComp(mClsName(j), v, vbTextCompare) = 0 Then Exit For
                Next j
                If StrComp(v, "UNMAPPED", vbTextCompare) = 0 Then
                    mScWarnings.Add "Set-up tab " & where & ": """ & v & """ is not a size class name (UNMAPPED marks " & _
                        "models not in the file) - ignored"
                ElseIf j <= mClsCount Then
                    mScWarnings.Add "Set-up tab " & where & ": """ & v & """ is listed again (first in " & firstAt(j) & ") - ignored"
                ElseIf mClsCount = MAX_CLASSES Then
                    problems.Add "The Set-up tab of " & mScName & " lists more than " & MAX_CLASSES & " size classes."
                    Exit Function
                Else
                    mClsCount = mClsCount + 1
                    mClsName(mClsCount) = v
                    mClsFam(mClsCount) = f
                    firstAt(mClsCount) = where
                    nFam(f) = nFam(f) + 1
                End If
            End If
        Next r
    Next f
    If mClsCount = 0 Then
        problems.Add "The Set-up tab of " & mScName & " lists no size classes (type them under CTL size class, MX size " & _
            "class and SSL size class)."
        Exit Function
    End If
    mClsLine = "Size classes: " & mClsCount & " from the Set-up tab (CTL " & nFam(1) & ", MX " & nFam(2) & ", SSL " & nFam(3) & ")"
    LoadClasses = True
End Function

' Family texts of File Format Table 1, by family number (1 CTL, 2 MX, 3 SSL).
Private Function FamilyAbbr(ByVal f As Long) As String
    FamilyAbbr = Choose(f, "CTL", "MX", "SSL")
End Function

Private Function FamilyName(ByVal f As Long) As String
    FamilyName = Choose(f, CTL_FAMILY, MX_FAMILY, SSL_FAMILY)
End Function

Private Function FamilyRule(ByVal f As Long) As String
    FamilyRule = Choose(f, CTL_RULE_TEXT, MX_RULE_TEXT, SSL_RULE_TEXT)
End Function

' Column letters of a column number (1 = A, 27 = AA). Built in s: in VBA the function's own name on the right of an
' assignment would be a call to itself.
Private Function ColLetters(ByVal n As Long) As String
    Dim s As String
    Do While n > 0
        s = Chr$(65 + (n - 1) Mod 26) & s
        n = (n - 1) \ 26
    Loop
    ColLetters = s
End Function

' First column whose header, without spaces and underscores, is one of aliases ("|A|B|"); 0 if none.
Private Function AliasColumn(t As SourceTable, ByVal aliases As String) As Long
    Dim c As Long, h As String
    For c = 1 To t.ColCount
        h = Replace(Replace(t.Headers(c), " ", ""), "_", "")
        If Len(h) > 0 Then
            If InStr(aliases, "|" & h & "|") > 0 Then
                AliasColumn = c
                Exit Function
            End If
        End If
    Next c
End Function

' "MSA, ORDER_DATE, ..." - the non-blank headers of a table.
Private Function HeaderList(t As SourceTable) As String
    Dim c As Long, s As String
    For c = 1 To t.ColCount
        If Len(t.Headers(c)) > 0 Then s = s & IIf(Len(s) > 0, ", ", "") & t.Headers(c)
    Next c
    HeaderList = s
End Function

' The size class of a model group (trimmed, upper-cased): the file's row with that exact name, else the first
' row whose name starts with it as a word ("T650 Compact Track Loader"); "" if none.
Private Function LookupClass(ByVal key As String) As String
    Dim j As Long
    For j = 1 To mScCount
        If mScKeys(j) = key Then
            LookupClass = mScClass(j)
            Exit Function
        End If
    Next j
    For j = 1 To mScCount
        If mScFirst(j) = key Then
            LookupClass = mScClass(j)
            Exit Function
        End If
    Next j
End Function

' The 17 built-in size classes in report order (v2.7): CTL by rated operating capacity, MX by nominal tonnes, SSL by
' frame size. Since v2.8 they apply only to a size class file without a Set-up tab (LoadClasses).
Private Function SizeClasses() As Variant
    SizeClasses = Array("0<2000", "2000<2300", "2300<2600", "2600<3000", "3000<3500", "3500&Over", _
                        ">00-01", ">01-02", ">02-03", ">03-04", ">04-05", ">05-06", ">06-10", _
                        "EX Small", "Small Frame", "Med Frame", "Large Frame")
End Function

' The allowed spelling of a size class (matched ignoring case after Trim), "" if it is not allowed.
Private Function CanonClass(ByVal cls As String) As String
    Dim i As Long
    For i = 1 To mClsCount
        If StrComp(mClsName(i), Trim$(cls), vbTextCompare) = 0 Then
            CanonClass = mClsName(i)
            Exit Function
        End If
    Next i
End Function

' Rank of a class: its place among the allowed classes; UNMAPPED (anything else) ranks after them all.
Private Function ClassRank(ByVal cls As String) As Long
    Dim i As Long
    For i = 1 To mClsCount
        If mClsName(i) = cls Then
            ClassRank = i
            Exit Function
        End If
    Next i
    ClassRank = UNMAPPED_RANK
End Function

' The order status order (plan 5.3.3), as printed on the File Format tab.
Private Function StatusOrder() As Variant
    StatusOrder = Array("ORDERED", "Not yet processed", "INVOICED", "Completely processed", "(any other status)", "CANCELLED")
End Function

' Rank of a trimmed, upper-cased status: ORDERED 1, NOT YET PROCESSED 2, INVOICED 3, COMPLETELY PROCESSED 4,
' any other 5, CANCELLED 6.
Private Function StatusRank(ByVal s As String) As Long
    Select Case s
        Case "ORDERED": StatusRank = 1
        Case "NOT YET PROCESSED": StatusRank = 2
        Case "INVOICED": StatusRank = 3
        Case "COMPLETELY PROCESSED": StatusRank = 4
        Case "CANCELLED": StatusRank = 6
        Case Else: StatusRank = 5
    End Select
End Function

Private Function ProductTypeName(ByVal i As Long) As String
    ProductTypeName = Choose(i + 1, "Compact Track Loaders", "Mini Excavators", "Skid-Steer Loaders")
End Function

' =====================================================================================================
'  INPUT DISCOVERY AND READING
' =====================================================================================================

' Fills mFiles with the export files of folder in alphabetical order of file name (case-insensitive). A file's
' kind is the start of its name; earlier reports (001..., ...Analysis By MSA...) and lock files are skipped.
' mFiles is sized once with a plain ReDim, so nothing of an earlier run is carried over.
Private Sub DiscoverFiles(ByVal folder As String)
    Dim fileNames() As String, nNames As Long, f As String, tmpName As String
    Dim baseName As String, upperBase As String, city As String, kind As String, prefixLen As Long
    Dim i As Long, j As Long, p As Long, dup As Boolean, n As Long
    Dim kinds() As String, cities() As String, picked() As String

    mNFiles = 0
    #If Mac Then
        f = Dir(folder & Application.PathSeparator)
    #Else
        f = Dir(folder & Application.PathSeparator & "*")
    #End If
    Do While Len(f) > 0
        If Left$(f, 2) <> "~$" And IsExcelFileName(f) Then
            dup = False
            For i = 1 To nNames
                If StrComp(fileNames(i), f, vbTextCompare) = 0 Then dup = True: Exit For
            Next i
            If Not dup Then
                nNames = nNames + 1
                ReDim Preserve fileNames(1 To nNames)
                fileNames(nNames) = f
            End If
        End If
        f = Dir
    Loop
    If nNames = 0 Then Exit Sub

    For i = 2 To nNames                                  ' insertion sort, case-insensitive
        tmpName = fileNames(i)
        j = i - 1
        Do While j >= 1
            If UCase$(fileNames(j)) <= UCase$(tmpName) Then Exit Do
            fileNames(j + 1) = fileNames(j)
            j = j - 1
        Loop
        fileNames(j + 1) = tmpName
    Next i

    ReDim kinds(1 To nNames)
    ReDim cities(1 To nNames)
    ReDim picked(1 To nNames)
    For i = 1 To nNames
        baseName = fileNames(i)
        p = InStrRev(baseName, ".")
        If p > 0 Then baseName = Left$(baseName, p - 1)
        baseName = Trim$(baseName)
        upperBase = UCase$(baseName)
        kind = ""
        If Left$(upperBase, 3) = "001" Or InStr(upperBase, "ANALYSIS BY MSA") > 0 Then
            kind = ""                                    ' an earlier report, not an input
        ElseIf Left$(upperBase, 12) = "ORDER INTAKE" Then
            kind = "OI": prefixLen = 12
        ElseIf Left$(upperBase, 18) = "INVENTORY BY MONTH" Then
            kind = "INV": prefixLen = 18
        ElseIf Left$(upperBase, 11) = "OPEN ORDERS" Then
            kind = "OO": prefixLen = 11
        ElseIf Left$(upperBase, 10) = "OPEN ORDER" Then
            kind = "OO": prefixLen = 10
        End If
        If Len(kind) > 0 Then
            city = Mid$(baseName, prefixLen + 1)
            p = InStr(city, "_")                         ' "Austin_ANONYMIZED" -> "Austin"
            If p > 0 Then city = Left$(city, p - 1)
            city = Trim$(city)
            Do While Len(city) > 0 And (Left$(city, 1) = "-" Or Left$(city, 1) = " ")
                city = Mid$(city, 2)
            Loop
            n = n + 1
            picked(n) = fileNames(i)
            kinds(n) = kind
            cities(n) = NormCity(city)
        End If
    Next i
    If n = 0 Then Exit Sub
    ReDim mFiles(1 To n)
    For i = 1 To n
        mFiles(i).Path = folder & Application.PathSeparator & picked(i)
        mFiles(i).FileName = picked(i)
        mFiles(i).Kind = kinds(i)
        mFiles(i).FileCity = cities(i)
    Next i
    mNFiles = n
End Sub

Private Function IsExcelFileName(ByVal f As String) As Boolean
    Dim p As Long, ext As String
    p = InStrRev(f, ".")
    If p = 0 Then Exit Function
    ext = LCase$(Mid$(f, p + 1))
    IsExcelFileName = (ext = "xlsx" Or ext = "xlsm" Or ext = "xls")
End Function

' The required columns of file i and its city column. A CityColumn that is not a header, or a file with
' neither a city column nor a city in its name, is a problem (cityOk = False).
Private Sub ResolveColumns(ByVal i As Long, opt As ReportOptions, problems As Collection, ByRef cityOk As Boolean)
    Dim names As Variant, k As Long, idx() As Long
    Select Case mFiles(i).Kind
        Case "OI": names = Array("ORDER_DATE", "REQUESTED_DATE", "ORDER_STATUS", "IS_OPEN_ORDER", "SUM(F.ORDERED_QUANTITY)", _
                                 "PRODUCT_TYPE_ABBR", "MODEL_GROUP", "MODEL_NUMBER", "ORGANIZATION_NAME")
        Case "INV": names = Array("PRODUCT_TYPE", "MODEL_GROUP", "MODEL_NUMBER", "INVENTORY_DATE", "TOTAL")
        Case Else: names = Array()
    End Select
    mFiles(i).Missing = ""
    If UBound(names) >= 0 Then
        ReDim idx(0 To UBound(names))
        For k = 0 To UBound(names)
            idx(k) = HeaderIndex(mFiles(i).Tbl, CStr(names(k)))
            If idx(k) = 0 Then mFiles(i).Missing = mFiles(i).Missing & IIf(Len(mFiles(i).Missing) > 0, ", ", "") & names(k)
        Next k
        mFiles(i).Idx = idx
    End If
    If mFiles(i).Kind = "INV" Then
        mFiles(i).OrgCol = HeaderIndex(mFiles(i).Tbl, "ORGANIZATION_NAME")
        mFiles(i).CustCol = HeaderIndex(mFiles(i).Tbl, "ORGANIZATION_CUSTOMER_NUMBER")
    End If
    mFiles(i).CityCol = FindCityColumn(mFiles(i).Tbl, opt.CityColumn)
    If mFiles(i).CityCol > 0 Then
        mFiles(i).CityHdr = mFiles(i).Tbl.Headers(mFiles(i).CityCol)
    ElseIf Len(opt.CityColumn) > 0 Then
        problems.Add "MSAColumn """ & opt.CityColumn & """ is not a header of " & mFiles(i).FileName & _
                     " (headers: " & HeaderList(mFiles(i).Tbl) & ")."
        cityOk = False
    ElseIf Len(mFiles(i).FileCity) = 0 Then
        problems.Add "No MSA column found in " & mFiles(i).FileName & " (headers: " & HeaderList(mFiles(i).Tbl) & "). " & _
                     "Type the header name into MSAColumn."
        cityOk = False
    End If
End Sub

' The MSA column: Config MSAColumn when set (that header, ignoring case); else the first header, in this order
' of tests, that equals MSA, contains MSA, equals CITY, contains CITY, METRO or MARKET (ignoring spaces and
' underscores). 0 = none.
Private Function FindCityColumn(t As SourceTable, ByVal cityColumn As String) As Long
    Dim test As Long, c As Long, h As String, hit As Boolean
    If Len(cityColumn) > 0 Then
        FindCityColumn = HeaderIndex(t, UCase$(Trim$(cityColumn)))
        Exit Function
    End If
    For test = 1 To 6
        For c = 1 To t.ColCount
            h = Replace(Replace(t.Headers(c), " ", ""), "_", "")
            If Len(h) > 0 Then
                Select Case test
                    Case 1: hit = (h = "MSA")
                    Case 2: hit = (InStr(h, "MSA") > 0)
                    Case 3: hit = (h = "CITY")
                    Case 4: hit = (InStr(h, "CITY") > 0)
                    Case 5: hit = (InStr(h, "METRO") > 0)
                    Case 6: hit = (InStr(h, "MARKET") > 0)
                End Select
                If hit Then
                    FindCityColumn = c
                    Exit Function
                End If
            End If
        Next c
    Next test
End Function

' ReadSource that turns an error (the file is open from another place, password protected, damaged) into errText.
Private Function TryReadSource(ByVal path As String, t As SourceTable, ByRef errText As String, _
                               Optional ByVal withSetup As Boolean = False) As Boolean
    On Error GoTo Failed
    t = ReadSource(path, withSetup)
    TryReadSource = True
    Exit Function
Failed:
    errText = Err.Description
End Function

' Copies the first worksheet of a workbook into memory. The file is opened read-only and closed again,
' unless it is already open in Excel, in which case it is read where it is and left open. A workbook with
' the same file name open from another place (another folder, or a OneDrive/SharePoint address) stops with
' an error: Excel cannot open both, and closing it could lose the user's edits.
Private Function ReadSource(ByVal path As String, Optional ByVal withSetup As Boolean = False) As SourceTable
    Dim t As SourceTable, wb As Workbook, wbx As Workbook, raw As Variant, srcName As String, nBefore As Long
    Dim r As Long, c As Long, filled As Long, wasOpen As Boolean, hdrs() As String, errNum As Long, errDesc As String
    Dim wsx As Worksheet, one(1 To 1, 1 To 1) As Variant
    On Error GoTo Fail
    srcName = FileNameOf(path)
    For Each wbx In Application.Workbooks
        If StrComp(wbx.Name, srcName, vbTextCompare) = 0 Then
            If StrComp(wbx.FullName, path, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 4, , _
                srcName & " is open in Excel from " & wbx.FullName & ". Close it and run again."
            Set wb = wbx
            wasOpen = True
            Exit For
        End If
    Next wbx
    If wb Is Nothing Then
        nBefore = Application.Workbooks.Count
        Set wb = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
        If Application.Workbooks.Count = nBefore Then wasOpen = True     ' Excel handed back a copy already open
    End If
    t.FirstRow = wb.Worksheets(1).UsedRange.Row
    raw = wb.Worksheets(1).UsedRange.Value
    If withSetup Then                                    ' v2.8: the size class file's Set-up tab, if it has one
        For Each wsx In wb.Worksheets
            If UCase$(Replace(Replace(Replace(wsx.Name, " ", ""), "-", ""), "_", "")) = "SETUP" Then
                t.SetupFound = True
                t.SetupRow = wsx.UsedRange.Row
                t.SetupCol = wsx.UsedRange.Column
                t.SetupRaw = wsx.UsedRange.Value
                If Not IsArray(t.SetupRaw) Then          ' one used cell: Value is a scalar
                    one(1, 1) = t.SetupRaw
                    t.SetupRaw = one
                End If
                Exit For
            End If
        Next wsx
    End If
    If Not wasOpen Then wb.Close SaveChanges:=False
    Set wb = Nothing
    If Not IsArray(raw) Then GoTo Done                   ' empty sheet: UsedRange is one cell, Value a scalar
    t.Raw = raw
    t.ColCount = UBound(raw, 2)
    t.LastRow = UBound(raw, 1)
    ' header = first row with at least three non-blank cells
    For r = 1 To t.LastRow
        filled = 0
        For c = 1 To t.ColCount
            If Len(Trim$(SafeStr(raw(r, c)))) > 0 Then filled = filled + 1
        Next c
        If filled >= 3 Then t.HdrRow = r: Exit For
    Next r
    If t.HdrRow = 0 Then GoTo Done
    ReDim hdrs(1 To t.ColCount)
    For c = 1 To t.ColCount
        hdrs(c) = UCase$(Trim$(SafeStr(raw(t.HdrRow, c))))
    Next c
    t.Headers = hdrs
    ' drop trailing blank rows, count the rows below the header (blank ones too) and the non-blank ones
    Do While t.LastRow > t.HdrRow
        If Not RowIsBlank(t, t.LastRow) Then Exit Do
        t.LastRow = t.LastRow - 1
    Loop
    t.NRead = t.LastRow - t.HdrRow
    For r = t.HdrRow + 1 To t.LastRow
        If Not RowIsBlank(t, r) Then t.DataRows = t.DataRows + 1
    Next r
    t.Ok = True
Done:
    ReadSource = t                                       ' on every path (an empty sheet returns t.Ok = False)
    Exit Function
Fail:                                                    ' close a workbook this call opened, then pass the error on
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If Not wasOpen And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise errNum, , errDesc
End Function

Private Function RowIsBlank(t As SourceTable, ByVal r As Long) As Boolean
    Dim c As Long
    For c = 1 To t.ColCount
        If Len(Trim$(SafeStr(t.Raw(r, c)))) > 0 Then Exit Function
    Next c
    RowIsBlank = True
End Function

Private Function HeaderIndex(t As SourceTable, ByVal colName As String) As Long
    Dim c As Long
    For c = 1 To t.ColCount
        If t.Headers(c) = UCase$(colName) Then HeaderIndex = c: Exit Function
    Next c
    HeaderIndex = 0
End Function

' =====================================================================================================
'  PREPARING THE ROWS  (every rule that decides what a row holds; the writers only lay them out)
' =====================================================================================================

Private Sub PrepareData(opt As ReportOptions)
    Dim i As Long, r As Long, k As Long, n As Long, cityIdx As Long, ti As Long, nTotal As Long
    Dim rawAbbr As String, abbr As String, mg As String, cls As String, v As Variant, uMg As String
    Dim present() As Boolean, m As Long, j As Long, tmp As Long, rowNum As Long

    mNCities = 0
    mLastCity = 0
    mUnmappedAll = ""
    ReDim mCityKeys(1 To 8)
    ReDim mCityNames(1 To 8)
    ReDim mUnmapped(1 To 8)
    mNDrop = 0
    ReDim mDropFile(1 To 8)
    ReDim mDropLabel(1 To 8)
    ReDim mDropCount(1 To 8)
    For i = 1 To mNFiles
        mFiles(i).Written = 0
        mFiles(i).BlankCity = 0
        mFiles(i).CityKeys = Chr$(1)
        mFiles(i).NCities = 0
    Next i

    ' Order Intake: rows of every usable file, files in alphabetical order, rows in order
    nTotal = 0
    For i = 1 To mNFiles
        If FileUsable(i, "OI") Then nTotal = nTotal + mFiles(i).Tbl.DataRows
    Next i
    ReDim mOI(1 To nTotal + 1, 1 To 11)
    ReDim mOICity(1 To nTotal + 1)
    ReDim mOIStatus(1 To nTotal + 1)
    ReDim mOIDay(1 To nTotal + 1)
    mNOI = 0
    For i = 1 To mNFiles
        If FileUsable(i, "OI") Then
            With mFiles(i).Tbl
                For r = .HdrRow + 1 To .LastRow
                    If RowIsBlank(mFiles(i).Tbl, r) Then GoTo NextOI
                    rawAbbr = Trim$(SafeStr(.Raw(r, mFiles(i).Idx(5))))
                    abbr = UCase$(rawAbbr)
                    If Not (abbr = "CTL" Or abbr = "MX" Or abbr = "SSL") Then
                        AddDrop i, IIf(Len(rawAbbr) = 0, "(blank)", rawAbbr)
                        GoTo NextOI
                    End If
                    mFiles(i).Written = mFiles(i).Written + 1
                    cityIdx = RowCity(i, r, True)
                    mNOI = mNOI + 1
                    n = mNOI
                    For k = 0 To 8
                        mOI(n, k + 3) = .Raw(r, mFiles(i).Idx(k))
                    Next k
                    mOI(n, 8) = abbr
                    rowNum = .FirstRow + r - 1
                    For k = 0 To 1                       ' ORDER_DATE, REQUESTED_DATE stored as text
                        v = mOI(n, k + 3)
                        Select Case DateStatus(v)
                            Case 1
                                mOI(n, k + 3) = v
                                mFiles(i).Conv(k) = mFiles(i).Conv(k) + 1
                            Case 2
                                mFiles(i).NBad(k) = mFiles(i).NBad(k) + 1
                                AddRowNumber mFiles(i).BadRows(k), mFiles(i).NBad(k), rowNum
                        End Select
                    Next k
                    mg = Trim$(SafeStr(.Raw(r, mFiles(i).Idx(6))))
                    uMg = UCase$(mg)
                    cls = LookupClass(uMg)
                    If Len(cls) = 0 Then                      ' not in the file: UNMAPPED, whatever the type or name
                        cls = "UNMAPPED"
                        AddToList mUnmapped(cityIdx), IIf(Len(mg) = 0, "(blank)", mg)
                        AddToList mUnmappedAll, IIf(Len(mg) = 0, "(blank)", mg)
                    End If
                    mOI(n, 1) = cls
                    mOI(n, 2) = cityIdx
                    mOICity(n) = cityIdx
                    mOIStatus(n) = StatusRank(UCase$(Trim$(SafeStr(mOI(n, 5)))))
                    If VarType(mOI(n, 3)) = vbDate Then mOIDay(n) = Int(CDbl(mOI(n, 3))) Else mOIDay(n) = 0
NextOI:
                Next r
            End With
        End If
    Next i

    ' Inventory by Month: the three product types, written with their canonical names
    nTotal = 0
    For i = 1 To mNFiles
        If FileUsable(i, "INV") Then nTotal = nTotal + mFiles(i).Tbl.DataRows
    Next i
    ReDim mINV(1 To nTotal + 1, 1 To 7)
    ReDim mINVCity(1 To nTotal + 1)
    ReDim mINVType(1 To nTotal + 1)
    ReDim mINVMonth(1 To nTotal + 1)
    ReDim mINVDay(1 To nTotal + 1)
    mNINV = 0
    For i = 1 To mNFiles
        If FileUsable(i, "INV") Then
            With mFiles(i).Tbl
                For r = .HdrRow + 1 To .LastRow
                    If RowIsBlank(mFiles(i).Tbl, r) Then GoTo NextInv
                    rawAbbr = Trim$(SafeStr(.Raw(r, mFiles(i).Idx(0))))
                    ti = -1
                    For k = 0 To 2
                        If StrComp(rawAbbr, ProductTypeName(k), vbTextCompare) = 0 Then ti = k
                    Next k
                    If ti < 0 Then
                        AddDrop i, IIf(Len(rawAbbr) = 0, "(blank)", rawAbbr)
                        GoTo NextInv
                    End If
                    mFiles(i).Written = mFiles(i).Written + 1
                    cityIdx = RowCity(i, r, True)
                    v = .Raw(r, mFiles(i).Idx(3))
                    Select Case DateStatus(v)
                        Case 1
                            mFiles(i).Conv(0) = mFiles(i).Conv(0) + 1
                        Case 2
                            mFiles(i).NBad(0) = mFiles(i).NBad(0) + 1
                            AddRowNumber mFiles(i).BadRows(0), mFiles(i).NBad(0), .FirstRow + r - 1
                    End Select
                    mNINV = mNINV + 1
                    n = mNINV
                    mINV(n, 1) = ProductTypeName(ti)
                    mINV(n, 2) = .Raw(r, mFiles(i).Idx(1))
                    mINV(n, 3) = .Raw(r, mFiles(i).Idx(2))
                    mINV(n, 4) = v
                    mINV(n, 5) = .Raw(r, mFiles(i).Idx(4))
                    If mFiles(i).OrgCol > 0 Then mINV(n, 6) = .Raw(r, mFiles(i).OrgCol)
                    If mFiles(i).CustCol > 0 Then mINV(n, 7) = .Raw(r, mFiles(i).CustCol)
                    mINVCity(n) = cityIdx
                    mINVType(n) = ti
                    If VarType(v) = vbDate Then
                        mINVMonth(n) = MonthKey(CDate(v))
                        mINVDay(n) = Int(CDbl(v))
                    Else
                        mINVMonth(n) = -1
                        mINVDay(n) = 0
                    End If
NextInv:
                Next r
            End With
        End If
    Next i

    ' Open Order files: only their cities are counted (for the log line); they feed nothing
    For i = 1 To mNFiles
        If mFiles(i).Kind = "OO" And mFiles(i).Tbl.Ok Then
            For r = mFiles(i).Tbl.HdrRow + 1 To mFiles(i).Tbl.LastRow
                If Not RowIsBlank(mFiles(i).Tbl, r) Then cityIdx = RowCity(i, r, False)
            Next r
        End If
    Next i

    ' the cities in alphabetical order (of their upper-case key)
    ReDim mCityOrder(1 To mNCities + 1)
    ReDim mCityRank(1 To mNCities + 1)
    For i = 1 To mNCities
        mCityOrder(i) = i
    Next i
    For i = 2 To mNCities
        tmp = mCityOrder(i)
        j = i - 1
        Do While j >= 1
            If mCityKeys(mCityOrder(j)) <= mCityKeys(tmp) Then Exit Do
            mCityOrder(j + 1) = mCityOrder(j)
            j = j - 1
        Loop
        mCityOrder(j + 1) = tmp
    Next i
    For i = 1 To mNCities
        mCityRank(mCityOrder(i)) = i
    Next i

    ' the months: every month in the inventory rows, the one shown (Config, else the latest) added if absent
    mMonth = opt.InvMonth
    mNMonths = 0
    For n = 1 To mNINV
        If mINVMonth(n) > mMonth And opt.InvMonth < 0 Then mMonth = mINVMonth(n)
    Next n
    If mNINV > 0 Or mMonth >= 0 Then
        ReDim mMonths(1 To mNINV + 1)
        For n = 1 To mNINV + 1
            If n <= mNINV Then m = mINVMonth(n) Else m = mMonth
            If m >= 0 Then
                For j = 1 To mNMonths
                    If mMonths(j) = m Then Exit For
                Next j
                If j > mNMonths Then
                    mNMonths = mNMonths + 1
                    mMonths(mNMonths) = m
                End If
            End If
        Next n
        For i = 2 To mNMonths                        ' ascending
            tmp = mMonths(i)
            j = i - 1
            Do While j >= 1
                If mMonths(j) <= tmp Then Exit Do
                mMonths(j + 1) = mMonths(j)
                j = j - 1
            Loop
            mMonths(j + 1) = tmp
        Next i
    End If
End Sub

Private Function FileUsable(ByVal i As Long, ByVal kind As String) As Boolean
    FileUsable = (mFiles(i).Kind = kind And mFiles(i).Tbl.Ok And Len(mFiles(i).Missing) = 0)
End Function

' The city of row r of file i as a registry index (0 when not registered, for Open Order files): the city
' column's text (trimmed, spaces collapsed; blank = (no city)), else the file name's city. The first spelling
' of a city wins; the file counts its distinct cities.
Private Function RowCity(ByVal i As Long, ByVal r As Long, ByVal register As Boolean) As Long
    Dim c As String, k As String, j As Long
    If mFiles(i).CityCol > 0 Then
        c = NormCity(SafeStr(mFiles(i).Tbl.Raw(r, mFiles(i).CityCol)))
        If Len(c) = 0 Then
            mFiles(i).BlankCity = mFiles(i).BlankCity + 1
            c = NO_CITY
        End If
    Else
        c = mFiles(i).FileCity
    End If
    k = UCase$(c)
    If InStr(1, mFiles(i).CityKeys, Chr$(1) & k & Chr$(1), vbBinaryCompare) = 0 Then
        mFiles(i).CityKeys = mFiles(i).CityKeys & k & Chr$(1)
        mFiles(i).NCities = mFiles(i).NCities + 1
    End If
    If Not register Then Exit Function
    If mLastCity > 0 Then
        If mCityKeys(mLastCity) = k Then
            RowCity = mLastCity
            Exit Function
        End If
    End If
    For j = 1 To mNCities
        If mCityKeys(j) = k Then
            mLastCity = j
            RowCity = j
            Exit Function
        End If
    Next j
    mNCities = mNCities + 1
    If mNCities > UBound(mCityKeys) Then
        ReDim Preserve mCityKeys(1 To 2 * mNCities)
        ReDim Preserve mCityNames(1 To 2 * mNCities)
        ReDim Preserve mUnmapped(1 To 2 * mNCities)
    End If
    mCityKeys(mNCities) = k
    mCityNames(mNCities) = c
    mLastCity = mNCities
    RowCity = mNCities
End Function

' Counts a dropped row of file i under its product type label (labels in the order first seen).
Private Sub AddDrop(ByVal i As Long, ByVal label As String)
    Dim j As Long
    For j = 1 To mNDrop
        If mDropFile(j) = i And mDropLabel(j) = label Then
            mDropCount(j) = mDropCount(j) + 1
            Exit Sub
        End If
    Next j
    mNDrop = mNDrop + 1
    If mNDrop > UBound(mDropFile) Then
        ReDim Preserve mDropFile(1 To 2 * mNDrop)
        ReDim Preserve mDropLabel(1 To 2 * mNDrop)
        ReDim Preserve mDropCount(1 To 2 * mNDrop)
    End If
    mDropFile(mNDrop) = i
    mDropLabel(mNDrop) = label
    mDropCount(mNDrop) = 1
End Sub

' Adds s to a "|"-separated list unless it is there already (case-sensitive), keeping the order first seen.
Private Sub AddToList(ByRef list As String, ByVal s As String)
    If InStr(1, "|" & list & "|", "|" & s & "|", vbBinaryCompare) = 0 Then list = list & IIf(Len(list) > 0, "|", "") & s
End Sub

' Registry index of the i-th city in alphabetical order.
Private Function SortedCity(ByVal i As Long) As Long
    SortedCity = mCityOrder(i)
End Function

' Sort orders (stable): Order Intake by size class, status, ORDER_DATE newest first (undated last), city;
' inventory by city, product type, INVENTORY_DATE newest first (undated last). Ties keep the export order.
Private Sub SortData()
    Dim sortKey() As Double, n As Long
    ReDim sortKey(1 To mNOI + 1)
    For n = 1 To mNOI
        sortKey(n) = ((ClassRank(CStr(mOI(n, 1))) * 10# + mOIStatus(n)) * 10000000# + (3000000# - mOIDay(n))) * 10000# + _
                     mCityRank(mOICity(n))
    Next n
    If mNOI > 0 Then StableSortIndex sortKey, mNOI, mOIOrd
    ReDim sortKey(1 To mNINV + 1)
    For n = 1 To mNINV
        sortKey(n) = (mCityRank(mINVCity(n)) * 10# + mINVType(n)) * 10000000# + (3000000# - mINVDay(n))
    Next n
    If mNINV > 0 Then StableSortIndex sortKey, mNINV, mINVOrd
End Sub

' =====================================================================================================
'  RUN LOG LINES
' =====================================================================================================

' The line of file i: rows read, written, dropped by type, blank, and the city column (or why it is not used).
Private Function FileSummary(ByVal i As Long) As String
    Dim s As String, j As Long, nd As Long, blank As Long, types As String
    With mFiles(i)
        If Not .Tbl.Ok Then
            FileSummary = .FileName & " is empty - not used"
            Exit Function
        End If
        If Len(.Missing) > 0 Then
            FileSummary = .FileName & " is missing column(s) " & .Missing & " - not used"
            Exit Function
        End If
        If .Kind = "OO" Then
            FileSummary = .FileName & ": " & CountText(.Tbl.DataRows, "row") & ", " & FileCityText(i) & " - not used by this report"
            Exit Function
        End If
        s = .FileName & ": " & CountText(.Tbl.NRead, "row") & " -> " & .Written & " written"
        For j = 1 To mNDrop
            If mDropFile(j) = i Then
                nd = nd + mDropCount(j)
                types = types & IIf(Len(types) > 0, ", ", "") & mDropLabel(j) & " " & mDropCount(j)
            End If
        Next j
        If nd > 0 Then s = s & ", " & nd & " dropped (" & types & ")"
        blank = .Tbl.NRead - .Tbl.DataRows
        If blank > 0 Then s = s & ", " & blank & " blank ignored"
        FileSummary = s & "; " & FileCityText(i)
    End With
End Function

Private Function FileCityText(ByVal i As Long) As String
    If mFiles(i).CityCol > 0 Then
        FileCityText = "MSA column " & mFiles(i).CityHdr & ", " & CityCountText(mFiles(i).NCities)
    Else
        FileCityText = "MSA from file name (" & mFiles(i).FileCity & ")"
    End If
End Function

' Every Run Log line from row 4, in the order of plan Section 5.2.4 (build_report_reference.py writes the same).
Private Sub BuildLogLines(opt As ReportOptions)
    Dim i As Long, c As Long, n As Long, ci As Long, kinds As Variant, labels As Variant, k As Long, names As String
    Dim nOI() As Long, nType() As Long, nCanc() As Long, nHid() As Long, nInv() As Long, units() As Double
    Dim s As String, mText As String, hasInv As Boolean, inMonth As Boolean
    Dim nOpen() As Long, nOrdered() As Long, openUnits() As Double, st As String, isOpen As Boolean
    Dim nSpaced As Long, nSpacedTexts As Long, spacedKeys As String, spacedText As String, nDisagree As Long
    Set mLog = New Collection
    For k = 0 To 1
        For i = 1 To mNFiles
            If mFiles(i).Kind = Choose(k + 1, "OI", "INV") Then LogLine "File " & FileSummary(i)
        Next i
    Next k
    LogLine "Size class file " & mScName & ": " & CountText(mScCount, "usable row") & ", " & CountText(mScWarnings.Count, "warning")
    LogLine mClsLine
    For i = 1 To mScWarnings.Count
        LogLine CStr(mScWarnings.Item(i))
    Next i

    ReDim nOI(1 To mNCities + 1)
    ReDim nType(1 To mNCities + 1, 0 To 2)
    ReDim nCanc(1 To mNCities + 1)
    ReDim nHid(1 To mNCities + 1)
    ReDim nInv(1 To mNCities + 1)
    ReDim units(1 To mNCities + 1, 0 To 2)
    ReDim nOpen(1 To mNCities + 1)
    ReDim nOrdered(1 To mNCities + 1)
    ReDim openUnits(1 To mNCities + 1, 0 To 2)
    mOpenRows = 0
    For k = 0 To 2
        mOpenUnits(k) = 0
    Next k
    spacedKeys = Chr$(1)
    For n = 1 To mNOI
        ci = mOICity(n)
        nOI(ci) = nOI(ci) + 1
        k = AbbrIndex(CStr(mOI(n, 8)))
        nType(ci, k) = nType(ci, k) + 1
        If mOIStatus(n) = 6 Then nCanc(ci) = nCanc(ci) + 1
        If OIHidden(n, opt) Then nHid(ci) = nHid(ci) + 1
        ' open orders: the rows the SUMIFS on the Open Orders by MSA tab count (ORDER_STATUS ORDERED or Not yet
        ' processed, the whole text, ignoring case, not ignoring stray spaces); all three product types
        st = UCase$(SafeStr(mOI(n, 5)))
        isOpen = (st = "ORDERED" Or st = "NOT YET PROCESSED")
        If isOpen Then
            nOpen(ci) = nOpen(ci) + 1
            mOpenRows = mOpenRows + 1
            If st = "ORDERED" Then nOrdered(ci) = nOrdered(ci) + 1
            If IsNumber(mOI(n, 7)) Then
                openUnits(ci, k) = openUnits(ci, k) + CDbl(mOI(n, 7))
                mOpenUnits(k) = mOpenUnits(k) + CDbl(mOI(n, 7))
            End If
        ElseIf Trim$(st) = "ORDERED" Or Trim$(st) = "NOT YET PROCESSED" Then
            nSpaced = nSpaced + 1                        ' stray spaces: SUMIFS will not count the row
            If InStr(1, spacedKeys, Chr$(1) & SafeStr(mOI(n, 5)) & Chr$(1), vbBinaryCompare) = 0 Then
                spacedKeys = spacedKeys & SafeStr(mOI(n, 5)) & Chr$(1)
                nSpacedTexts = nSpacedTexts + 1
                If nSpacedTexts <= MAX_ROWS_LISTED Then spacedText = spacedText & IIf(Len(spacedText) > 0, ", ", "") & _
                    """" & SafeStr(mOI(n, 5)) & """"
            End If
        End If
        If VarType(mOI(n, 6)) = vbBoolean Then           ' IS_OPEN_ORDER: a Boolean that disagrees with the status
            If CBool(mOI(n, 6)) <> isOpen Then nDisagree = nDisagree + 1
        End If
    Next n
    For n = 1 To mNINV
        ci = mINVCity(n)
        nInv(ci) = nInv(ci) + 1
        If mMonth >= 0 And mINVMonth(n) = mMonth Then
            If IsNumber(mINV(n, 5)) Then units(ci, mINVType(n)) = units(ci, mINVType(n)) + CDbl(mINV(n, 5))
            inMonth = True
        End If
    Next n
    If mMonth >= 0 Then mText = MonthText(mMonth) Else mText = "(no month)"
    For c = 1 To mNCities
        ci = SortedCity(c)
        If nOI(ci) > 0 Then
            s = mCityNames(ci) & ": Order Intake " & CountText(nOI(ci), "row") & " (CTL " & nType(ci, 0) & ", MX " & nType(ci, 1) & _
                ", SSL " & nType(ci, 2) & "; " & nCanc(ci) & " CANCELLED), " & nHid(ci) & " hidden by the filters"
            If Len(mUnmapped(ci)) > 0 Then s = s & "; UNMAPPED model groups: " & Replace(mUnmapped(ci), "|", ", ")
            LogLine s
        Else
            LogLine mCityNames(ci) & ": no Order Intake rows"
        End If
        If nInv(ci) > 0 Then
            LogLine mCityNames(ci) & ": Inventory " & CountText(nInv(ci), "row") & "; " & mText & ": CTL " & CStr(units(ci, 0)) & _
                    ", MX " & CStr(units(ci, 1)) & ", SSL " & CStr(units(ci, 2)) & " units"
        Else
            LogLine mCityNames(ci) & ": no inventory rows"
        End If
        If nOpen(ci) > 0 Then
            LogLine mCityNames(ci) & ": Open orders " & CountText(nOpen(ci), "row") & " (ORDERED " & nOrdered(ci) & _
                    ", Not yet processed " & (nOpen(ci) - nOrdered(ci)) & "): CTL " & CStr(openUnits(ci, 0)) & ", MX " & _
                    CStr(openUnits(ci, 1)) & ", SSL " & CStr(openUnits(ci, 2)) & " units"
        Else
            LogLine mCityNames(ci) & ": no open orders (ORDERED or Not yet processed)"
        End If
    Next c
    For i = 1 To mNFiles
        If mFiles(i).Kind = "OO" Then LogLine "File " & FileSummary(i)
    Next i
    For k = 0 To 1
        For i = 1 To mNFiles
            If FileUsable(i, Choose(k + 1, "OI", "INV")) Then
                If k = 0 Then
                    LogDates mFiles(i).FileName, "ORDER_DATE", mFiles(i).Conv(0), mFiles(i).NBad(0), mFiles(i).BadRows(0)
                    LogDates mFiles(i).FileName, "REQUESTED_DATE", mFiles(i).Conv(1), mFiles(i).NBad(1), mFiles(i).BadRows(1)
                Else
                    LogDates mFiles(i).FileName, "INVENTORY_DATE", mFiles(i).Conv(0), mFiles(i).NBad(0), mFiles(i).BadRows(0)
                End If
            End If
        Next i
    Next k
    kinds = Array("OI", "INV", "OO")
    labels = Array("Order Intake", "Inventory by Month", "Open Order")
    For k = 0 To 2
        n = 0
        names = ""
        For i = 1 To mNFiles
            If mFiles(i).Kind = kinds(k) Then
                n = n + 1
                names = names & IIf(n > 1, ", ", "") & mFiles(i).FileName
            End If
        Next i
        If n > 1 Then LogLine n & " " & labels(k) & " files found, all are read: " & names
    Next k
    For i = 1 To mNFiles
        If mFiles(i).BlankCity > 0 And mFiles(i).Kind <> "OO" Then LogLine mFiles(i).FileName & ": " & _
            CountText(mFiles(i).BlankCity, "row") & " with a blank MSA - shown as " & NO_CITY
    Next i
    For i = 1 To mNFiles
        If FileUsable(i, "INV") Then hasInv = True
    Next i
    If Not hasInv Then
        LogLine "Inventory: no Inventory by Month file - the summary shows zeros"
    ElseIf mMonth < 0 Then
        LogLine "Inventory: no valid INVENTORY_DATE values - the summary shows zeros"
    ElseIf Not inMonth Then
        LogLine "Inventory: no rows for " & MonthText(mMonth) & " - the summary shows zeros"
    End If
    If mOpenRows = 0 Then LogLine "Open orders: no rows with ORDER_STATUS ORDERED or Not yet processed - the Open Orders by MSA tab shows zeros"
    If nSpaced > 0 Then LogLine "WARNING: " & CountText(nSpaced, "row") & " have ORDER_STATUS ORDERED or Not yet processed with " & _
        "extra spaces and are not counted on the Open Orders by MSA tab: " & spacedText & IIf(nSpacedTexts > MAX_ROWS_LISTED, ", ...", "")
    If nDisagree > 0 Then LogLine "Open orders: " & CountText(nDisagree, "row") & " where IS_OPEN_ORDER disagrees with the status " & _
        "(TRUE with another status, or FALSE with ORDERED / Not yet processed)"
    If mNonIsoText Then
        If Application.International(xlDateOrder) <> 0 Then LogLine "WARNING: text dates not in yyyy-mm-dd form were read " & _
            "with this PC's date order, which is not month-first; check them against the export."
    End If
End Sub

Private Function AbbrIndex(ByVal abbr As String) As Long
    Select Case abbr
        Case "CTL": AbbrIndex = 0
        Case "MX": AbbrIndex = 1
        Case Else: AbbrIndex = 2
    End Select
End Function

' True when Order Intake row n is hidden by the default filters: Excel's "<>CANCELLED" ignores case but not
' stray spaces, so the status text is compared as written; SSL rows unless ShowSSL.
Private Function OIHidden(ByVal n As Long, opt As ReportOptions) As Boolean
    OIHidden = (Not opt.ShowCancelled And UCase$(SafeStr(mOI(n, 5))) = "CANCELLED") Or (Not opt.ShowSSL And CStr(mOI(n, 8)) = "SSL")
End Function

' "Options: InventoryMonth=..., ShowCancelled=..., ShowSSL=..., MSAColumn=..., SizeClassFile=<file name>"
Private Function OptionsLine(opt As ReportOptions) As String
    OptionsLine = "Options: InventoryMonth=" & IIf(opt.InvMonth < 0, "latest in file", MonthText(opt.InvMonth)) & _
                  ", ShowCancelled=" & IIf(opt.ShowCancelled, "TRUE", "FALSE") & ", ShowSSL=" & IIf(opt.ShowSSL, "TRUE", "FALSE") & _
                  ", MSAColumn=" & IIf(Len(opt.CityColumn) = 0, "auto", opt.CityColumn) & ", SizeClassFile=" & FileNameOf(opt.SizeClassFile)
End Function

' =====================================================================================================
'  OUTPUT SHEETS  (style sheet: plan Section 5.5)
' =====================================================================================================

Private Sub WriteFileFormat(ws As Worksheet, opt As ReportOptions, ByVal builtAt As Date)
    Dim arr As Variant, i As Long, r As Long, k As Long, n As Long, kinds As Variant, labels As Variant, u As Long
    SheetFont ws, FONT_REPORT
    StyleTitle ws, "File Format", "Layout and rules of this report.", FF_MONTH_COL

    r = FF_CLASS_ROW
    CaptionCell ws.Cells(r, 1), "Table 1.", "Size classes, in report order."
    PutRow ws, r + 1, Array("Class", "Family", "Rule")
    For i = 1 To mClsCount                                          ' v2.8: the allowed classes, in rank order
        ws.Cells(r + 1 + i, 1).Value = SafeCellText(mClsName(i))
        ws.Cells(r + 1 + i, 2).Value = FamilyName(mClsFam(i))
        ws.Cells(r + 1 + i, 3).Value = FamilyRule(mClsFam(i))
    Next i
    u = r + 2 + mClsCount                                           ' the UNMAPPED row
    PutRow ws, u, Array("UNMAPPED", "Any", "model not in the size class file; add it there and build again")
    ' the UNMAPPED row is left out of the banding, or on an even row the green would hide the peach (v2.7)
    StyleTable ws, r + 1, u, "LLL", 1, 1                            ' col 1, the last (UNMAPPED) row not banded
    For i = 2 To mClsCount                                          ' a medium rule above each family after the first
        If mClsFam(i) <> mClsFam(i - 1) Then SetRule ws.Range(ws.Cells(r + 1 + i, 1), ws.Cells(r + 1 + i, 3)), xlEdgeTop, "medium"
    Next i
    ws.Cells(u, 1).Font.Bold = True
    ws.Cells(u, 1).Interior.Color = CLR_FLAG

    r = u + 2
    CaptionCell ws.Cells(r, 1), "Table 2.", "Order status rank inside a size class."
    PutRow ws, r + 1, Array("Rank", "ORDER_STATUS")
    arr = StatusOrder()
    For i = 0 To 5
        ws.Cells(r + 2 + i, 1).Value = i + 1
        ws.Cells(r + 2 + i, 2).Value = arr(i)
    Next i
    StyleTable ws, r + 1, r + 7, "CL"

    r = r + FF_TABLE_GAP
    CaptionCell ws.Cells(r, 1), "Table 3.", "Default filters."
    PutRow ws, r + 1, Array("Tab", "Column", "Shown", "Config setting")
    PutRow ws, r + 2, Array("Order Intake", "ORDER_STATUS", IIf(opt.ShowCancelled, "all", "all except CANCELLED"), _
                            "ShowCancelled = " & IIf(opt.ShowCancelled, "TRUE", "FALSE"))
    PutRow ws, r + 3, Array("Order Intake", "PRODUCT_TYPE_ABBR", IIf(opt.ShowSSL, "all (CTL, MX, SSL)", "CTL, MX"), _
                            "ShowSSL = " & IIf(opt.ShowSSL, "TRUE", "FALSE"))
    PutRow ws, r + 4, Array("Inventory by Month", "Month", IIf(mMonth >= 0, MonthText(mMonth), "(no inventory month)"), _
                            "InventoryMonth = " & IIf(opt.InvMonth < 0, "latest in file", MonthText(opt.InvMonth)))
    PutRow ws, r + 5, Array("Inventory by Month", "PRODUCT_TYPE", IIf(opt.ShowSSL, "all (Compact Track Loaders, Mini Excavators, " & _
                            "Skid-Steer Loaders)", "Compact Track Loaders, Mini Excavators"), "ShowSSL = " & IIf(opt.ShowSSL, "TRUE", "FALSE"))
    PutRow ws, r + 6, Array("Open Orders by MSA", "PRODUCT_TYPE_ABBR", IIf(opt.ShowSSL, "CTL, MX, SSL on the chart", _
                            "CTL, MX on the chart"), "ShowSSL = " & IIf(opt.ShowSSL, "TRUE", "FALSE"))
    StyleTable ws, r + 1, r + 6, "LLLL"
    ws.Cells(r + 7, 1).Value = "Note. Clear a filter to see the rows it hides."
    StyleNote ws.Cells(r + 7, 1)

    r = r + FF_TABLE_GAP
    CaptionCell ws.Cells(r, 1), "Table 4.", "This build."
    PutRow ws, r + 1, Array("Item", "Value")
    n = r + 2
    PutPair ws, n, "Built", Format$(builtAt, "yyyy-mm-dd hh:mm")
    PutPair ws, n, "Builder", BUILDER_TEXT
    PutPair ws, n, "Input folder", opt.InputFolder
    kinds = Array("OI", "INV", "OO")
    labels = Array("Order Intake file", "Inventory by Month file", "Open Order file")
    For k = 0 To 2
        For i = 1 To mNFiles
            If mFiles(i).Kind = kinds(k) Then PutPair ws, n, CStr(labels(k)), FileSummary(i)
        Next i
    Next k
    PutPair ws, n, "Size class file", opt.SizeClassFile & " (" & CountText(mScCount, "usable row") & ")"
    PutPair ws, n, "Inventory month", IIf(mMonth >= 0, MonthText(mMonth), "(none)")
    PutPair ws, n, "MSAs", CityNames()
    PutPair ws, n, "Open orders", OpenOrdersText()
    PutPair ws, n, "Options", Mid$(OptionsLine(opt), Len("Options: ") + 1)
    StyleTable ws, r + 1, n - 1, "LLLL"                                        ' Value runs over C:D

    CaptionCell ws.Cells(FF_CLASS_ROW, FF_MONTH_COL), "Table 5.", "Months in this build."
    For i = 1 To mNMonths
        ws.Cells(FF_CLASS_ROW + i, FF_MONTH_COL).Value = SafeCellText(MonthText(mMonths(i)))
    Next i
    If mNMonths > 0 Then
        Grid ws.Range(ws.Cells(FF_CLASS_ROW + 1, FF_MONTH_COL), ws.Cells(FF_CLASS_ROW + mNMonths, FF_MONTH_COL))
        AlignCols ws.Range(ws.Cells(FF_CLASS_ROW + 1, FF_MONTH_COL), ws.Cells(FF_CLASS_ROW + mNMonths, FF_MONTH_COL)), "L"
        Band ws.Range(ws.Cells(FF_CLASS_ROW + 1, FF_MONTH_COL), ws.Cells(FF_CLASS_ROW + mNMonths, FF_MONTH_COL))
    End If
    SetWidths ws, Array(22, 26, 62, 32, 3, 20)
End Sub

Private Sub WriteOrderIntake(ws As Worksheet, opt As ReportOptions, ByVal builtAt As Date)
    Dim n As Long, i As Long, k As Long, block() As Variant, hiddenText As String, nFiles As Long, startRow As Long
    Dim rng As Range, hid As Boolean, wrong As Long, lastOI As Long
    n = mNOI
    If n < 1 Then lastOI = 3 Else lastOI = 2 + n
    For i = 1 To mNFiles
        If FileUsable(i, "OI") Then nFiles = nFiles + 1
    Next i
    If Not opt.ShowCancelled And Not opt.ShowSSL Then
        hiddenText = "Filtered out: cancelled orders, skid-steer loaders."
    ElseIf Not opt.ShowCancelled Then
        hiddenText = "Filtered out: cancelled orders."
    ElseIf Not opt.ShowSSL Then
        hiddenText = "Filtered out: skid-steer loaders."
    Else
        hiddenText = "No rows filtered out."
    End If
    SheetFont ws, FONT_REPORT
    StyleTitle ws, "Order Intake", "Source: Snowflake Order Intake export, " & CountText(nFiles, "file") & ", built " & _
               Format$(builtAt, "yyyy-mm-dd") & ". " & hiddenText, 11, OI_NOTE_COL
    With ws.Range("F1")                                  ' the visible total, in the frozen title row
        .Value = OI_TOTAL_LABEL
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    With ws.Range("G1")                                  ' SUBTOTAL 109 skips the rows a filter (or a person) hid
        .Formula = "=SUBTOTAL(109," & AbsRange("G", 3, lastOI) & ")"
        .NumberFormat = "#,##0"
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlLeft                    ' read on from the label, clear of the note
        .IndentLevel = 1
    End With
    PutRow ws, 2, Array("Size Class", "MSA", "ORDER_DATE", "REQUESTED_DATE", "ORDER_STATUS", "IS_OPEN_ORDER", _
                        "SUM(F.ORDERED_QUANTITY)", "PRODUCT_TYPE_ABBR", "MODEL_GROUP", "MODEL_NUMBER", "ORGANIZATION_NAME")
    SetWidths ws, Array(20.29, 18, 18.86, 22.57, 19.29, 21, 35.57, 22.29, 17.86, 24, 56.86)
    StyleTable ws, 2, 2 + n, OI_ALIGN
    If n > 0 Then AddFillRule ws.Range(ws.Cells(3, 1), ws.Cells(2 + n, 1)), "=$A3=""UNMAPPED""", CLR_FLAG, True
    If n > 0 Then
        ReDim block(1 To n, 1 To 11)
        For i = 1 To n
            For k = 1 To 11
                block(i, k) = mOI(mOIOrd(i), k)
            Next k
            block(i, 2) = mCityNames(mOICity(mOIOrd(i)))
        Next i
        ProtectText block, n, 1, 11                      ' v2.8: column A too, a class from the Set-up tab can look like a number
        ws.Range(ws.Cells(3, 1), ws.Cells(2 + n, 11)).Value = block
        ws.Range(ws.Cells(3, 3), ws.Cells(2 + n, 4)).NumberFormat = "m/d/yyyy"
        ws.Range(ws.Cells(3, 7), ws.Cells(2 + n, 7)).NumberFormat = "0"
        startRow = 1                                     ' a medium rule above each size class; UNMAPPED in bold
        For i = 2 To n + 1
            If i > n Then
                ClassGroup ws, startRow, i - 1, CStr(block(startRow, 1))
            ElseIf CStr(block(i, 1)) <> CStr(block(startRow, 1)) Then
                ClassGroup ws, startRow, i - 1, CStr(block(startRow, 1))
                startRow = i
            End If
        Next i
    End If
    Set rng = ws.Range(ws.Cells(2, 1), ws.Cells(2 + n, 11))
    SetAutoFilter ws, rng
    If n > 0 Then
        If Not opt.ShowCancelled Then rng.AutoFilter Field:=5, Criteria1:="<>CANCELLED"
        If Not opt.ShowSSL Then rng.AutoFilter Field:=8, Criteria1:=Array("CTL", "MX"), Operator:=xlFilterValues
        For i = 1 To n                                   ' the rows the filters hid must be the rows the rules hide
            hid = OIHidden(mOIOrd(i), opt)
            If ws.Rows(2 + i).Hidden <> hid Then
                ws.Rows(2 + i).Hidden = hid
                wrong = wrong + 1
            End If
        Next i
        If wrong > 0 Then LogLine "WARNING: Excel's filters on the Order Intake tab hid other rows than the rules; " & _
            CountText(wrong, "row") & " set directly. Please report this."
    End If
    HideGridlines ws
    FreezeBelow ws, 3
End Sub

' Rows first..last of the Order Intake block (1-based) form one size class: a medium rule above it (the first
' class sits under the header's double rule), the Size Class cells in bold when the class is UNMAPPED (their
' peach fill is a conditional format, so it stays above the banding).
Private Sub ClassGroup(ws As Worksheet, ByVal first As Long, ByVal last As Long, ByVal cls As String)
    If first > 1 Then SetRule ws.Range(ws.Cells(2 + first, 1), ws.Cells(2 + first, 11)), xlEdgeTop, "medium"
    If cls = "UNMAPPED" Then ws.Range(ws.Cells(2 + first, 1), ws.Cells(2 + last, 1)).Font.Bold = True
End Sub

Private Sub WriteInventory(ws As Worksheet, opt As ReportOptions, ByVal builtAt As Date)
    Dim nc As Long, nd As Long, h As Long, first As Long, last As Long, i As Long, j As Long, r As Long, n As Long
    Dim nFiles As Long, mText As String, block() As Variant, rng As Range, col As String, hid As Boolean, wrong As Long
    Dim detailText As String, e As Variant
    nc = mNCities
    nd = mNINV
    For i = 1 To mNFiles
        If FileUsable(i, "INV") Then nFiles = nFiles + 1
    Next i
    If mMonth >= 0 Then mText = MonthText(mMonth)
    If Len(mText) > 0 Then
        detailText = " Detail filtered to " & mText & IIf(opt.ShowSSL, "", ", CTL and MX") & "."
    Else
        detailText = " No inventory rows."
    End If
    SheetFont ws, FONT_REPORT
    StyleTitle ws, "Inventory by Month", "Source: Snowflake Inventory by Month export, " & CountText(nFiles, "file") & _
               ", built " & Format$(builtAt, "yyyy-mm-dd") & ". Units: sum of TOTAL for the month in B3." & detailText, 9
    ws.Range("A3").Value = "Units in"
    ws.Range("A3").Font.Bold = True
    ws.Range("B3").NumberFormat = "@"                   ' Text: a month picked in the drop-down stays text, not a date
    If Len(mText) > 0 Then ws.Range("B3").Value = mText
    For Each e In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)   ' the month cell: yellow, boxed
        With ws.Range("B3").Borders(e)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = 0
        End With
    Next e
    ws.Range("B3").Interior.Color = CLR_INPUT
    AlignCols ws.Range("B3"), "L"
    If mNMonths > 0 Then
        With ws.Range("B3").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                 Formula1:="='File Format'!$F$" & (FF_CLASS_ROW + 1) & ":$F$" & (FF_CLASS_ROW + mNMonths)
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    End If
    PutRow ws, 4, Array("MSA", "CTL", "MX", "SSL", "Total")

    h = 7 + nc
    first = h + 1
    last = h + IIf(nd < 1, 1, nd)
    For i = 1 To nc
        r = 4 + i
        ws.Cells(r, 1).Value = SafeCellText(mCityNames(SortedCity(i)))
        For j = 0 To 2
            ws.Cells(r, 2 + j).Formula = "=SUMIFS(" & AbsRange("G", first, last) & "," & AbsRange("A", first, last) & ",$A" & r & "," & _
                AbsRange("B", first, last) & ",""" & ProductTypeName(j) & """," & AbsRange("F", first, last) & ",$B$3)"
        Next j
        ws.Cells(r, 5).Formula = "=SUM(B" & r & ":D" & r & ")"
    Next i
    If nc > 0 Then ws.Range(ws.Cells(5, 2), ws.Cells(4 + nc, 5)).NumberFormat = "#,##0"
    r = 5 + nc
    ws.Cells(r, 1).Value = "Total"
    For j = 2 To 5
        col = Mid$("ABCDE", j, 1)
        If nc > 0 Then ws.Cells(r, j).Formula = "=SUM(" & col & "5:" & col & (4 + nc) & ")" Else ws.Cells(r, j).Value = 0
    Next j
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).NumberFormat = "#,##0"
    StyleTable ws, 4, r, SUM_ALIGN, 1, 1                        ' the Total row is not banded
    SetRule ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)), xlEdgeTop, "double"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Font.Bold = True

    PutRow ws, h, Array("MSA", "PRODUCT_TYPE", "MODEL_GROUP", "MODEL_NUMBER", "INVENTORY_DATE", "Month", "TOTAL", _
                        "ORGANIZATION_NAME", "ORGANIZATION_CUSTOMER_NUMBER")
    SetWidths ws, Array(18, 25.14, 14, 16, 17, 16, 11.29, 56.86, 36)
    StyleTable ws, h, h + nd, INV_ALIGN
    If nd > 0 Then
        ReDim block(1 To nd, 1 To 9)
        For i = 1 To nd
            n = mINVOrd(i)
            block(i, 1) = mCityNames(mINVCity(n))
            For j = 1 To 4
                block(i, j + 1) = mINV(n, j)
            Next j
            If mINVMonth(n) >= 0 Then block(i, 6) = MonthText(mINVMonth(n))
            block(i, 7) = mINV(n, 5)
            block(i, 8) = mINV(n, 6)
            block(i, 9) = mINV(n, 7)
        Next i
        ProtectText block, nd, 1, 9
        ws.Range(ws.Cells(first, 1), ws.Cells(h + nd, 9)).Value = block
        ws.Range(ws.Cells(first, 5), ws.Cells(h + nd, 5)).NumberFormat = "m/d/yyyy"
        ws.Range(ws.Cells(first, 7), ws.Cells(h + nd, 7)).NumberFormat = "#,##0"
        For i = 2 To nd                                  ' a medium rule above each MSA
            If block(i, 1) <> block(i - 1, 1) Then SetRule ws.Range(ws.Cells(h + i, 1), ws.Cells(h + i, 9)), xlEdgeTop, "medium"
        Next i
    End If
    Set rng = ws.Range(ws.Cells(h, 1), ws.Cells(h + nd, 9))
    SetAutoFilter ws, rng
    If nd > 0 Then
        If mMonth >= 0 Then rng.AutoFilter Field:=6, Criteria1:=Array(mText), Operator:=xlFilterValues
        If Not opt.ShowSSL Then rng.AutoFilter Field:=2, Criteria1:=Array(ProductTypeName(0), ProductTypeName(1)), Operator:=xlFilterValues
        For i = 1 To nd                                  ' the rows the filters hid must be the rows the rules hide
            n = mINVOrd(i)
            hid = (mMonth >= 0 And mINVMonth(n) <> mMonth) Or (Not opt.ShowSSL And mINVType(n) = 2)
            If ws.Rows(h + i).Hidden <> hid Then
                ws.Rows(h + i).Hidden = hid
                wrong = wrong + 1
            End If
        Next i
        If wrong > 0 Then LogLine "WARNING: Excel's filters on the Inventory by Month tab hid other rows than the rules; " & _
            CountText(wrong, "row") & " set directly. Please report this."
    End If
    HideGridlines ws
    FreezeBelow ws, first
End Sub

' The Open Orders by MSA tab (plan open-orders-by-msa-tab-implementation-plan.md Section 5.2): the ordered quantity of
' the Order Intake rows whose ORDER_STATUS is ORDERED or Not yet processed, by MSA (every MSA of the report,
' alphabetical, as in the inventory summary) and product type, as live SUMIFS over the Order Intake tab, so the table
' always agrees with the rows the report holds; a Total column and row; then the chart of the table. SUMIFS compares
' the whole text, ignoring case but not stray spaces; BuildLogLines warns about statuses with stray spaces.
Private Sub WriteOpenOrders(ws As Worksheet, opt As ReportOptions, ByVal builtAt As Date)
    Dim nc As Long, nt As Long, i As Long, j As Long, r As Long, lastOI As Long, anchorRow As Long
    Dim col As String, f As String, crit As Variant, rG As String, rB As String, rH As String, rE As String
    nc = mNCities
    If opt.ShowSSL Then nt = 3 Else nt = 2               ' product types on the chart (the table always has SSL)
    If mNOI < 1 Then lastOI = 3 Else lastOI = 2 + mNOI
    rG = "'Order Intake'!" & AbsRange("G", 3, lastOI)    ' SUM(F.ORDERED_QUANTITY)
    rB = "'Order Intake'!" & AbsRange("B", 3, lastOI)    ' MSA
    rH = "'Order Intake'!" & AbsRange("H", 3, lastOI)    ' PRODUCT_TYPE_ABBR
    rE = "'Order Intake'!" & AbsRange("E", 3, lastOI)    ' ORDER_STATUS
    SheetFont ws, FONT_REPORT
    StyleTitle ws, "Open Orders by MSA", "Source: the Order Intake tab, built " & Format$(builtAt, "yyyy-mm-dd") & _
               ". Ordered quantity of orders with ORDER_STATUS ORDERED or Not yet processed, by MSA and product type." & _
               IIf(opt.ShowSSL, " Chart: CTL, MX and SSL.", " Chart: CTL and MX."), 5
    CaptionCell ws.Range("A3"), "Table 1.", "Open orders by MSA: sum of SUM(F.ORDERED_QUANTITY) over the Order Intake " & _
                "rows whose ORDER_STATUS is ORDERED or Not yet processed."
    PutRow ws, 4, Array("MSA", "CTL", "MX", "SSL", "Total")
    For i = 1 To nc
        r = 4 + i
        ws.Cells(r, 1).Value = SafeCellText(mCityNames(SortedCity(i)))
        For j = 2 To 4
            col = Mid$("ABCDE", j, 1)
            f = ""
            For Each crit In Array("ORDERED", "Not yet processed")
                f = f & IIf(Len(f) = 0, "=", "+") & "SUMIFS(" & rG & "," & rB & ",$A" & r & "," & rH & "," & col & "$4," & _
                    rE & ",""" & crit & """)"
            Next crit
            ws.Cells(r, j).Formula = f
        Next j
        ws.Cells(r, 5).Formula = "=SUM(B" & r & ":D" & r & ")"
    Next i
    If nc > 0 Then ws.Range(ws.Cells(5, 2), ws.Cells(4 + nc, 5)).NumberFormat = "#,##0"
    r = 5 + nc
    ws.Cells(r, 1).Value = "Total"
    For j = 2 To 5
        col = Mid$("ABCDE", j, 1)
        If nc > 0 Then ws.Cells(r, j).Formula = "=SUM(" & col & "5:" & col & (4 + nc) & ")" Else ws.Cells(r, j).Value = 0
    Next j
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).NumberFormat = "#,##0"
    StyleTable ws, 4, r, SUM_ALIGN, 1, 1                        ' the Total row is not banded
    SetRule ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)), xlEdgeTop, "double"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Font.Bold = True
    SetWidths ws, Array(36, 12, 12, 12, 12)
    anchorRow = 7 + nc                                         ' two empty rows under the Total row
    If nc > 0 Then
        DrawOpenOrdersChart ws, nc, nt, anchorRow
    Else
        ws.Cells(anchorRow, 1).Value = "No MSAs: no chart."      ' a chart of the header row alone would fail
        StyleNote ws.Cells(anchorRow, 1)
    End If
    HideGridlines ws
End Sub

' The clustered column chart of the Open Orders by MSA table: source A4 to the last product type column of the last
' MSA row, plotted by rows, so one series per MSA (named from column A) in each product type group (CTL, MX and, with
' ShowSSL, SSL from row 4); the Total column and row stay out. Anchored at A<anchorRow>, 700 x 320 pt. Excel 2010
' members only, and no Select, ActiveChart or ActiveSheet. The only chart code in the module
' (run_macro_libreoffice.py replaces its one call; build_report_reference.py draws the same chart).
Private Sub DrawOpenOrdersChart(ws As Worksheet, ByVal nc As Long, ByVal nt As Long, ByVal anchorRow As Long)
    Dim co As ChartObject, ch As Chart, pal As Variant, i As Long
    pal = ChartPalette()
    Set co = ws.ChartObjects.Add(ws.Cells(anchorRow, 1).Left, ws.Cells(anchorRow, 1).Top, 700, 320)
    co.Name = "OpenOrdersChart"
    Set ch = co.Chart
    ch.ChartType = xlColumnClustered
    ch.SetSourceData Source:=ws.Range(ws.Cells(4, 1), ws.Cells(4 + nc, 1 + nt)), PlotBy:=xlRows
    Do While ch.SeriesCollection.Count > nc               ' had Excel read the range another way, the series are
        ch.SeriesCollection(ch.SeriesCollection.Count).Delete     ' set to the exact rows below
    Loop
    Do While ch.SeriesCollection.Count < nc
        ch.SeriesCollection.NewSeries
    Loop
    For i = 1 To nc
        With ch.SeriesCollection(i)
            .Name = "='" & ws.Name & "'!$A$" & (4 + i)
            .Values = ws.Range(ws.Cells(4 + i, 2), ws.Cells(4 + i, 1 + nt))
            .XValues = ws.Range(ws.Cells(4, 2), ws.Cells(4, 1 + nt))
            .Format.Fill.Visible = True
            .Format.Fill.Solid
            .Format.Fill.ForeColor.RGB = pal((i - 1) Mod 8)
            .HasDataLabels = True
            .DataLabels.ShowValue = True
            .DataLabels.NumberFormat = "#,##0"
        End With
    Next i
    ch.ChartGroups(1).GapWidth = 80
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    With ch.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "Ordered quantity"
        .TickLabels.NumberFormat = "#,##0"
        .HasMajorGridlines = True
        .HasMinorGridlines = False
    End With
    ch.Axes(xlCategory).HasMajorGridlines = False
    ch.ChartArea.Format.Fill.Visible = False                 ' no fill: white on the page (True / False, not mso constants)
    ch.PlotArea.Format.Fill.Visible = False
    With ch.ChartArea.Format.Line                            ' a thin grey border
        .Visible = True
        .ForeColor.RGB = RGB(217, 217, 217)
        .Weight = 0.75
    End With
    ch.HasTitle = True
    ch.ChartTitle.Text = "Open orders by product type and MSA (ordered quantity)"
    With ch.ChartArea.Font                                   ' every chart text in Cambria 10, then the title 12 bold
        .Name = FONT_REPORT
        .Size = BODY_SIZE
        .Bold = False                                        ' the default style bolds the axis title
    End With
    With ch.ChartTitle.Font
        .Name = FONT_REPORT
        .Size = 12
        .Bold = True
    End With
End Sub

' The series fills of the Open Orders chart, in MSA order (wrapping after eight); the first is the report green.
Private Function ChartPalette() As Variant
    ChartPalette = Array(HEAD_FILL, RGB(127, 127, 127), RGB(47, 85, 151), RGB(197, 90, 17), RGB(112, 48, 160), _
                         RGB(112, 173, 71), RGB(165, 165, 165), RGB(38, 68, 120))
End Function

' The File Format build table's "Open orders" item: the open rows and their units over all MSAs, as BuildLogLines
' counted them (all three product types, whatever ShowSSL says).
Private Function OpenOrdersText() As String
    OpenOrdersText = "ORDERED and Not yet processed: " & CountText(mOpenRows, "row") & ", " & _
                     CStr(mOpenUnits(0) + mOpenUnits(1) + mOpenUnits(2)) & " units (CTL " & CStr(mOpenUnits(0)) & ", MX " & _
                     CStr(mOpenUnits(1)) & ", SSL " & CStr(mOpenUnits(2)) & "), on the Open Orders by MSA tab"
End Function

Private Sub WriteRunLog(ws As Worksheet, opt As ReportOptions, ByVal builtAt As Date, ByVal secs As Single)
    Dim i As Long
    SheetFont ws, FONT_LOG
    ws.Range("A1").Value = SafeCellText("Built " & Format$(builtAt, "yyyy-mm-dd hh:mm") & " from " & opt.InputFolder & _
                                        " in " & Format$(secs, "0.0") & " s")
    ws.Range("A2").Value = SafeCellText(OptionsLine(opt))
    For i = 1 To mLog.Count
        ws.Cells(3 + i, 1).Value = SafeCellText(CStr(mLog.Item(i)))
    Next i
    ws.Columns("A").ColumnWidth = 150
End Sub

' Page setup for printing: landscape, one page wide (pagesTall 0 = as many pages as needed), the report name
' and the build date in the page header, the tab name and "Page n of m" in the footer, titleRows repeated on
' every page. A printer driver that refuses a setting leaves a WARNING in the Run Log; the report is complete.
Private Sub SetUpPage(ws As Worksheet, ByVal titleRows As String, ByVal pagesTall As Long, ByVal builtAt As Date)
    On Error GoTo Fail
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        If pagesTall > 0 Then
            .FitToPagesTall = pagesTall
        Else
            .FitToPagesTall = False                   ' as many pages tall as needed
        End If
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .LeftHeader = HF_FONT & REPORT_TITLE
        .RightHeader = HF_FONT & "Built " & Format$(builtAt, "yyyy-mm-dd")
        .LeftFooter = HF_FONT & "&A"
        .CenterFooter = HF_FONT & "Page &P of &N"
        If Len(titleRows) > 0 Then .PrintTitleRows = titleRows
    End With
    Exit Sub
Fail:
    LogLine "WARNING: the print setup of the " & ws.Name & " tab stopped (" & Err.Description & "); the report itself is complete."
End Sub

' Application.PrintCommunication (Excel 2010+) off while the pages are set up, which makes it much faster.
Private Sub SetPrintCommunication(ByVal turnOn As Boolean)
    On Error Resume Next
    Application.PrintCommunication = turnOn
End Sub

' ----- style helpers -----

' Every cell of the sheet in one font at the body size; titles, headers and notes change only bold, italic
' or size.
Private Sub SheetFont(ws As Worksheet, ByVal fontName As String)
    ws.Cells.Font.Name = fontName
    ws.Cells.Font.Size = BODY_SIZE
End Sub

' Title in A1 (14 bold), the note in D1 (italic; Order Intake: H1), a medium rule under row 1 across the sheet's
' ncols columns.
Private Sub StyleTitle(ws As Worksheet, ByVal titleText As String, ByVal note As String, ByVal ncols As Long, _
                       Optional ByVal noteCol As Long = 4)
    With ws.Range("A1")
        .Value = titleText
        .Font.Size = 14
        .Font.Bold = True
    End With
    ws.Cells(1, noteCol).Value = SafeCellText(note)
    StyleNote ws.Cells(1, noteCol)
    SetRule ws.Range(ws.Cells(1, 1), ws.Cells(1, ncols)), xlEdgeBottom, "medium"
End Sub

Private Sub StyleNote(c As Range)
    c.Font.Italic = True
End Sub

' A table caption: "Table 1." in bold, then the description.
Private Sub CaptionCell(c As Range, ByVal label As String, ByVal s As String)
    c.Value = label & " " & s
    c.Font.Bold = False
    c.Characters(1, Len(label)).Font.Bold = True
End Sub

' A header row at head and its body to last, columns from col, one per letter of spec: the header bold white on
' HEAD_FILL, HEAD_HEIGHT high, a medium rule above it and a double rule under it; thin rules between the rows,
' no vertical lines; the columns aligned and indented by spec; body rows ROW_HEIGHT high and banded, all but the
' last unbanded rows.
Private Sub StyleTable(ws As Worksheet, ByVal head As Long, ByVal last As Long, ByVal spec As String, _
                       Optional ByVal col As Long = 1, Optional ByVal unbanded As Long = 0)
    Dim ncols As Long, hdr As Range, body As Range
    ncols = Len(spec)
    If last < head Then last = head
    Set hdr = ws.Range(ws.Cells(head, col), ws.Cells(head, col + ncols - 1))
    Grid ws.Range(ws.Cells(head, col), ws.Cells(last, col + ncols - 1))
    hdr.Font.Bold = True
    hdr.Font.Color = HEAD_TEXT
    hdr.Interior.Color = HEAD_FILL
    hdr.RowHeight = HEAD_HEIGHT
    AlignCols hdr, spec
    SetRule hdr, xlEdgeTop, "medium"
    SetRule hdr, xlEdgeBottom, "double"
    If last > head Then
        Set body = ws.Range(ws.Cells(head + 1, col), ws.Cells(last, col + ncols - 1))
        AlignCols body, spec
        body.RowHeight = ROW_HEIGHT
        If last - unbanded > head Then Band ws.Range(ws.Cells(head + 1, col), ws.Cells(last - unbanded, col + ncols - 1))
    End If
End Sub

' Every second visible row of rng light green: conditional formatting that counts the visible rows from the
' first row of rng down to this one (SUBTOTAL 103 skips hidden rows), so the bands follow any filter.
Private Sub Band(rng As Range)
    Dim c As String
    c = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", rng.Column, 1)  ' report tables start in A to F
    AddFillRule rng, "=MOD(SUBTOTAL(103,$" & c & "$" & rng.Row & ":$" & c & rng.Row & "),2)=0", BAND_FILL, False
End Sub

' A conditional format on rng: formula f, written as for the first cell of rng, filling fillColor. first: the
' rule goes first and stops the rules after it, so its fill wins over the banding.
Private Sub AddFillRule(rng As Range, ByVal f As String, ByVal fillColor As Long, ByVal first As Boolean)
    f = Replace(f, ",", ListSeparator())
    rng.Worksheet.Activate                               ' Excel reads the formula relative to the active cell
    rng.Cells(1, 1).Select
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=f)
        .Interior.Color = fillColor
        .StopIfTrue = first
        If first Then .SetFirstPriority
    End With
End Sub

' The PC's list separator ("," or ";"): conditional formatting formulas are read in the PC's own notation.
Private Function ListSeparator() As String
    Dim sep As String
    On Error Resume Next
    sep = CStr(Application.International(xlListSeparator))
    On Error GoTo 0
    If Len(sep) = 1 Then ListSeparator = sep Else ListSeparator = ","   ' anything odd: the usual comma
End Function

' L left, R right (both indented one step), C centre; every cell centred vertically.
Private Sub AlignCols(rng As Range, ByVal spec As String)
    Dim i As Long
    For i = 1 To Len(spec)
        With rng.Columns(i)
            .VerticalAlignment = xlCenter
            Select Case Mid$(spec, i, 1)
                Case "R"
                    .HorizontalAlignment = xlRight
                    .IndentLevel = 1
                Case "C"
                    .HorizontalAlignment = xlCenter
                Case Else
                    .HorizontalAlignment = xlLeft
                    .IndentLevel = 1
            End Select
        End With
    Next i
End Sub

' Thin black rules above, between and below the rows of rng; no vertical lines.
Private Sub Grid(rng As Range)
    Dim e As Variant
    For Each e In Array(xlEdgeTop, xlEdgeBottom, xlInsideHorizontal)
        If Not (e = xlInsideHorizontal And rng.Rows.Count = 1) Then
            With rng.Borders(e)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = 0
            End With
        End If
    Next e
End Sub

' One edge of rng as a "thin", "medium" or "double" black rule.
Private Sub SetRule(rng As Range, ByVal edge As Long, ByVal style As String)
    With rng.Borders(edge)
        If style = "double" Then
            .LineStyle = xlDouble
        Else
            .LineStyle = xlContinuous
            .Weight = IIf(style = "medium", xlMedium, xlThin)
        End If
        .Color = 0
    End With
End Sub

' Text values in a row from column A, kept as text (a value that looks like a date or number stays as written).
Private Sub PutRow(ws As Worksheet, ByVal r As Long, ByVal values As Variant)
    Dim i As Long
    For i = LBound(values) To UBound(values)
        ws.Cells(r, 1 + i - LBound(values)).Value = SafeCellText(CStr(values(i)))
    Next i
End Sub

' An item and its value in row r (columns A and B, kept as text); r moves to the next row.
Private Sub PutPair(ws As Worksheet, ByRef r As Long, ByVal item As String, ByVal value As String)
    ws.Cells(r, 1).Value = SafeCellText(item)
    ws.Cells(r, 2).Value = SafeCellText(value)
    r = r + 1
End Sub

' Column widths as Excel's UI shows them (the file stores about 0.71 more).
Private Sub SetWidths(ws As Worksheet, ByVal widths As Variant)
    Dim i As Long
    For i = 0 To UBound(widths)
        ws.Columns(i + 1).ColumnWidth = widths(i)
    Next i
End Sub

Private Function AbsRange(ByVal col As String, ByVal r1 As Long, ByVal r2 As Long) As String
    AbsRange = "$" & col & "$" & r1 & ":$" & col & "$" & r2
End Function

Private Sub HideGridlines(ws As Worksheet)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

' Freezes the rows above firstRow (the title and headers stay in view while scrolling).
Private Sub FreezeBelow(ws As Worksheet, ByVal firstRow As Long)
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .ScrollRow = 1
        .ScrollColumn = 1
        .SplitColumn = 0
        .SplitRow = firstRow - 1
        .FreezePanes = True
    End With
    ws.Range("A1").Select
End Sub

' =====================================================================================================
'  DATES STORED AS TEXT
' =====================================================================================================

' A cell from a date column: 0 = a date or empty, 1 = text that reads as a date (v is replaced by the date),
' 2 = anything else (v is left as it is).
Private Function DateStatus(ByRef v As Variant) As Long
    Dim d As Date
    If IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate
            DateStatus = 0
        Case vbString
            If Len(Trim$(v)) = 0 Then Exit Function
            If HasDay(CStr(v)) And TextToDate(CStr(v), d) Then
                v = d
                DateStatus = 1
            Else
                DateStatus = 2
            End If
        Case Else
            DateStatus = 2
    End Select
End Function

' A text date in an export must name a day, not only a month: three runs of digits (2026-08-05, 8/5/2026), or
' two and a month name (Aug 5, 2026). CDate alone would read "August 2026" as the 1st (the reference does not).
Private Function HasDay(ByVal s As String) As Boolean
    Dim q As Long, runs As Long, inRun As Boolean
    For q = 1 To Len(s)
        If Mid$(s, q, 1) Like "#" Then
            If Not inRun Then runs = runs + 1
            inRun = True
        Else
            inRun = False
        End If
    Next q
    HasDay = (runs >= 3) Or (runs = 2 And s Like "*[A-Za-z]*")
End Function

' Text that reads as a date -> d. Accepts what CDate accepts (the PC's date formats and ISO yyyy-mm-dd) plus
' a "T" between date and time, a trailing "Z" and fractional seconds, as in Snowflake's "2026-08-05 00:00:00.000".
' The text must contain a 4-digit year and give a date, not only a time. Text not in yyyy-mm-dd form is read
' with the PC's date order; on a PC that is not month-first the Run Log gets a warning (mNonIsoText).
Private Function TextToDate(ByVal s As String, ByRef d As Date) As Boolean
    Dim p As Long
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    If Right$(s, 1) = "Z" Or Right$(s, 1) = "z" Then s = Left$(s, Len(s) - 1)
    If Len(s) >= 11 Then
        If Mid$(s, 11, 1) = "T" Or Mid$(s, 11, 1) = "t" Then s = Left$(s, 10) & " " & Mid$(s, 12)
    End If
    p = InStrRev(s, ".")
    If p > 0 And InStr(s, ":") > 0 And p > InStrRev(s, ":") Then
        If Len(Mid$(s, p + 1)) > 0 And Not (Mid$(s, p + 1) Like "*[!0-9]*") Then s = Left$(s, p - 1)
    End If
    If Not (s Like "*####*") Then Exit Function
    If IsDate(s) Then
        d = CDate(s)
        If Int(CDbl(d)) = 0 Then Exit Function
        If Not (s Like "####-##-##*") Then mNonIsoText = True
        TextToDate = True
    End If
End Function

Private Sub AddRowNumber(ByRef list As String, ByVal count As Long, ByVal rowNum As Long)
    If count <= MAX_ROWS_LISTED Then list = list & IIf(Len(list) > 0, ", ", "") & rowNum
End Sub

' One Run Log line per date column that held text: how many were converted, which rows were not.
Private Sub LogDates(ByVal label As String, ByVal colName As String, ByVal nConv As Long, ByVal nBad As Long, ByVal badRows As String)
    Dim s As String
    If nConv = 0 And nBad = 0 Then Exit Sub
    s = label & " " & colName & ": " & nConv & " text value(s) converted to dates, " & nBad & " not a date and left as is"
    If nBad > 0 Then s = s & " (" & IIf(nBad = 1, "row ", "rows ") & badRows & IIf(nBad > MAX_ROWS_LISTED, ", ...", "") & ")"
    LogLine s
End Sub

Private Function MonthKey(ByVal d As Date) As Long
    MonthKey = Year(d) * 12 + Month(d) - 1
End Function

' "August 2026" for a month key (English month names whatever the PC's language); "" for -1 (no month). The
' guard matters: IIf evaluates both of its branches, so MonthText(-1) is called whenever no month is set.
Private Function MonthText(ByVal key As Long) As String
    If key < 0 Then Exit Function
    MonthText = EnglishMonth((key Mod 12) + 1) & " " & (key \ 12)
End Function

Private Function EnglishMonth(ByVal m As Long) As String
    EnglishMonth = Choose(m, "January", "February", "March", "April", "May", "June", "July", "August", _
                          "September", "October", "November", "December")
End Function

' =====================================================================================================
'  SMALL HELPERS
' =====================================================================================================

Private Sub LogLine(ByVal msg As String)
    mLog.Add msg
End Sub

Private Sub ShowStatus(ByVal msg As String)
    Application.StatusBar = STATUS_PREFIX & msg
End Sub

' CStr that survives #N/A, Null and Empty cells.
Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    SafeStr = CStr(v)
End Function

Private Function IsNumber(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte: IsNumber = True
    End Select
End Function

' City text: trimmed, runs of spaces collapsed.
Private Function NormCity(ByVal s As String) As String
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormCity = s
End Function

' The file name of a path (after the last \ or /).
Private Function FileNameOf(ByVal p As String) As String
    Dim k As Long
    k = InStrRev(p, "\")
    If InStrRev(p, "/") > k Then k = InStrRev(p, "/")
    FileNameOf = Mid$(p, k + 1)
End Function

' Text that Excel would turn into a number, a date, a Boolean or a formula when written ("00123", "1-2",
' "TRUE", "1 1/2", "5%", "=x", or text starting with an apostrophe) stays text: the quote prefix keeps it.
Private Function SafeCellText(ByVal s As String) As String
    SafeCellText = s
    If Len(s) = 0 Then Exit Function
    If IsNumeric(s) Or IsDate(s) Or InStr("=+-@'", Left$(s, 1)) > 0 Or s Like "*#/#*" Or Right$(s, 1) = "%" _
       Or StrComp(s, "TRUE", vbTextCompare) = 0 Or StrComp(s, "FALSE", vbTextCompare) = 0 Then SafeCellText = "'" & s
End Function

Private Sub ProtectText(arr() As Variant, ByVal nRows As Long, ByVal firstCol As Long, ByVal lastCol As Long)
    Dim r As Long, c As Long
    For r = 1 To nRows
        For c = firstCol To lastCol
            If VarType(arr(r, c)) = vbString Then arr(r, c) = SafeCellText(CStr(arr(r, c)))
        Next c
    Next r
End Sub

' Stable merge sort: ord(1..n) lists 1..n in ascending order of sortKey; equal keys keep their order.
Private Sub StableSortIndex(sortKey() As Double, ByVal n As Long, ord() As Long)
    Dim tmp() As Long, runLen As Long, lo As Long, md As Long, hi As Long, i As Long, j As Long, k As Long
    ReDim ord(1 To n)
    ReDim tmp(1 To n)
    For i = 1 To n
        ord(i) = i
    Next i
    runLen = 1
    Do While runLen < n
        lo = 1
        Do While lo <= n
            md = lo + runLen - 1
            hi = lo + 2 * runLen - 1
            If md > n Then md = n
            If hi > n Then hi = n
            i = lo: j = md + 1: k = lo
            Do While i <= md And j <= hi
                If sortKey(ord(j)) < sortKey(ord(i)) Then
                    tmp(k) = ord(j)
                    j = j + 1
                Else
                    tmp(k) = ord(i)
                    i = i + 1
                End If
                k = k + 1
            Loop
            Do While i <= md
                tmp(k) = ord(i)
                i = i + 1
                k = k + 1
            Loop
            Do While j <= hi
                tmp(k) = ord(j)
                j = j + 1
                k = k + 1
            Loop
            lo = lo + 2 * runLen
        Loop
        For i = 1 To n
            ord(i) = tmp(i)
        Next i
        runLen = runLen * 2
    Loop
End Sub

Private Sub SetAutoFilter(ws As Worksheet, rng As Range)
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    rng.AutoFilter
End Sub

Private Function SheetExists(wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Function StripTrailingSeparator(ByVal p As String) As String
    p = Trim$(p)
    Do While Len(p) > 1 And (Right$(p, 1) = "\" Or Right$(p, 1) = "/")
        p = Left$(p, Len(p) - 1)
    Loop
    StripTrailingSeparator = p
End Function

Private Function PickFolder(ByVal title As String, ByVal startFolder As String) As String
    Dim fd As Object
    On Error Resume Next
    Set fd = Application.FileDialog(4)                 ' msoFileDialogFolderPicker
    If Err.Number = 0 And Not fd Is Nothing Then
        fd.Title = title
        If Len(startFolder) > 0 Then
            If FolderExists(StripTrailingSeparator(startFolder)) Then fd.InitialFileName = StripTrailingSeparator(startFolder) & Application.PathSeparator
        End If
        fd.AllowMultiSelect = False
        If fd.Show = -1 Then PickFolder = fd.SelectedItems(1)
        Exit Function
    End If
    On Error GoTo 0
    PickFolder = Trim$(InputBox(title, "Open Orders report"))
End Function

' A file picker for Excel workbooks; an InputBox where Excel has no file dialog.
Private Function PickFile(ByVal title As String, ByVal startPath As String) As String
    Dim fd As Object
    On Error Resume Next
    Set fd = Application.FileDialog(3)                 ' msoFileDialogFilePicker
    If Err.Number = 0 And Not fd Is Nothing Then
        fd.Title = title
        fd.AllowMultiSelect = False
        fd.Filters.Clear
        fd.Filters.Add "Excel workbooks", "*.xlsx; *.xlsm; *.xls"
        If Len(startPath) > 0 Then fd.InitialFileName = startPath
        If fd.Show = -1 Then PickFile = fd.SelectedItems(1)
        Exit Function
    End If
    On Error GoTo 0
    PickFile = Trim$(InputBox(title & vbLf & "(the full path of the workbook)", "Open Orders report"))
End Function

' The path to save the report to: p, or a time-stamped copy next to it when p exists and the user says No.
' Excel cannot save under the name of a workbook that is open (last time's report often still is, from
' any folder or a OneDrive address): such a workbook is closed if it has no unsaved changes, else the run stops.
Private Function ConfirmOutputPath(ByVal p As String, wbOut As Workbook) As String
    Dim dotPos As Long, wbx As Workbook, outName As String
    If Len(Dir(p)) > 0 Then
        If MsgBox(p & vbCrLf & vbCrLf & "already exists. Overwrite it?" & vbCrLf & "(No = save a time-stamped copy next to it)", _
                  vbYesNo + vbQuestion, "Open Orders report") <> vbYes Then
            dotPos = InStrRev(p, ".")
            If dotPos = 0 Then dotPos = Len(p) + 1
            p = Left$(p, dotPos - 1) & " " & Format$(Now, "yyyy-mm-dd hhmm") & Mid$(p, dotPos)
        End If
    End If
    outName = FileNameOf(p)
    For Each wbx In Application.Workbooks
        If Not wbx Is wbOut And StrComp(wbx.Name, outName, vbTextCompare) = 0 Then
            If Not wbx.Saved Then Err.Raise vbObjectError + 5, , outName & " is open in Excel with unsaved changes. Save or close it, then run again."
            wbx.Close SaveChanges:=False
            Exit For
        End If
    Next wbx
    ConfirmOutputPath = p
End Function

Private Function FolderExists(ByVal p As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(p & Application.PathSeparator) And vbDirectory) = vbDirectory)
End Function

Private Function FileExists(ByVal p As String) As Boolean
    On Error Resume Next
    FileExists = ((GetAttr(p) And vbDirectory) = 0)
    If Err.Number <> 0 Then FileExists = False
End Function
