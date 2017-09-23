VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_formApplicationLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'
' Project : Import Vehicle Commodity Tax Rebate
' Author : Al Fan
' Version : 0.1.0
' Coding date : 2017/08/22
'
'
' Define constant strings in registry
'
Const ApplicationName                       As String = "Vehicle Tax Rebate Access DB"
Const RegistrySectionName                   As String = "Add New Record"
Const RegistryChecklistFolder               As String = "Checklist Folder"
'
' Constants regarding to range definition in 檢核表
'
Const Vehicle_Body_ID_Range                 As String = "C8"
Const Dealer_Range                          As String = "D8"
Const Undertaker_Range                      As String = "E8"
Const Brand_Range                           As String = "F8"
Const Submit_Date_Range                     As String = "G8"
Const Tax_Rebate_Payee_Range                As String = "H8"
Const Cause_to_Refund_Range                 As String = "J8"
Const Import_Date_Range                     As String = "K8"
Const Arrival_Port_Range                    As String = "L8"
Const Number_of_Record_Range                As String = "CA1"
'
' Range of new vehicle information
'
Const New_Vehicle_Brand_Range               As String = "F8"
Const New_Vehicle_Factory_Date_Range        As String = "E29"   ' 新車出廠年月
Const New_Vehicle_Owner_Name_Range          As String = "F29"   ' 新車車主
Const New_Vehicle_Owner_ID_Range            As String = "G29"
Const New_Vehicle_Type_Range                As String = "H29"   ' 新車車別
Const New_Vehicle_Model_Range               As String = "I29"   ' 新車車型
Const New_Vehicle_Plate_ID_Range            As String = "J29"
Const New_Vehicle_Body_ID_Range             As String = "K29"   ' 新車車身碼
Const New_Vehicle_Registration_Date_Range   As String = "L29"   ' 新車領牌/發照日期
'
' Range of old vehicle information
'
Const Old_Vehicle_Brand_Range               As String = "E17"   ' 舊車品牌
Const Old_Vehicle_Owner_Name_Range          As String = "F17"   ' 舊車車主
Const Old_Vehicle_Owner_ID_Range            As String = "E25"
Const Vehicle_Owner_Relation_Range          As String = "E27"   ' 新舊車主關係
Const Vehicle_Owner_Same_Location_Range     As String = "F27"   ' 新舊車主同戶籍
Const Old_Vehicle_Type_Range                As String = "G17"   ' 舊車車別
Const Old_Vehicle_Plate_ID_Range            As String = "H17"
Const Old_Vehicle_Body_ID_Range             As String = "I17"   ' 舊車車身碼
Const Old_Vehicle_Engine_ID_Range           As String = "J17"   ' 舊車引擎碼
Const Old_Vehicle_Factory_Date_Range        As String = "K17"
Const Old_Vehicle_Registration_Date_Range   As String = "L17"
Const Old_Vehicle_Customs_ID_Range          As String = "E21"
Const Old_Vehicle_Customs_Date_Range        As String = "F23"
Const Old_Vehicle_Recycle_Date_Range        As String = "F21"
Const Old_Vehicle_Change_Date_Range         As String = "E19"
Const Old_Vehicle_Export_Declaration_Range  As String = "E23"   ' 舊車出口報單號碼


'
' Function routine handling open checklist file button click event
' TODO :
'   1. check data validity, mandatory fields
'   2. check duplicated data
'   3. check if data changed before saving the record
'
Private Sub btnOpenChecklist_Click()
    Dim fDialog As Office.FileDialog
    Dim selectedFolder As String
    Dim selectedFile   As String
    
    Dim targetWorkbook  As Excel.Workbook
    Dim targetWorksheet As Excel.Worksheet
    Dim targetWorksheetTitle As String
    Dim subformAddNewRecord As Form
        
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .InitialFileName = Me.textChecklistFolder.Value
        .AllowMultiSelect = False
        .Title = "選取檢核表"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls"
        .Filters.Add "All Files", "*.*"
        If .Show = False Then GoTo StopCode
        selectedFile = .SelectedItems(1)
    End With
    Debug.Print "選取檢核表為 " & selectedFile
    
    Me.textChecklistFile.Value = Right(selectedFile, Len(selectedFile) - Len(Me.textChecklistFolder.Value) - 1)
    Set targetWorkbook = Me.Open_ExistingWorkbook_to_Object(selectedFile)
    Set targetWorksheet = targetWorkbook.ActiveSheet
'
' Experiment to check if I can extract data stored in an Excel file
'
    targetWorksheetTitle = targetWorksheet.Range("B5")
    Debug.Print "worksheet title : " & targetWorksheetTitle
'
' Extract commodity tax rebate details and populate the data to database tables
'
    Set subformAddNewRecord = Me!formEmbeddedAddNewRecord.Form
    
'
' 退稅申請資料
'
    subformAddNewRecord!textEngineID.Value = targetWorksheet.Range(Vehicle_Body_ID_Range)
    subformAddNewRecord!textAgent.Value = targetWorksheet.Range(Dealer_Range)
    subformAddNewRecord!textUndertaker.Value = targetWorksheet.Range(Undertaker_Range)
    subformAddNewRecord!textBrandName.Value = targetWorksheet.Range(Brand_Range)
    subformAddNewRecord!textApplicationDate.Value = targetWorksheet.Range(Submit_Date_Range)
    subformAddNewRecord!textTaxRebatePayee.Value = targetWorksheet.Range(Tax_Rebate_Payee_Range)
    subformAddNewRecord!textTaxRebateReason.Value = targetWorksheet.Range(Cause_to_Refund_Range)
    subformAddNewRecord!textArrivalDate.Value = targetWorksheet.Range(Import_Date_Range)
    subformAddNewRecord!textArrivalPort.Value = targetWorksheet.Range(Arrival_Port_Range)
'
' 舊車資料
'
    subformAddNewRecord!textOldVehicleBrand.Value = targetWorksheet.Range(Old_Vehicle_Brand_Range)
    subformAddNewRecord!textOldVehicleOwnerName.Value = targetWorksheet.Range(Old_Vehicle_Owner_Name_Range)
    subformAddNewRecord!textOldVehicleType.Value = targetWorksheet.Range(Old_Vehicle_Type_Range)
    subformAddNewRecord!textOldVehiclePlateID.Value = targetWorksheet.Range(Old_Vehicle_Plate_ID_Range)
    subformAddNewRecord!textOldVehicleBodyID.Value = targetWorksheet.Range(Old_Vehicle_Body_ID_Range)
    subformAddNewRecord!textOldVehicleEngineID.Value = targetWorksheet.Range(Old_Vehicle_Engine_ID_Range)
    subformAddNewRecord!textOldVehicleFactoryDate.Value = targetWorksheet.Range(Old_Vehicle_Factory_Date_Range)
    subformAddNewRecord!textOldVehicleRegistrationDate.Value = targetWorksheet.Range(Old_Vehicle_Registration_Date_Range)
    subformAddNewRecord!textOldVehicleChangeRegistrationDate.Value = targetWorksheet.Range(Old_Vehicle_Change_Date_Range)
    subformAddNewRecord!textOldVehicleControlID.Value = targetWorksheet.Range(Old_Vehicle_Customs_ID_Range)
    subformAddNewRecord!textOldVehicleRecycleDate.Value = targetWorksheet.Range(Old_Vehicle_Recycle_Date_Range)
    subformAddNewRecord!textOldVehicleExportDeclarationNumber.Value = targetWorksheet.Range(Old_Vehicle_Export_Declaration_Range)
    subformAddNewRecord!textOldVehicleExportDate.Value = targetWorksheet.Range(Old_Vehicle_Customs_Date_Range)
    subformAddNewRecord!textOldVehicleOwnerID.Value = targetWorksheet.Range(Old_Vehicle_Owner_ID_Range)
    subformAddNewRecord!textOwnerRelationship.Value = targetWorksheet.Range(Vehicle_Owner_Relation_Range)
    subformAddNewRecord!textOwnerSameRegistrationLocation.Value = targetWorksheet.Range(Vehicle_Owner_Same_Location_Range)
'
' 新車資料
'
    subformAddNewRecord!textNewVehicleFatoryDate.Value = targetWorksheet.Range(New_Vehicle_Factory_Date_Range)
    subformAddNewRecord!textNewVehicleOwnerName.Value = targetWorksheet.Range(New_Vehicle_Owner_Name_Range)
    subformAddNewRecord!textNewVehicleOwnerID.Value = targetWorksheet.Range(New_Vehicle_Owner_ID_Range)
    subformAddNewRecord!textNewVehicleType.Value = targetWorksheet.Range(New_Vehicle_Type_Range)
    subformAddNewRecord!textNewVehicleModel.Value = targetWorksheet.Range(New_Vehicle_Model_Range)
    subformAddNewRecord!textNewVehiclePlateID.Value = targetWorksheet.Range(New_Vehicle_Plate_ID_Range)
    subformAddNewRecord!textNewVehicleBodyID.Value = targetWorksheet.Range(New_Vehicle_Body_ID_Range)
    subformAddNewRecord!textNewVehicleRegistrationDate.Value = targetWorksheet.Range(New_Vehicle_Registration_Date_Range)
        
    targetWorkbook.Close savechanges:=False
        


StopCode:
    Set fDialog = Nothing
End Sub


'
' Demonstrates how a function works and an object is returned by this function
'
' ToDo :
'       - check if specified file exists or not
'
Function Open_ExistingWorkbook_to_Object(ByVal fileName As String) As Workbook
    Set Open_ExistingWorkbook_to_Object = Workbooks.Open(fileName)
End Function



'
' Pop out directory select dialog and return selected folder
'
Public Function GetFolder(strPath As String) As String
    Dim fldr As FileDialog
    Dim sItem As String

    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "選取資料夾"
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With

NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

'
' Get_ImportFolderSetting() returns default or stored folder of source 檢核表
'
Function Get_ImportFolderSetting() As String
    Dim ImportFolder As String
    
    ImportFolder = GetSetting(ApplicationName, RegistrySectionName, RegistryChecklistFolder, _
                        "C:\")
                        
    Get_ImportFolderSetting = ImportFolder
End Function

'
' Select the default folder of checklist files
' TODO : save selected file folder to a registry for later use, e.g.  HKEY_CURRENT_USER\Software\VB and VBA Program Settings\<application name>\<registry name>
'
Private Sub btnSelectFolder_Click()
    Dim Directory_String As String
    
    If Me.textChecklistFolder.Value <> "" Then
        Directory_String = GetFolder(Me.textChecklistFolder.Value)
    Else
        Directory_String = GetFolder("C:\")
    End If
    
    Me.textChecklistFolder.Value = Directory_String
'
'   Stores selected folder to system registry
'
    Call SaveSetting(ApplicationName, RegistrySectionName, RegistryChecklistFolder, Directory_String)
    
End Sub

Private Sub cmdClearSearchQuiteria_Click()
    textNewVehicleOwnerFilter.Value = ""
    textNewVehiclePlateIDFilter.Value = ""
    textNewVehicleEngineIDFilter.Value = ""
End Sub

'
' 選擇控管表，匯入資料庫 tblApplicationDetails 及 tblControl
'
Private Sub cmdSelectFile_Click()
    Dim fDialog As Office.FileDialog
    Dim selectedFolder As String
    Dim selectedFile As String
    
    Dim targetWorkbook As Excel.Workbook
    Dim targetWorksheet As Excel.Worksheet
    Dim targetWorksheetTitle As String
    Dim subformControl As Form
    Dim subformDetails As Form
    Dim nextControlID As String
    Dim targetStringToCheck As String
    Dim lngOriginalForeColor, lngWarningRed As Long
    lngWarningRed = RGB(255, 0, 0)
                
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .InitialFileName = Me.textChkListFolder.Value
        .AllowMultiSelect = False
        .Title = "選取檢核表"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls"
        .Filters.Add "All Files", "*.*"
        If .Show = False Then GoTo StopCode
        selectedFile = .SelectedItems(1)
    End With
    Debug.Print "選取檢核表為 " & selectedFile
    
    Me.textChklistFile.Value = Right(selectedFile, Len(selectedFile) - Len(Me.textChkListFolder.Value) - 1)
    Set targetWorkbook = Me.Open_ExistingWorkbook_to_Object(selectedFile)
    Set targetWorksheet = targetWorkbook.ActiveSheet
'
' Experiment to check if I can extract data stored in an Excel file
'
    targetWorksheetTitle = targetWorksheet.Range("B5")
    
    Set subformControl = Me!formAddNewRecord_Control.Form
    Set subformDetails = subformControl!formAddNewRecord_Details.Form
'
' Check if vehicle body ID or Engine exists or not
' Only populating the data if not existing
'
    targetStringToCheck = targetWorksheet.Range(Vehicle_Body_ID_Range)
    If Me.isEngineIDExisted(targetStringToCheck) = True Then
        MsgBox ("Duplicated record, Engine ID exists!")
        GoTo StopCode
    End If
'
' 退稅申請基本資料
'
    subformControl!textBodyID.Value = targetWorksheet.Range(Vehicle_Body_ID_Range)
    subformControl!textAgent.Value = targetWorksheet.Range(Dealer_Range)
    subformControl!textUndertaker.Value = targetWorksheet.Range(Undertaker_Range)
    subformControl!textNewVehicleBrand.Value = targetWorksheet.Range(Brand_Range)
    
'
' need to checck if Submit Date is of correct date format
'
    subformControl!textApplicationDate.ForeColor = RGB(64, 64, 64)
    targetStringToCheck = targetWorksheet.Range(Submit_Date_Range)
    If (IsDate(targetStringToCheck)) Then
        subformControl!textApplicationDate.Value = targetStringToCheck
    ElseIf (IsDate(Replace(targetStringToCheck, ".", "/"))) Then
        subformControl!textApplicationDate.Value = Replace(targetStringToCheck, ".", "/")
        subformControl!listProcessInfo.AddItem "檢核表 " & Me.textChklistFile.Value
        subformControl!listProcessInfo.AddItem ">>> 申辦日期輸入值為" & targetStringToCheck & _
            "更正為" & subformControl!textApplicationDate.Value
    Else
        subformControl!textApplicationDate.Value = "1900/1/1"
        lngOriginalForeColor = subformControl!textApplicationDate.BackColor
        MsgBox "請檢查申辦日期資料格式" & vbCrLf & _
            "原輸入申辦日期 : " & targetStringToCheck, vbOKOnly
        subformControl!listProcessInfo.AddItem "檢核表 " & Me.textChklistFile.Value
        subformControl!listProcessInfo.AddItem ">>> 申辦日期輸入值為" & targetStringToCheck
    End If
    
    subformControl!textCheckPayee.Value = targetWorksheet.Range(Tax_Rebate_Payee_Range)
    subformControl!textRebateReason.Value = targetWorksheet.Range(Cause_to_Refund_Range)
'
' need to check if Import Date is of correct date format
'
    subformControl!textArrivalDate.ForeColor = RGB(64, 64, 64)
    targetStringToCheck = targetWorksheet.Range(Import_Date_Range)
    If (IsDate(targetStringToCheck)) Then
        subformControl!textArrivalDate.Value = targetStringToCheck
    ElseIf (IsDate(Replace(targetStringToCheck, ".", "/"))) Then
        targetStringToCheck = Replace(targetStringToCheck, ".", "/")
        subformControl!textArrivalDate.Value = targetStringToCheck
        subformControl!listProcessInfo.AddItem "檢核表 " & Me.textChklistFile.Value
        subformControl!listProcessInfo.AddItem ">>> 到港日期輸入值為" & targetStringToCheck & _
            "更正為" & subformControl!textApplicationDate.Value
    Else
        subformControl!textArrivalDate.Value = "1900/1/1"
        lngOriginalForeColor = subformControl!textArrivalDate.ForeColor
        subformControl!textArrivalDate.ForeColor = RGB(255, 0, 0)
        MsgBox "請檢查到港日期資料格式" & vbCrLf & _
            "原輸入到港日期 : " & targetStringToCheck, vbOKOnly
        subformControl!listProcessInfo.AddItem "檢核表 " & Me.textChklistFile.Value
        subformControl!listProcessInfo.AddItem ">>> 到港日期輸入值為" & targetStringToCheck
    End If
    subformControl!textArrivalPort.Value = targetWorksheet.Range(Arrival_Port_Range)
        
    nextControlID = Next_Control_ID()
    subformControl!textControlID.Value = nextControlID
    subformControl!textChecklistFile.Value = Me.textChklistFile.Value
'
' 舊車資料
'
    subformDetails!textOldVehicleBrand.Value = targetWorksheet.Range(Old_Vehicle_Brand_Range)
    subformDetails!textOldVehicleOwnerName.Value = targetWorksheet.Range(Old_Vehicle_Owner_Name_Range)
    subformDetails!textOldVehicleOwnerID.Value = targetWorksheet.Range(Old_Vehicle_Owner_ID_Range)
    subformDetails!textOldVehicleTypeCode.Value = targetWorksheet.Range(Old_Vehicle_Type_Range)
    subformDetails!textOldVehiclePlateID.Value = targetWorksheet.Range(Old_Vehicle_Plate_ID_Range)
    subformDetails!textOldVehicleBodyID.Value = targetWorksheet.Range(Old_Vehicle_Body_ID_Range)
    subformDetails!textOldVehicleEngineID.Value = targetWorksheet.Range(Old_Vehicle_Engine_ID_Range)
    subformDetails!textOldVehicleFactoryDate.Value = targetWorksheet.Range(Old_Vehicle_Factory_Date_Range)
    subformDetails!textOldVehicleRegistrationDate.Value = targetWorksheet.Range(Old_Vehicle_Registration_Date_Range)
    subformDetails!textOldVehicleRecycleDate.Value = targetWorksheet.Range(Old_Vehicle_Change_Date_Range)
    subformDetails!textOldVehicleRecycleControlID.Value = targetWorksheet.Range(Old_Vehicle_Customs_ID_Range)
    subformDetails!textOldVehicleRecycleControlDate.Value = targetWorksheet.Range(Old_Vehicle_Recycle_Date_Range)
    subformDetails!textOldVehicleExportDeclarationID.Value = targetWorksheet.Range(Old_Vehicle_Export_Declaration_Range)
    subformDetails!textOldVehicleExportDeclarationDate.Value = targetWorksheet.Range(Old_Vehicle_Customs_Date_Range)
    subformDetails!textVehicleOwnerRelationship.Value = targetWorksheet.Range(Vehicle_Owner_Relation_Range)
    If (targetWorksheet.Range(Vehicle_Owner_Same_Location_Range) = "Y") Then
        subformDetails!textVehicleOwnerSameHousehold.Value = "True"
    ElseIf (targetWorksheet.Range(Vehicle_Owner_Same_Location_Range) = "N") Then
        subformDetails!textVehicleOwnerSameHousehold.Value = "False"
    Else    ' N/A
        subformDetails!textVehicleOwnerSameHousehold.Value = "True"
    End If
    
'
' 新車資料
'
    subformDetails!textNewVehicleFactoryDate.Value = targetWorksheet.Range(New_Vehicle_Factory_Date_Range)
    subformDetails!textNewVehicleOwnerName.Value = targetWorksheet.Range(New_Vehicle_Owner_Name_Range)
    subformDetails!textNewVehicleOwnerID.Value = targetWorksheet.Range(New_Vehicle_Owner_ID_Range)
    subformDetails!textNewVehicleTypeCode.Value = targetWorksheet.Range(New_Vehicle_Type_Range)
    subformDetails!textNewVehicleModel.Value = targetWorksheet.Range(New_Vehicle_Model_Range)
    subformDetails!textNewVehiclePlateID.Value = targetWorksheet.Range(New_Vehicle_Plate_ID_Range)
    subformDetails!textNewVehicleEngineID.Value = targetWorksheet.Range(New_Vehicle_Body_ID_Range)
    subformDetails!textNewVehicleRegistrationDate.Value = targetWorksheet.Range(New_Vehicle_Registration_Date_Range)
    
    subformDetails!textControlID.Value = nextControlID
      
    targetWorkbook.Close savechanges:=False
    
StopCode:

End Sub

'
' isValidDate(targetStringToCheck)
'
Function isValidDate(tString As String) As Boolean
    Dim tDate As Date
    tDate = CDate(tString)
    
    If tDate.NumberFormat = "yyyy/mm/dd" Then
        isValidDate = True
    ElseIf tDate.NumberFormat = "yyyy.mm.dd" Then
        isValidDate = True
    Else
        isValidDate = False
    End If
End Function

'
' isEngineIDExisted(tString as String) checks if an application for EngineID exists or not
'
Function isEngineIDExisted(tString As String) As Boolean
    Dim sqlString As String
    
    sqlString = "Select count(*) from tblControl where EngineID=""" & tString & """;"
    Debug.Print "isEngineIDExisted exectues the SQL command " & vbNewLine & vbTab & sqlString
    
    If CurrentDb.OpenRecordset(sqlString).Fields(0) > 0 Then
        isEngineIDExisted = True
        Debug.Print "EngineID " & tString & " exists!"
    Else
        isEngineIDExisted = False
        Debug.Print "EngineID " & tString & " does not exist!"
    End If
    
End Function
'
' Subroutine cmdSelectFolder selects the target folder of 檢核表
'
Private Sub cmdSelectFolder_Click()
    Dim Directory_String As String
    
    If Me.textChkListFolder.Value <> "" Then
        Directory_String = GetFolder(Me.textChecklistFolder.Value)
    Else
        Directory_String = GetFolder("C:\")
    End If
    
    Me.textChkListFolder.Value = Directory_String
'
'   Stores selected folder to system registry
'
    Call SaveSetting(ApplicationName, RegistrySectionName, RegistryChecklistFolder, Directory_String)
End Sub

Private Sub Form_Load()
'    Me.tabControlMain.Pages(2).Controls!textChecklistFolder.Value = _
'        Me.Get_ImportFolderSetting()
    Me.tabControlMain.Pages(2).Controls!textChkListFolder.Value = _
        Me.Get_ImportFolderSetting()
    Me!formAddNewRecord_Control.Form!listProcessInfo.RowSource = ""
            
    Debug.Print "Form_Load of formApplicationLookup is called"
End Sub

Private Sub tabControlMain_Change()
    If (Me.tabControlMain.Value = 2) Then
        Debug.Print "Entering tab 新增退稅資料"
    Else
        Debug.Print "I just changed to tab " & CInt(Me.tabControlMain.Value)
    End If
End Sub

Private Sub textNewVehicleEngineIDFilter_Change()
    Debug.Print "textNewVehicleOwnerFilter = " & textNewVehicleOwnerFilter.Value
    Debug.Print "textNewVehicleEngineIDFilter = " & textNewVehicleEngineIDFilter.Value
    Debug.Print "textNewVehiclePlateIDFilter = " & textNewVehiclePlateIDFilter.Value
    
    textNewVehicleOwnerFilter.Value = ""
    textNewVehiclePlateIDFilter.Value = ""
    Me.Requery
End Sub

Private Sub textNewVehicleOwnerFilter_Change()
    Debug.Print "textNewVehicleOwnerFilter = " & textNewVehicleOwnerFilter.Value
    Debug.Print "textNewVehicleEngineIDFilter = " & textNewVehicleEngineIDFilter.Value
    Debug.Print "textNewVehiclePlateIDFilter = " & textNewVehiclePlateIDFilter.Value
    textNewVehiclePlateIDFilter.Value = ""
    textNewVehicleEngineIDFilter.Value = ""
    Me.Requery
End Sub

Private Sub textNewVehiclePlateIDFilter_Change()
    Debug.Print "textNewVehicleOwnerFilter = " & textNewVehicleOwnerFilter.Value
    Debug.Print "textNewVehicleEngineIDFilter = " & textNewVehicleEngineIDFilter.Value
    Debug.Print "textNewVehiclePlateIDFilter = " & textNewVehiclePlateIDFilter.Value
    textNewVehicleOwnerFilter.Value = ""
    textNewVehicleEngineIDFilter.Value = ""
    Me.Requery
End Sub
'
' Returns (max ControlID + 1) of tblControl
'
Function Next_Control_ID() As String
    Dim dbs As Database, rst As Recordset
    Dim nextControlID As String
    
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("SELECT Max(tblControl.[ControlID]) AS ControlIDOfMax FROM tblControl;")
    
'    rst.MoveLast

    If rst.RecordCount > 0 Then
        nextControlID = rst.Fields(0).Value
    Else
        nextControlID = "0"
    End If
    
    Debug.Print "Next ControlID is " & nextControlID
    dbs.Close
    Next_Control_ID = CStr(CInt(nextControlID) + 1)
    
End Function
