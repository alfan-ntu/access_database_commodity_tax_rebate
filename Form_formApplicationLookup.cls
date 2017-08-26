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
        
    targetWorkbook.Close
        
' Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
'    With fDialog
'        .Title = "選取檢核表資料夾"
'        .AllowMultiSelect = False
'        If .Show <> -1 Then GoTo StopCode
'        selectedFolder = .SelectedItems(1)
'    End With
'    Debug.Print "selected folder is " & selectedFolder
    

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
' Select the default folder of checklist files
' TODO : save selected file folder to a registry for later use, e.g.  HKEY_CURRENT_USER\Software\VB and VBA Program Settings\<application name>\<registry name>
'
Private Sub btnSelectFolder_Click()
    Dim Directory_String As String
    
    If Me.textChecklistFolder.Value <> "" Then
        Directory_String = GetFolder(Me.textChecklistFolder.Value)
    Else
        Directory_String = GetFolder("C:/")
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
