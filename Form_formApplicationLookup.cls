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
' Constants regarding to range definition in 檢核表
'
Const Vehicle_Body_ID_Range                 As String = "C8"
Const Dealer_Range                          As String = "D8"
Const Brand_Range                           As String = "F8"
Const Submit_Date_Range                     As String = "G8"
Const Cause_to_Refund_Range                 As String = "J8"
Const Number_of_Record_Range                As String = "CA1"
'
' Range of new vehicle information
'
Const New_Vehicle_Brand_Range               As String = "F8"
Const New_Vehicle_Factory_Date_Range        As String = "E29"   ' 新車出廠年月
Const New_Vehicle_Owner_Name_Range          As String = "F29"   ' 新車車主
Const New_Vehicle_Owner_ID_Range            As String = "G29"
Const New_Vehicle_Type_Range                As String = "H29"   ' 新車車別
Const New_Vehicle_Plate_ID_Range            As String = "J29"
Const New_Vehicle_Engine_ID_Range           As String = "K29"
Const New_Vehicle_Registration_Date_Range   As String = "L29"   ' 新車領牌/發照日期
'
' Range of old vehicle information
'
Const Old_Vehicle_Owner_Name_Range          As String = "F17"   ' 舊車車主
Const Old_Vehicle_Owner_ID_Range            As String = "E25"
Const Vehicle_Owner_Relation_Range          As String = "E27"   ' 新舊車主關係
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

Private Sub btnOpenChecklist_Click()
    Dim fDialog As Office.FileDialog
    Dim selectedFolder As String
    Dim selectedFile   As String
    Dim targetWorkbook  As Excel.Workbook
    Dim targetWorksheet As Excel.Worksheet
    Dim targetWorksheetTitle As String
                
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = False
        .Title = "選取檢核表"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls"
        .Filters.Add "All Files", "*.*"
        If .Show = False Then GoTo StopCode
        selectedFile = .SelectedItems(1)
    End With
    Debug.Print "選取檢核表為 " & selectedFile
    
    Set targetWorkbook = Me.Open_ExistingWorkbook_to_Object(selectedFile)
    Set targetWorksheet = targetWorkbook.Sheets("文件檢核暨完稅證明申請表")
    targetWorksheetTitle = targetWorksheet.Range("B5")
    Debug.Print "worksheet title : " & targetWorksheetTitle

'
' Extract commodity tax rebate details and populate the data to database tables
'
    
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
