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
' Constants regarding to range definition in �ˮ֪�
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
Const New_Vehicle_Factory_Date_Range        As String = "E29"   ' �s���X�t�~��
Const New_Vehicle_Owner_Name_Range          As String = "F29"   ' �s�����D
Const New_Vehicle_Owner_ID_Range            As String = "G29"
Const New_Vehicle_Type_Range                As String = "H29"   ' �s�����O
Const New_Vehicle_Plate_ID_Range            As String = "J29"
Const New_Vehicle_Engine_ID_Range           As String = "K29"
Const New_Vehicle_Registration_Date_Range   As String = "L29"   ' �s����P/�o�Ӥ��
'
' Range of old vehicle information
'
Const Old_Vehicle_Owner_Name_Range          As String = "F17"   ' �¨����D
Const Old_Vehicle_Owner_ID_Range            As String = "E25"
Const Vehicle_Owner_Relation_Range          As String = "E27"   ' �s�¨��D���Y
Const Old_Vehicle_Type_Range                As String = "G17"   ' �¨����O
Const Old_Vehicle_Plate_ID_Range            As String = "H17"
Const Old_Vehicle_Body_ID_Range             As String = "I17"   ' �¨������X
Const Old_Vehicle_Engine_ID_Range           As String = "J17"   ' �¨������X
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
            
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = False
        .Title = "����ˮ֪�"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls"
        .Filters.Add "All Files", "*.*"
        If .Show = False Then GoTo StopCode
        selectedFile = .SelectedItems(1)
    End With
    Debug.Print "����ˮ֪��� " & selectedFile
        
' Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
'    With fDialog
'        .Title = "����ˮ֪���Ƨ�"
'        .AllowMultiSelect = False
'        If .Show <> -1 Then GoTo StopCode
'        selectedFolder = .SelectedItems(1)
'    End With
'    Debug.Print "selected folder is " & selectedFolder
    

StopCode:
    Set fDialog = Nothing
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