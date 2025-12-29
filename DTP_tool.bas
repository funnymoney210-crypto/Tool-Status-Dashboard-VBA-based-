Attribute VB_Name = "DTP_tool"
Option Explicit
' Sub function to down the tool selected by user

Sub Tool_DTP()

Dim EID As Long
Dim Select_Tool As Range
Dim Select_DTP As Range
Dim Select_DTPString As String
Dim Reason_for_DTP As Range
Dim cell As Range

If Range("D28") = "" Then
    Exit Sub
End If

Set Select_Tool = Range("D28")
If Range("D30") = "" Then
Range("D30") = "DTP - Waiting for Maintenance"
End If
Set Select_DTP = Range("D30")
Set Reason_for_DTP = Range("D32")


    If IsEmpty(Range("D32")) Then
        MsgBox "Please type Reason for DTP and try again"
        Exit Sub
    End If

EID = Application.InputBox("Please Enter Employee Identification" & vbCrLf & "נא הכנס מספר עובד", Type:=1) 'EID = Employee Identification

        For Each cell In Range("A100:A108") 'for example change "DTP - Waiting for Maintenance" to "DTP - Maint."
            If cell.Value = Select_DTP Then
                Select_DTPString = Cells(cell.Row, 3).Value
            End If
        Next

Select Case Range("D28")
'///////DAFLS TOOLS////////
    Case Is = "DAFL 1"
        Sheets("Takala").Range("B2").Value = "DAFL 1"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 2, 1) ' Last two numbers are cells cordinations for UTP cells for tools
    Case Is = "DAFL 2"
        Sheets("Takala").Range("B2").Value = "DAFL 2"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 2, 2)
    Case Is = "DAFL 3"
        Sheets("Takala").Range("B2").Value = "DAFL 3"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 2, 3)
     Case Is = "DAFL 4"
        Sheets("Takala").Range("B2").Value = "DAFL 4"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 2, 4)
     Case Is = "DAFL 5"
        Sheets("Takala").Range("B2").Value = "DAFL 5"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 2, 5)
    Case Is = "DAFL 6"
        Sheets("Takala").Range("B2").Value = "DAFL 6"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 2, 6)
    Case Is = "DAFL 7"
        Sheets("Takala").Range("B2").Value = "DAFL 7"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 2, 7)
    Case Is = "DAFL 8"
        Sheets("Takala").Range("B2").Value = "DAFL 8"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 2, 8)
    Case Is = "DAFL 9"
        Sheets("Takala").Range("B2").Value = "DAFL 9"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 2, 9)
    Case Is = "DAFL 10"
        Sheets("Takala").Range("B2").Value = "DAFL 10"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 2, 10)
                     
'///////Washing Station TOOLS ////////
    Case Is = "Washing Station 1"
        Sheets("Takala").Range("B2").Value = "Washing Station 1"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 12, 1)
    Case Is = "Washing Station 2"
        Sheets("Takala").Range("B2").Value = "Washing Station 2"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 12, 2)
    Case Is = "Washing Station 3"
        Sheets("Takala").Range("B2").Value = "Washing Station 3"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 12, 3)
        
'///////Chemical Etch TOOLS ////////
    Case Is = "CHEMICAL ETCH 1"
        Sheets("Takala").Range("B2").Value = "CHEMICAL ETCH 1"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 19, 1)
    Case Is = "CHEMICAL ETCH 2"
        Sheets("Takala").Range("B2").Value = "CHEMICAL ETCH 2"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 19, 2)

'///////Anti-tarnish 2 ////////
    Case Is = "ANTI-TARNISH 1"
        Sheets("Takala").Range("B2").Value = "ANTI-TARNISH 1"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 19, 3)

'///////Manual Ponsage ////////
    Case Is = "Manual Ponsage"
        Sheets("Takala").Range("B2").Value = "Manual Ponsage"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 19, 4)

'///////SC TECHNIC STRIPPING ////////
    Case Is = "SC TECHNIC STRIPPING"
        Sheets("Takala").Range("B2").Value = "SC TECHNIC STRIPPING"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 26, 1)

'///////Manual Stripping ////////
    Case Is = "Manual Stripping"
        Sheets("Takala").Range("B2").Value = "Manual Stripping"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 26, 2)

'///////Electro-clean (room 27) ////////
    Case Is = "Electro-clean (room 27)"
        Sheets("Takala").Range("B2").Value = "Electro-clean (room 27)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 26, 3)

'///////Electro-clean (WE) ////////
    Case Is = "Electro-clean (WE)"
        Sheets("Takala").Range("B2").Value = "Electro-clean (WE)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call up_downTool(Range("D28").Value, Range("D30").Value, Range("D32").Value, 26, 4)


'///////GOLD 1 ////////
    Case Is = "GOLD 1"
        Sheets("Dashboard").Gold1_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "GOLD 1"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "GOLD 2"
        Sheets("Dashboard").Gold2_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "GOLD 2"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "Strike 1-2 (GOLD)"
        Sheets("Dashboard").Strike_1_2_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "Strike 1-2 (GOLD)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "Activation 1-2 (GOLD)"
        Sheets("Dashboard").Activation_1_2_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "Activation 1-2 (GOLD)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "HCL 1-2 (Dip Gold)"
        Sheets("Dashboard").HCL_1_2_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "HCL 1-2 (Dip Gold)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "GOLD 1-2 All"
        Sheets("Dashboard").Gold_1_2_All_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "GOLD 1-2 All"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail

'///////GOLD 3 ///////
    Case Is = "GOLD 3 (103)"
        Sheets("Dashboard").Gold3_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "GOLD 3 (103)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "Strike 3 (109)"
        Sheets("Dashboard").Strike3_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "Strike 3 (109)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "Activation 3 (113)"
        Sheets("Dashboard").Activation3_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "Activation 3 (113)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "HCL 3 (114)"
        Sheets("Dashboard").HCL3_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "HCL 3 (114)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "GOLD 3 All"
        Sheets("Dashboard").Gold3_All_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "GOLD 3 All"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail

'///////COPPER 1 ///////
    Case Is = "ANTI-TARNISH 1"
        Sheets("Dashboard").Antitarnish1_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "ANTI-TARNISH 1"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "Copper 1"
        Sheets("Dashboard").Copper1_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "Copper 1"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "Sulfuric 1"
        Sheets("Dashboard").Sulfuric1_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "Sulfuric 1"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "Strike 1 (COPPER)"
        Sheets("Dashboard").Strike1_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "Strike 1 (COPPER)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "Strike 1"
        Sheets("Dashboard").Strike1_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "Strike 1"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "Activation 1 (COPPER)"
        Sheets("Dashboard").Activation1_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "Activation 1 (COPPER)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "Activation 1"
        Sheets("Dashboard").Activation1_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "Activation 1"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "HCL 1 (Dip Copper)"
        Sheets("Dashboard").HCL1_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "HCL 1 (Dip Copper)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "HCL 1"
        Sheets("Dashboard").HCL1_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "HCL 1"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "Copper 1 All"
        Sheets("Dashboard").Copper1_All_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "Copper 1 All"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail



'///////COPPER 2 ///////
    Case Is = "Copper 2 (205)"
        Sheets("Dashboard").Copper2_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "Copper 2 (205)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "Sulfuric 2 (207)"
        Sheets("Dashboard").Sulfuric2_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "Sulfuric 2 (207)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "Strike 2 (208)"
        Sheets("Dashboard").Strike2_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "Strike 2 (208)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "Activation 2 (212)"
        Sheets("Dashboard").Activation2_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "Activation 2 (212)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "HCL 2 (213)"
        Sheets("Dashboard").HCL2_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "HCL 3 (114)"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "Copper 2 All"
        Sheets("Dashboard").Copper2_All_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "Copper 2 All"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail


'///////COPPER 2 ///////
    Case Is = "SRD 1"
        Sheets("Dashboard").SRD1_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 1"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "SRD 1_Upper"
        Sheets("Dashboard").SRD1_UP_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 1_Upper"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "SRD 1_Lower"
        Sheets("Dashboard").SRD1_Low_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 1_Lower"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
        
    Case Is = "SRD 2"
        Sheets("Dashboard").SRD2_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 2"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
     Case Is = "SRD 2_Upper"
        Sheets("Dashboard").SRD2_UP_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 2_Upper"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "SRD 2_Lower"
        Sheets("Dashboard").SRD2_Low_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 2_Lower"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
        
    Case Is = "SRD 3"
        Sheets("Dashboard").SRD3_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 3"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "SRD 4"
        Sheets("Dashboard").SRD4_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 4"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
     Case Is = "SRD 4_Upper"
        Sheets("Dashboard").SRD4_UP_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 4_Upper"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "SRD 4_Lower"
        Sheets("Dashboard").SRD4_Low_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 4_Lower"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail

    Case Is = "SRD 5"
        Sheets("Dashboard").SRD5_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 5"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
     Case Is = "SRD 5_Upper"
        Sheets("Dashboard").SRD5_UP_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 5_Upper"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "SRD 5_Lower"
        Sheets("Dashboard").SRD5_Low_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 5_Lower"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail

    Case Is = "SRD 6"
        Sheets("Dashboard").SRD6_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 6"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "SRD 7"
        Sheets("Dashboard").SRD7_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 7"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
     Case Is = "SRD 7_Upper"
        Sheets("Dashboard").SRD7_UP_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 7_Upper"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "SRD 7_Lower"
        Sheets("Dashboard").SRD7_Low_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 7_Lower"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail


    Case Is = "SRD 8"
        Sheets("Dashboard").SRD8_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 8"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
     Case Is = "SRD 8_Upper"
        Sheets("Dashboard").SRD8_UP_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 8_Upper"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail
    Case Is = "SRD 8_Lower"
        Sheets("Dashboard").SRD8_Low_ComboBox.Value = Select_DTPString
        Sheets("Takala").Range("B2").Value = "SRD 8_Lower"
        Sheets("Takala").Range("B3").Value = Reason_for_DTP
        Sheets("Takala").Range("B4") = Now()
        Call CreateOutlookEmail




End Select






ThisWorkbook.Sheets("DB").Select

Range("L10000").End(xlUp).Offset(1, 0) = Now()
Range("L10000").End(xlUp).Offset(0, 1) = Select_Tool
Range("L10000").End(xlUp).Offset(0, 2) = Select_DTP
Range("L10000").End(xlUp).Offset(0, 3) = Reason_for_DTP
Range("L10000").End(xlUp).Offset(0, 4) = "DTP"
Range("L10000").End(xlUp).Offset(0, 5) = EID


ThisWorkbook.Sheets(1).Select
        Range("D28") = ""
        Range("D30") = ""
        Range("D32").ClearContents
ActiveWorkbook.Save

End Sub








'                        D28                    D30                      D32
Function up_downTool(Select_Tool As String, Select_DTP As String, Reason_for_DTP As String, Cell_Xlocation As Long, Cell_Ylocation As Long)


If (Select_DTP = "DTP - Waiting for Maintenance" Or Select_DTP = "") Then
    Cells(Cell_Xlocation, Cell_Ylocation) = "DTP - Waiting for Maintenance"
   On Error Resume Next
    Cells(Cell_Xlocation, Cell_Ylocation).Comment.Delete
   On Error GoTo 0
    Cells(Cell_Xlocation, Cell_Ylocation).AddComment   'ADD Comments for Reason for DTP
    Cells(Cell_Xlocation, Cell_Ylocation).Comment.Visible = False
   On Error Resume Next
    Cells(Cell_Xlocation, Cell_Ylocation).Comment.Text Text:=Reason_for_DTP
   On Error GoTo 0
    Call CreateOutlookEmail  'Function send email
Else
    Cells(Cell_Xlocation, Cell_Ylocation) = Select_DTP
    On Error Resume Next
   Cells(Cell_Xlocation, Cell_Ylocation).Comment.Delete
   On Error GoTo 0
    Cells(Cell_Xlocation, Cell_Ylocation).AddComment   'ADD Comments for Reason for DTP
    Cells(Cell_Xlocation, Cell_Ylocation).Comment.Visible = False
   On Error Resume Next
    Cells(Cell_Xlocation, Cell_Ylocation).Comment.Text Text:=Reason_for_DTP
   On Error GoTo 0
    Call CreateOutlookEmail  'Function send email
End If


End Function



