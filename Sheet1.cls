VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents wb As Workbook
Attribute wb.VB_VarHelpID = -1
'****************************************************************************ComboBoxes*****************************************************************************************************************
Sub Antitarnish1_ComboBox_Change()

If Antitarnish1_ComboBox.Value = "UTP" Then
    Antitarnish1_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Antitarnish1_ComboBox.Value, "DTP") Then
    Antitarnish1_TextBox.BackColor = RGB(255, 0, 0)
End If

If Antitarnish1_ComboBox.Value = "Under Quall" Or Antitarnish1_ComboBox.Value = "Tool Install" Then
    Antitarnish1_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Copper1_ComboBox_Change()

If Copper1_ComboBox.Value = "UTP" Then
    Copper1_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Copper1_ComboBox.Value, "DTP") Then
    Copper1_TextBox.BackColor = RGB(255, 0, 0)
End If

If Copper1_ComboBox.Value = "Under Quall" Or Copper1_ComboBox.Value = "Tool Install" Then
    Copper1_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Copper2_ComboBox_Change()

If Copper2_ComboBox.Value = "UTP" Then
    Copper2_TextBox.BackColor = RGB(0, 255, 0)
    'Copper2_button.BackColor = RGB(0, 255, 0)
End If

If InStr(Copper2_ComboBox.Value, "DTP") Then
    Copper2_TextBox.BackColor = RGB(255, 0, 0)
    'Copper2_button.BackColor = RGB(255, 0, 0)
End If

If Copper2_ComboBox.Value = "Under Quall" Or Copper2_ComboBox.Value = "Tool Install" Then
    Copper2_TextBox.BackColor = RGB(255, 192, 0)
    'Copper2_button.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub ORP1_Click()

Dim ORP1Value As Variant

ORP1Value = InputBox("Please Enter ORP for Chemcut 1")

If Not IsNumeric(ORP1Value) Then
MsgBox "Please enter correct number only and try again"
Exit Sub
End If


Sheets("DB").Range("AA10000").End(xlUp).Offset(1, 0) = Now()
Sheets("DB").Range("AA10000").End(xlUp).Offset(0, 1) = ORP1Value

End Sub


Sub CHEM_density1_Click()

Dim density1Value As Variant

density1Value = InputBox("Please Enter Density for Chemcut 1")

If Not IsNumeric(density1Value) Then
MsgBox "Please enter correct number only and try again"
Exit Sub
End If

Sheets("DB").Range("AA10000").End(xlUp).Offset(1, 0) = Now()
Sheets("DB").Range("AA10000").End(xlUp).Offset(0, 2) = density1Value

End Sub


Sub ORP2_Click()

Dim ORP2Value As Variant

ORP2Value = InputBox("Please Enter ORP for Chemcut 2")

If Not IsNumeric(ORP2Value) Then
MsgBox "Please enter correct number only and try again"
Exit Sub
End If

Sheets("DB").Range("AE10000").End(xlUp).Offset(1, 0) = Now()
Sheets("DB").Range("AE10000").End(xlUp).Offset(0, 1) = ORP2Value

End Sub

Sub CHEM_density2_Click()

Dim density2Value As Variant

density2Value = InputBox("Please Enter Density for Chemcut 2")

If Not IsNumeric(density2Value) Then
MsgBox "Please enter correct number only and try again"
Exit Sub
End If

Sheets("DB").Range("AE10000").End(xlUp).Offset(1, 0) = Now()
Sheets("DB").Range("AE10000").End(xlUp).Offset(0, 2) = density2Value

End Sub

Sub Sulfuric1_ComboBox_Change()

If Sulfuric1_ComboBox.Value = "UTP" Then
    Sulf1_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Sulfuric1_ComboBox.Value, "DTP") Then
    Sulf1_TextBox.BackColor = RGB(255, 0, 0)
End If

If Sulfuric1_ComboBox.Value = "Under Quall" Or Sulfuric1_ComboBox.Value = "Tool Install" Then
    Sulf1_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Sulfuric2_ComboBox_Change()

If Sulfuric2_ComboBox.Value = "UTP" Then
    Sulf2_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Sulfuric2_ComboBox.Value, "DTP") Then
    Sulf2_TextBox.BackColor = RGB(255, 0, 0)
End If

If Sulfuric2_ComboBox.Value = "Under Quall" Or Sulfuric2_ComboBox.Value = "Tool Install" Then
    Sulf2_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Strike1_ComboBox_Change()

If Strike1_ComboBox.Value = "UTP" Then
    Strike1_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Strike1_ComboBox.Value, "DTP") Then
    Strike1_TextBox.BackColor = RGB(255, 0, 0)
End If

If Strike1_ComboBox.Value = "Under Quall" Or Strike1_ComboBox.Value = "Tool Install" Then
    Strike1_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Strike2_ComboBox_Change()

If Strike2_ComboBox.Value = "UTP" Then
    Strike2_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Strike2_ComboBox.Value, "DTP") Then
    Strike2_TextBox.BackColor = RGB(255, 0, 0)
End If

If Strike2_ComboBox.Value = "Under Quall" Or Strike2_ComboBox.Value = "Tool Install" Then
    Strike2_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Activation1_ComboBox_Change()

If Activation1_ComboBox.Value = "UTP" Then
    Activation1_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Activation1_ComboBox.Value, "DTP") Then
    Activation1_TextBox.BackColor = RGB(255, 0, 0)
End If

If Activation1_ComboBox.Value = "Under Quall" Or Activation1_ComboBox.Value = "Tool Install" Then
    Activation1_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Activation2_ComboBox_Change()

If Activation2_ComboBox.Value = "UTP" Then
    Activation2_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Activation2_ComboBox.Value, "DTP") Then
    Activation2_TextBox.BackColor = RGB(255, 0, 0)
End If

If Activation2_ComboBox.Value = "Under Quall" Or Activation2_ComboBox.Value = "Tool Install" Then
    Activation2_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub HCL1_ComboBox_Change()

If HCL1_ComboBox.Value = "UTP" Then
    HCL1_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(HCL1_ComboBox.Value, "DTP") Then
    HCL1_TextBox.BackColor = RGB(255, 0, 0)
End If

If HCL1_ComboBox.Value = "Under Quall" Or HCL1_ComboBox.Value = "Tool Install" Then
    HCL1_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub HCL2_ComboBox_Change()

If HCL2_ComboBox.Value = "UTP" Then
    HCL2_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(HCL2_ComboBox.Value, "DTP") Then
    HCL2_TextBox.BackColor = RGB(255, 0, 0)
End If

If HCL2_ComboBox.Value = "Under Quall" Or HCL2_ComboBox.Value = "Tool Install" Then
    HCL2_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Copper2_All_ComboBox_Change()

If Copper2_All_ComboBox.Value = "UTP" Then
    Copper2_All_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Copper2_All_ComboBox.Value, "DTP") Then
    Copper2_All_TextBox.BackColor = RGB(255, 0, 0)
End If

If Copper2_All_ComboBox.Value = "Under Quall" Or Copper2_All_ComboBox.Value = "Tool Install" Then
    Copper2_All_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Copper1_All_ComboBox_Change()

If Copper1_All_ComboBox.Value = "UTP" Then
    Copper1_All_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Copper1_All_ComboBox.Value, "DTP") Then
    Copper1_All_TextBox.BackColor = RGB(255, 0, 0)
End If

If Copper1_All_ComboBox.Value = "Under Quall" Or Copper1_All_ComboBox.Value = "Tool Install" Then
    Copper1_All_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

' ###############################################     GOLD TOOLS     ############################################################

Sub Gold1_ComboBox_Change()

If Gold1_ComboBox.Value = "UTP" Then
    Gold1_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Gold1_ComboBox.Value, "DTP") Then
    Gold1_TextBox.BackColor = RGB(255, 0, 0)
End If

If Gold1_ComboBox.Value = "Under Quall" Or Gold1_ComboBox.Value = "Tool Install" Then
    Gold1_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Gold2_ComboBox_Change()

If Gold2_ComboBox.Value = "UTP" Then
    Gold2_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Gold2_ComboBox.Value, "DTP") Then
    Gold2_TextBox.BackColor = RGB(255, 0, 0)
End If

If Gold2_ComboBox.Value = "Under Quall" Or Gold2_ComboBox.Value = "Tool Install" Then
    Gold2_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Strike_1_2_ComboBox_Change()

If Strike_1_2_ComboBox.Value = "UTP" Then
    Strike_1_2_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Strike_1_2_ComboBox.Value, "DTP") Then
    Strike_1_2_TextBox.BackColor = RGB(255, 0, 0)
End If

If Strike_1_2_ComboBox.Value = "Under Quall" Or Strike_1_2_ComboBox.Value = "Tool Install" Then
    Strike_1_2_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Activation_1_2_ComboBox_Change()

If Activation_1_2_ComboBox.Value = "UTP" Then
    Activation_1_2_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Activation_1_2_ComboBox.Value, "DTP") Then
    Activation_1_2_TextBox.BackColor = RGB(255, 0, 0)
End If

If Activation_1_2_ComboBox.Value = "Under Quall" Or Activation_1_2_ComboBox.Value = "Tool Install" Then
    Activation_1_2_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub HCL_1_2_ComboBox_Change()

If HCL_1_2_ComboBox.Value = "UTP" Then
    HCL_1_2_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(HCL_1_2_ComboBox.Value, "DTP") Then
    HCL_1_2_TextBox.BackColor = RGB(255, 0, 0)
End If

If HCL_1_2_ComboBox.Value = "Under Quall" Or HCL_1_2_ComboBox.Value = "Tool Install" Then
    HCL_1_2_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Gold_1_2_All_ComboBox_Change()

If Gold_1_2_All_ComboBox.Value = "UTP" Then
    Gold_1_2_All_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Gold_1_2_All_ComboBox.Value, "DTP") Then
    Gold_1_2_All_TextBox.BackColor = RGB(255, 0, 0)
End If

If Gold_1_2_All_ComboBox.Value = "Under Quall" Or Gold_1_2_All_ComboBox.Value = "Tool Install" Then
    Gold_1_2_All_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Gold3_ComboBox_Change()

If Gold3_ComboBox.Value = "UTP" Then
    Gold3_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Gold3_ComboBox.Value, "DTP") Then
    Gold3_TextBox.BackColor = RGB(255, 0, 0)
End If

If Gold3_ComboBox.Value = "Under Quall" Or Gold3_ComboBox.Value = "Tool Install" Then
    Gold3_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Strike3_ComboBox_Change()

If Strike3_ComboBox.Value = "UTP" Then
    Strike3_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Strike3_ComboBox.Value, "DTP") Then
    Strike3_TextBox.BackColor = RGB(255, 0, 0)
End If

If Strike3_ComboBox.Value = "Under Quall" Or Strike3_ComboBox.Value = "Tool Install" Then
    Strike3_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Activation3_ComboBox_Change()

If Activation3_ComboBox.Value = "UTP" Then
   Activation3_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Activation3_ComboBox.Value, "DTP") Then
    Activation3_TextBox.BackColor = RGB(255, 0, 0)
End If

If Activation3_ComboBox.Value = "Under Quall" Or Activation3_ComboBox.Value = "Tool Install" Then
    Activation3_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub HCL3_ComboBox_Change()

If HCL3_ComboBox.Value = "UTP" Then
    HCL3_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(HCL3_ComboBox.Value, "DTP") Then
    HCL3_TextBox.BackColor = RGB(255, 0, 0)
End If

If HCL3_ComboBox.Value = "Under Quall" Or HCL3_ComboBox.Value = "Tool Install" Then
    HCL3_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

Sub Gold3_All_ComboBox_Change()

If Gold3_All_ComboBox.Value = "UTP" Then
    Gold3_All_TextBox.BackColor = RGB(0, 255, 0)
End If

If InStr(Gold3_All_ComboBox.Value, "DTP") Then
    Gold3_All_TextBox.BackColor = RGB(255, 0, 0)
End If

If Gold3_All_ComboBox.Value = "Under Quall" Or Gold3_All_ComboBox.Value = "Tool Install" Then
    Gold3_All_TextBox.BackColor = RGB(255, 192, 0)
End If

End Sub

' ###############################################     SRD TOOLS     ############################################################
Sub SRD1_ComboBox_Change()

If SRD1_ComboBox.Value = "UTP" Then
     SRD1_button.BackColor = RGB(0, 255, 0)
     SRD1_UP_button.BackColor = RGB(0, 255, 0)
     SRD1_Low_button.BackColor = RGB(0, 255, 0)
     
     SRD1_UP_ComboBox.Value = "UTP"
     SRD1_Low_ComboBox.Value = "UTP"
End If

If InStr(SRD1_ComboBox.Value, "DTP") Then

     SRD1_button.BackColor = RGB(255, 0, 0)
     SRD1_UP_button.BackColor = RGB(255, 0, 0)
     SRD1_Low_button.BackColor = RGB(255, 0, 0)
     
     SRD1_UP_ComboBox.Value = SRD1_ComboBox.Value
     SRD1_Low_ComboBox.Value = SRD1_ComboBox.Value

End If

If SRD1_ComboBox.Value = "Under Quall" Or SRD1_ComboBox.Value = "Tool Install" Then
            
     SRD1_button.BackColor = RGB(255, 192, 0)
     SRD1_UP_button.BackColor = RGB(255, 192, 0)
     SRD1_Low_button.BackColor = RGB(255, 192, 0)
     
     SRD1_UP_ComboBox.Value = SRD1_ComboBox.Value
     SRD1_Low_ComboBox.Value = SRD1_ComboBox.Value
    
End If
End Sub

Sub SRD2_ComboBox_Change()

If SRD2_ComboBox.Value = "UTP" Then
     SRD2_button.BackColor = RGB(0, 255, 0)
     SRD2_UP_button.BackColor = RGB(0, 255, 0)
     SRD2_Low_button.BackColor = RGB(0, 255, 0)
     
     SRD2_UP_ComboBox.Value = "UTP"
     SRD2_Low_ComboBox.Value = "UTP"
End If

If InStr(SRD2_ComboBox.Value, "DTP") Then

     SRD2_button.BackColor = RGB(255, 0, 0)
     SRD2_UP_button.BackColor = RGB(255, 0, 0)
     SRD2_Low_button.BackColor = RGB(255, 0, 0)
     
     SRD2_UP_ComboBox.Value = SRD2_ComboBox.Value
     SRD2_Low_ComboBox.Value = SRD2_ComboBox.Value

End If

If SRD2_ComboBox.Value = "Under Quall" Or SRD2_ComboBox.Value = "Tool Install" Then
            
     SRD2_button.BackColor = RGB(255, 192, 0)
     SRD2_UP_button.BackColor = RGB(255, 192, 0)
     SRD2_Low_button.BackColor = RGB(255, 192, 0)
     
     SRD2_UP_ComboBox.Value = SRD2_ComboBox.Value
     SRD2_Low_ComboBox.Value = SRD2_ComboBox.Value
    
End If
End Sub

Sub SRD3_ComboBox_Change()

If SRD3_ComboBox.Value = "UTP" Then
     SRD3_button.BackColor = RGB(0, 255, 0)
End If

If InStr(SRD3_ComboBox.Value, "DTP") Then
     SRD3_button.BackColor = RGB(255, 0, 0)
End If

If SRD3_ComboBox.Value = "Under Quall" Or SRD3_ComboBox.Value = "Tool Install" Then
     SRD3_button.BackColor = RGB(255, 192, 0)
End If
End Sub

Sub SRD4_ComboBox_Change()

If SRD4_ComboBox.Value = "UTP" Then
     SRD4_button.BackColor = RGB(0, 255, 0)
     SRD4_UP_button.BackColor = RGB(0, 255, 0)
     SRD4_Low_button.BackColor = RGB(0, 255, 0)
     
     SRD4_UP_ComboBox.Value = "UTP"
     SRD4_Low_ComboBox.Value = "UTP"
End If

If InStr(SRD4_ComboBox.Value, "DTP") Then

     SRD4_button.BackColor = RGB(255, 0, 0)
     SRD4_UP_button.BackColor = RGB(255, 0, 0)
     SRD4_Low_button.BackColor = RGB(255, 0, 0)
     
     SRD4_UP_ComboBox.Value = SRD4_ComboBox.Value
     SRD4_Low_ComboBox.Value = SRD4_ComboBox.Value

End If

If SRD4_ComboBox.Value = "Under Quall" Or SRD4_ComboBox.Value = "Tool Install" Then
            
     SRD4_button.BackColor = RGB(255, 192, 0)
     SRD4_UP_button.BackColor = RGB(255, 192, 0)
     SRD4_Low_button.BackColor = RGB(255, 192, 0)
     
     SRD4_UP_ComboBox.Value = SRD4_ComboBox.Value
     SRD4_Low_ComboBox.Value = SRD4_ComboBox.Value
    
End If
End Sub

Sub SRD5_ComboBox_Change()

If SRD5_ComboBox.Value = "UTP" Then
     SRD5_button.BackColor = RGB(0, 255, 0)
     SRD5_UP_button.BackColor = RGB(0, 255, 0)
     SRD5_Low_button.BackColor = RGB(0, 255, 0)
     
     SRD5_UP_ComboBox.Value = "UTP"
     SRD5_Low_ComboBox.Value = "UTP"
End If

If InStr(SRD5_ComboBox.Value, "DTP") Then

     SRD5_button.BackColor = RGB(255, 0, 0)
     SRD5_UP_button.BackColor = RGB(255, 0, 0)
     SRD5_Low_button.BackColor = RGB(255, 0, 0)
     
     SRD5_UP_ComboBox.Value = SRD5_ComboBox.Value
     SRD5_Low_ComboBox.Value = SRD5_ComboBox.Value

End If

If SRD5_ComboBox.Value = "Under Quall" Or SRD5_ComboBox.Value = "Tool Install" Then
            
     SRD5_button.BackColor = RGB(255, 192, 0)
     SRD5_UP_button.BackColor = RGB(255, 192, 0)
     SRD5_Low_button.BackColor = RGB(255, 192, 0)
     
     SRD5_UP_ComboBox.Value = SRD5_ComboBox.Value
     SRD5_Low_ComboBox.Value = SRD5_ComboBox.Value
    
End If
End Sub

Sub SRD6_ComboBox_Change()

If SRD6_ComboBox.Value = "UTP" Then
     SRD6_button.BackColor = RGB(0, 255, 0)
End If

If InStr(SRD6_ComboBox.Value, "DTP") Then
     SRD6_button.BackColor = RGB(255, 0, 0)
End If

If SRD6_ComboBox.Value = "Under Quall" Or SRD6_ComboBox.Value = "Tool Install" Then
     SRD6_button.BackColor = RGB(255, 192, 0)
End If
End Sub

Sub SRD7_ComboBox_Change()

If SRD7_ComboBox.Value = "UTP" Then
     SRD7_button.BackColor = RGB(0, 255, 0)
     SRD7_UP_button.BackColor = RGB(0, 255, 0)
     SRD7_Low_button.BackColor = RGB(0, 255, 0)
     
     SRD7_UP_ComboBox.Value = "UTP"
     SRD7_Low_ComboBox.Value = "UTP"
End If

If InStr(SRD7_ComboBox.Value, "DTP") Then

     SRD7_button.BackColor = RGB(255, 0, 0)
     SRD7_UP_button.BackColor = RGB(255, 0, 0)
     SRD7_Low_button.BackColor = RGB(255, 0, 0)
     
     SRD7_UP_ComboBox.Value = SRD7_ComboBox.Value
     SRD7_Low_ComboBox.Value = SRD7_ComboBox.Value

End If

If SRD7_ComboBox.Value = "Under Quall" Or SRD7_ComboBox.Value = "Tool Install" Then
            
     SRD7_button.BackColor = RGB(255, 192, 0)
     SRD7_UP_button.BackColor = RGB(255, 192, 0)
     SRD7_Low_button.BackColor = RGB(255, 192, 0)
     
     SRD7_UP_ComboBox.Value = SRD7_ComboBox.Value
     SRD7_Low_ComboBox.Value = SRD7_ComboBox.Value
    
End If
End Sub

Sub SRD8_ComboBox_Change()

If SRD8_ComboBox.Value = "UTP" Then
     SRD8_button.BackColor = RGB(0, 255, 0)
     SRD8_UP_button.BackColor = RGB(0, 255, 0)
     SRD8_Low_button.BackColor = RGB(0, 255, 0)
     
     SRD8_UP_ComboBox.Value = "UTP"
     SRD8_Low_ComboBox.Value = "UTP"
End If

If InStr(SRD8_ComboBox.Value, "DTP") Then

     SRD8_button.BackColor = RGB(255, 0, 0)
     SRD8_UP_button.BackColor = RGB(255, 0, 0)
     SRD8_Low_button.BackColor = RGB(255, 0, 0)
     
     SRD8_UP_ComboBox.Value = SRD8_ComboBox.Value
     SRD8_Low_ComboBox.Value = SRD8_ComboBox.Value

End If

If SRD8_ComboBox.Value = "Under Quall" Or SRD8_ComboBox.Value = "Tool Install" Then
            
     SRD8_button.BackColor = RGB(255, 192, 0)
     SRD8_UP_button.BackColor = RGB(255, 192, 0)
     SRD8_Low_button.BackColor = RGB(255, 192, 0)
     
     SRD8_UP_ComboBox.Value = SRD8_ComboBox.Value
     SRD8_Low_ComboBox.Value = SRD8_ComboBox.Value
    
End If
End Sub


'Upper/Lower chambers combox
Sub SRD1_UP_ComboBox_Change()

If SRD1_UP_ComboBox.Value = "UTP" Then
     SRD1_UP_button.BackColor = RGB(0, 255, 0)
End If

If InStr(SRD1_UP_ComboBox.Value, "DTP") Then
     SRD1_UP_button.BackColor = RGB(255, 0, 0)
End If

If SRD1_UP_ComboBox.Value = "Under Quall" Or SRD1_UP_ComboBox.Value = "Tool Install" Then
     SRD1_UP_button.BackColor = RGB(255, 192, 0)
End If
End Sub

Sub SRD1_Low_ComboBox_Change()

If SRD1_Low_ComboBox.Value = "UTP" Then
     SRD1_Low_button.BackColor = RGB(0, 255, 0)
End If

If InStr(SRD1_Low_ComboBox.Value, "DTP") Then
     SRD1_Low_button.BackColor = RGB(255, 0, 0)
End If

If SRD1_Low_ComboBox.Value = "Under Quall" Or SRD1_Low_ComboBox.Value = "Tool Install" Then
     SRD1_Low_button.BackColor = RGB(255, 192, 0)
End If
End Sub

Sub SRD2_UP_ComboBox_Change()

If SRD2_UP_ComboBox.Value = "UTP" Then
     SRD2_UP_button.BackColor = RGB(0, 255, 0)
End If

If InStr(SRD2_UP_ComboBox.Value, "DTP") Then
     SRD2_UP_button.BackColor = RGB(255, 0, 0)
End If

If SRD2_UP_ComboBox.Value = "Under Quall" Or SRD2_UP_ComboBox.Value = "Tool Install" Then
     SRD2_UP_button.BackColor = RGB(255, 192, 0)
End If
End Sub

Sub SRD2_Low_ComboBox_Change()

If SRD2_Low_ComboBox.Value = "UTP" Then
     SRD2_Low_button.BackColor = RGB(0, 255, 0)
End If

If InStr(SRD2_Low_ComboBox.Value, "DTP") Then
     SRD2_Low_button.BackColor = RGB(255, 0, 0)
End If

If SRD2_Low_ComboBox.Value = "Under Quall" Or SRD2_Low_ComboBox.Value = "Tool Install" Then
     SRD2_Low_button.BackColor = RGB(255, 192, 0)
End If
End Sub

Sub SRD4_UP_ComboBox_Change()

If SRD4_UP_ComboBox.Value = "UTP" Then
     SRD4_UP_button.BackColor = RGB(0, 255, 0)
End If

If InStr(SRD4_UP_ComboBox.Value, "DTP") Then
     SRD4_UP_button.BackColor = RGB(255, 0, 0)
End If

If SRD4_UP_ComboBox.Value = "Under Quall" Or SRD4_UP_ComboBox.Value = "Tool Install" Then
     SRD4_UP_button.BackColor = RGB(255, 192, 0)
End If
End Sub

Sub SRD4_Low_ComboBox_Change()

If SRD4_Low_ComboBox.Value = "UTP" Then
     SRD4_Low_button.BackColor = RGB(0, 255, 0)
End If

If InStr(SRD4_Low_ComboBox.Value, "DTP") Then
     SRD4_Low_button.BackColor = RGB(255, 0, 0)
End If

If SRD4_Low_ComboBox.Value = "Under Quall" Or SRD4_Low_ComboBox.Value = "Tool Install" Then
     SRD4_Low_button.BackColor = RGB(255, 192, 0)
End If
End Sub

Sub SRD5_UP_ComboBox_Change()

If SRD5_UP_ComboBox.Value = "UTP" Then
     SRD5_UP_button.BackColor = RGB(0, 255, 0)
End If

If InStr(SRD5_UP_ComboBox.Value, "DTP") Then
     SRD5_UP_button.BackColor = RGB(255, 0, 0)
End If

If SRD5_UP_ComboBox.Value = "Under Quall" Or SRD5_UP_ComboBox.Value = "Tool Install" Then
     SRD5_UP_button.BackColor = RGB(255, 192, 0)
End If
End Sub

Sub SRD5_Low_ComboBox_Change()

If SRD5_Low_ComboBox.Value = "UTP" Then
     SRD5_Low_button.BackColor = RGB(0, 255, 0)
End If

If InStr(SRD5_Low_ComboBox.Value, "DTP") Then
     SRD5_Low_button.BackColor = RGB(255, 0, 0)
End If

If SRD5_Low_ComboBox.Value = "Under Quall" Or SRD5_Low_ComboBox.Value = "Tool Install" Then
     SRD5_Low_button.BackColor = RGB(255, 192, 0)
End If
End Sub

Sub SRD7_UP_ComboBox_Change()

If SRD7_UP_ComboBox.Value = "UTP" Then
     SRD7_UP_button.BackColor = RGB(0, 255, 0)
End If

If InStr(SRD7_UP_ComboBox.Value, "DTP") Then
     SRD7_UP_button.BackColor = RGB(255, 0, 0)
End If

If SRD7_UP_ComboBox.Value = "Under Quall" Or SRD7_UP_ComboBox.Value = "Tool Install" Then
     SRD7_UP_button.BackColor = RGB(255, 192, 0)
End If
End Sub

Sub SRD7_Low_ComboBox_Change()

If SRD7_Low_ComboBox.Value = "UTP" Then
     SRD7_Low_button.BackColor = RGB(0, 255, 0)
End If

If InStr(SRD7_Low_ComboBox.Value, "DTP") Then
     SRD7_Low_button.BackColor = RGB(255, 0, 0)
End If

If SRD7_Low_ComboBox.Value = "Under Quall" Or SRD7_Low_ComboBox.Value = "Tool Install" Then
     SRD7_Low_button.BackColor = RGB(255, 192, 0)
End If
End Sub

Sub SRD8_UP_ComboBox_Change()

If SRD8_UP_ComboBox.Value = "UTP" Then
     SRD8_UP_button.BackColor = RGB(0, 255, 0)
End If

If InStr(SRD8_UP_ComboBox.Value, "DTP") Then
     SRD8_UP_button.BackColor = RGB(255, 0, 0)
End If

If SRD8_UP_ComboBox.Value = "Under Quall" Or SRD8_UP_ComboBox.Value = "Tool Install" Then
     SRD8_UP_button.BackColor = RGB(255, 192, 0)
End If
End Sub

Sub SRD8_Low_ComboBox_Change()

If SRD8_Low_ComboBox.Value = "UTP" Then
     SRD8_Low_button.BackColor = RGB(0, 255, 0)
End If

If InStr(SRD8_Low_ComboBox.Value, "DTP") Then
     SRD8_Low_button.BackColor = RGB(255, 0, 0)
End If

If SRD8_Low_ComboBox.Value = "Under Quall" Or SRD8_Low_ComboBox.Value = "Tool Install" Then
     SRD8_Low_button.BackColor = RGB(255, 192, 0)
End If
End Sub
'Upper/Lower chambers combox  ---- THE END ----

' ###############################################     THE END    ##########################################################################################################################################

Sub Antitarnish1_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "COPPER Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Antitarnish_1"

End Sub

Sub Antitarnish2_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "COPPER Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Antitarnish_2"

End Sub

Sub Copper1_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "COPPER Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Copper_1"

End Sub

Sub Sulfuric1_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "COPPER Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Sulfuric_1"

End Sub

Sub Strike1_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "COPPER Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Strike_1"

End Sub

Sub Activation1_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "COPPER Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Activation_1"

End Sub

Sub HCL1_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "COPPER Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "HCL_1"

End Sub

Sub Copper2_button_Click()
Copper2_button.Caption = "Copper 2" & vbLf & vbLf & "205"
ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "COPPER Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Copper_2"

End Sub

Sub Sulfric2_button_Click()
Sulfric2_button.Caption = "Sulfuric" & vbLf & vbLf & "207"
ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "COPPER Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Sulfuric_2"
End Sub

Sub Strike2_button_Click()
Strike2_button.Caption = "Strike" & vbLf & vbLf & "208"
ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "COPPER Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Strike_2"
End Sub

Sub Activation2_button_Click()
Activation2_button.Caption = "Activ." & vbLf & vbLf & "212"
ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "COPPER Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Activation_2"
End Sub

Sub HCL2_button_Click()
HCL2_button.Caption = "HCL" & vbLf & vbLf & "213"
ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "COPPER Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "HCL_2"
End Sub

'**************************************************************************************************************************************************************************************************************
                                                                                        'DAFLS Buttons

Sub Dafl1_Button1_Click()


ThisWorkbook.Worksheets("Bath_Repl").Activate
Worksheets("Bath_Repl").Range("A1") = "DAFLS"
Worksheets("Bath_Repl").Range("C4") = "DAFL_1"
Worksheets("Bath_Repl").Range("B4") = Now

'Sheets(1).Dafl1_Button1.BackColor = RGB(0, 255, 0)


'Application.OnTime Now + TimeValue("00:00:01"), "blink", , True
'Dafl1_Button1.BackColor = RGB(255, 0, 0)

End Sub



Sub Dafl2_Button2_Click()


ThisWorkbook.Worksheets("Bath_Repl").Activate
Worksheets("Bath_Repl").Range("A1") = "DAFLS"
Worksheets("Bath_Repl").Range("C4") = "DAFL_2"
Worksheets("Bath_Repl").Range("B4") = Now
End Sub



Sub Dafl3_Button3_Click()


ThisWorkbook.Worksheets("Bath_Repl").Activate
Worksheets("Bath_Repl").Range("A1") = "DAFLS"
Worksheets("Bath_Repl").Range("C4") = "DAFL_3"
Worksheets("Bath_Repl").Range("B4") = Now

End Sub




Sub Dafl4_Button4_Click()


ThisWorkbook.Worksheets("Bath_Repl").Activate
Worksheets("Bath_Repl").Range("A1") = "DAFLS"
Worksheets("Bath_Repl").Range("C4") = "DAFL_4"
Worksheets("Bath_Repl").Range("B4") = Now
End Sub



Sub Dafl5_Button5_Click()


ThisWorkbook.Worksheets("Bath_Repl").Activate
Worksheets("Bath_Repl").Range("A1") = "DAFLS"
Worksheets("Bath_Repl").Range("C4") = "DAFL_5"
Worksheets("Bath_Repl").Range("B4") = Now
End Sub



Sub Dafl6_Button6_Click()


ThisWorkbook.Worksheets("Bath_Repl").Activate
Worksheets("Bath_Repl").Range("A1") = "DAFLS"
Worksheets("Bath_Repl").Range("C4") = "DAFL_6"
Worksheets("Bath_Repl").Range("B4") = Now
End Sub



Sub Dafl7_Button7_Click()


ThisWorkbook.Worksheets("Bath_Repl").Activate
Worksheets("Bath_Repl").Range("A1") = "DAFLS"
Worksheets("Bath_Repl").Range("C4") = "DAFL_7"
Worksheets("Bath_Repl").Range("B4") = Now
End Sub



Sub Dafl8_Button8_Click()


ThisWorkbook.Worksheets("Bath_Repl").Activate
Worksheets("Bath_Repl").Range("A1") = "DAFLS"
Worksheets("Bath_Repl").Range("C4") = "DAFL_8"
Worksheets("Bath_Repl").Range("B4") = Now
End Sub



Sub Dafl9_Button9_Click()


ThisWorkbook.Worksheets("Bath_Repl").Activate
Worksheets("Bath_Repl").Range("A1") = "DAFLS"
Worksheets("Bath_Repl").Range("C4") = "DAFL_9"
Worksheets("Bath_Repl").Range("B4") = Now
End Sub



Sub Dafl10_Button10_Click()


ThisWorkbook.Worksheets("Bath_Repl").Activate
Worksheets("Bath_Repl").Range("A1") = "DAFLS"
Worksheets("Bath_Repl").Range("C4") = "DAFL_10"
Worksheets("Bath_Repl").Range("B4") = Now
End Sub

'**************************************************************************************************************************************************************************************************************
                                                                                        'DAFL_DENSITY Buttons
                                                                                        
Sub Density1_Click()

ThisWorkbook.Worksheets("Dafl").Activate
Worksheets("Dafl").Range("B1") = "DAFL_1"
Worksheets("Dafl").Range("A3") = Now
Worksheets("Dafl").Range("G3") = Now
End Sub

Sub Density2_Click()

ThisWorkbook.Worksheets("Dafl").Activate
Worksheets("Dafl").Range("B1") = "DAFL_2"
Worksheets("Dafl").Range("A3") = Now
Worksheets("Dafl").Range("G3") = Now
End Sub

Sub Density3_Click()

ThisWorkbook.Worksheets("Dafl").Activate
Worksheets("Dafl").Range("B1") = "DAFL_3"
Worksheets("Dafl").Range("A3") = Now
Worksheets("Dafl").Range("G3") = Now
End Sub

Sub Density4_Click()

ThisWorkbook.Worksheets("Dafl").Activate
Worksheets("Dafl").Range("B1") = "DAFL_4"
Worksheets("Dafl").Range("A3") = Now
Worksheets("Dafl").Range("G3") = Now
End Sub

Sub Density5_Click()

ThisWorkbook.Worksheets("Dafl").Activate
Worksheets("Dafl").Range("B1") = "DAFL_5"
Worksheets("Dafl").Range("A3") = Now
Worksheets("Dafl").Range("G3") = Now
End Sub

Sub Density6_Click()

ThisWorkbook.Worksheets("Dafl").Activate
Worksheets("Dafl").Range("B1") = "DAFL_6"
Worksheets("Dafl").Range("A3") = Now
Worksheets("Dafl").Range("G3") = Now
End Sub

Sub Density7_Click()

ThisWorkbook.Worksheets("Dafl").Activate
Worksheets("Dafl").Range("B1") = "DAFL_7"
Worksheets("Dafl").Range("A3") = Now
Worksheets("Dafl").Range("G3") = Now
End Sub

Sub Density8_Click()

ThisWorkbook.Worksheets("Dafl").Activate
Worksheets("Dafl").Range("B1") = "DAFL_8"
Worksheets("Dafl").Range("A3") = Now
Worksheets("Dafl").Range("G3") = Now
End Sub

Sub Density9_Click()

ThisWorkbook.Worksheets("Dafl").Activate
Worksheets("Dafl").Range("B1") = "DAFL_9"
Worksheets("Dafl").Range("A3") = Now
Worksheets("Dafl").Range("G3") = Now
End Sub

Sub Density10_Click()

ThisWorkbook.Worksheets("Dafl").Activate
Worksheets("Dafl").Range("B1") = "DAFL_10"
Worksheets("Dafl").Range("A3") = Now
Worksheets("Dafl").Range("G3") = Now
End Sub


'**************************************************************************************************************************************************************************************************************
                                                                                        'DAFL_CLEAN Buttons

Sub Clean1_Click()
Call Density1_Click
End Sub
Sub Clean2_Click()
Call Density2_Click
End Sub
Sub Clean3_Click()
Call Density3_Click
End Sub
Sub Clean4_Click()
Call Density4_Click
End Sub
Sub Clean5_Click()
Call Density5_Click
End Sub
Sub Clean6_Click()
Call Density6_Click
End Sub
Sub Clean7_Click()
Call Density7_Click
End Sub
Sub Clean8_Click()
Call Density8_Click
End Sub
Sub Clean9_Click()
Call Density9_Click
End Sub
Sub Clean10_Click()
Call Density10_Click
End Sub


'**************************************************************************************************************************************************************************************************************
                                                                                        'WASHING STATIONS Buttons

Sub Ammonia1_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Washing Station 1"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Amonia_1"

End Sub

Sub Ammonia2_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Washing Station 2"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Amonia_2"

End Sub

Sub Ammonia3_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Washing Station 3"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Amonia_3"

End Sub


Sub CAN1_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Washing Station 1"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "CAN_1"

End Sub

Sub CAN2_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Washing Station 2"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "CAN_2"

End Sub

Sub CAN3_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Washing Station 3"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "CAN_3"

End Sub

Sub Enstrip1_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Washing Station 1"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Enstrip_1"

End Sub

Sub Enstrip2_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Washing Station 2"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Enstrip_2"

End Sub

Sub Enstrip3_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Washing Station 3"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Enstrip_3"

End Sub

Sub Gold1_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Gold Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Gold_1"

End Sub

Sub Gold2_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Gold Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Gold_2"

End Sub

Sub Strike_1_2_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Gold Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Strike_1_2"

End Sub

Sub Active_1_2_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Gold Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Activation_1_2"

End Sub

Sub HCL_1_2_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Gold Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "HCL_1_2"

End Sub

Sub Gold3_Click()
Gold3.Caption = "Gold 3" & vbLf & vbLf & "103"
ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Gold Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Gold_3"

End Sub

Sub Strike3_Click()

Strike3.Caption = "Strike" & vbLf & vbLf & "109"
ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Gold Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Strike_3"

End Sub

Sub Activ3_Click()

Activ3.Caption = "Activ." & vbLf & vbLf & "113"
ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Gold Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Activation_3"

End Sub

Sub HCL3_Click()

HCL3.Caption = "HCL" & vbLf & vbLf & "114"
ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Gold Plating"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "HCL_3"

End Sub

Sub CHEM1_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Chemical Etch"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "CHEM_1"

End Sub

Sub CHEM2_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Chemical Etch"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "CHEM_2"

End Sub


Sub SC_Stripping_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "SC Technic Stripping"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "SC_Technic_Stripping"

End Sub


Sub Electroclean_room27_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Electro-clean"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Electroclean_room27"

End Sub

Sub Electroclean_WE_Click()

ThisWorkbook.Sheets("Bath_Repl").Activate
Sheets("Bath_Repl").Range("A1") = "Electro-clean"
Sheets("Bath_Repl").Range("B4") = Date
Sheets("Bath_Repl").Range("C4") = "Electroclean_WE"

End Sub











'Scroll Down Comboboxes  call to subs from "Scroll_COMBOX" module
'***************************************************************************************************
Private Sub DropListTemp_GotFocus()
    Set wb = ThisWorkbook
    MakeScrollableWithMouseWheel(DropListTemp) = True
End Sub

Private Sub DropListTemp_LostFocus()
    MakeScrollableWithMouseWheel(DropListTemp) = False
End Sub


Private Sub wb_BeforeClose(Cancel As Boolean)
    If MakeScrollableWithMouseWheel(DropListTemp) Then
        MakeScrollableWithMouseWheel(DropListTemp) = False
    End If
End Sub


'***************************************************************************************************















'Auto fill List in combobox "DropListTemp"
Sub Worksheet_SelectionChange(ByVal Target As Range)

Dim xCombox As OLEObject
Dim xStr As String
Dim xWs As Worksheet
Dim xArr
Set xWs = Application.ActiveSheet
On Error Resume Next
Set xCombox = xWs.OLEObjects("DropListTemp")
With xCombox
.ListFillRange = ""
.LinkedCell = ""
.Visible = False
End With

If Target.Validation.Type = 3 Then
Target.Validation.InCellDropdown = False
Cancel = True
xStr = Target.Validation.Formula1
xStr = Right(xStr, Len(xStr) - 1)

If xStr = "" Then Exit Sub

With xCombox
.Visible = True
.Left = Target.Left
.Top = Target.Top
.Width = Target.Width + 5
.Height = Target.Height + 5
.ListFillRange = xStr

If .ListFillRange = "" Then
xArr = Split(xStr, ",")
Me.DropListTemp.List = xArr
End If
.LinkedCell = Target.Address
End With

xCombox.Activate
Me.DropListTemp.DropDown
End If
End Sub

Sub DropListTemp_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Select Case KeyCode
Case 9
Application.ActiveCell.Offset(0, 1).Activate
Case 13
Application.ActiveCell.Offset(1, 0).Activate
End Select
End Sub




