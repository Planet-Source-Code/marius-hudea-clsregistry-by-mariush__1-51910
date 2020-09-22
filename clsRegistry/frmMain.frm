VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Dim a As New clsRegistry ' load the class
MsgBox "This will test the registry functions." + vbCrLf + _
       "Start Regedit in order to see the changes."
a.CreateKey HKEY_LOCAL_MACHINE, "Software\Test"
MsgBox "Created key HKEY_LOCAL_MACHINE\Software\Test"
a.SetStringValue HKEY_LOCAL_MACHINE, "Software\Test", "Testing String Values", "Test OK"
MsgBox "Created string key Testing String Values with the value Test OK"
a.SetLongValue HKEY_LOCAL_MACHINE, "Software\Test", "Testing Long Values", 999
MsgBox "Created dword key Testing Long Values with the value 999"
a.SetBooleanValue HKEY_LOCAL_MACHINE, "Software\Test", "Testing Boolean Values", True
MsgBox "Created dword key Testing Boolean Values with the value True (1)"
MsgBox a.GetStringValue(HKEY_LOCAL_MACHINE, "Software\Test", "Testing String Values", "error"), , "String Value retreived from registry"
MsgBox a.GetLongValue(HKEY_LOCAL_MACHINE, "Software\Test", "Testing Long Values", -1), , "Long Value retreived from registry"
MsgBox a.GetBooleanValue(HKEY_LOCAL_MACHINE, "Software\Test", "Testing Boolean Values", False), , "Boolean Value retreived from registry"
a.DeleteKey HKEY_LOCAL_MACHINE, "Software\Test"
MsgBox "Key Software\Test deleted."
a.CreateKey HKEY_LOCAL_MACHINE, "Software\Test"
MsgBox "Empty key HKEY_LOCAL_MACHINE\Software\Test created"
a.DeleteEmptyKey HKEY_LOCAL_MACHINE, "Software\Test"
MsgBox "Empty key HKEY_LOCAL_MACHINE\Software\Test deleted"
Set a = Nothing  'unload the class

End Sub
