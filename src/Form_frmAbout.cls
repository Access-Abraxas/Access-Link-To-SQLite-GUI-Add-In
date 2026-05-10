VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Module Name:  Form_frmAbout
' Author:       Geoffrey Griffith
' Company:      Imagine Thought Software
' License:
'
'     Copyright © 2026 - Geoffrey L. Griffith
'
'     Permission is hereby granted, free of charge, to any person obtaining a copy
'     of this software and associated documentation files (the “Software”), to deal
'     in the Software without restriction, including without limitation the rights
'     to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'     copies of the Software, and to permit persons to whom the Software is
'     furnished to do so, subject to the following conditions:
'
'     The above copyright notice and this permission notice shall be included in
'     all copies or substantial portions of the Software.
'
'     THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'     IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'     FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'     AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'     LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'     OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'     SOFTWARE.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database
Option Explicit


Private Sub btnClose_Click()
On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSavePrompt
End Sub


Private Sub Form_Close()
On Error GoTo HandleErrors

    If (Application.Forms.Count + Application.Reports.Count) < 2 Then
        DoCmd.OpenForm "frmStart", acNormal
    End If

ExitMethod:
    Exit Sub
HandleErrors:
    MsgBox Err.Description, vbCritical, "Error " & Nz(Err.Number, "")
    Resume ExitMethod
End Sub


Private Sub Form_Open(Cancel As Integer)
On Error GoTo HandleErrors

    Me.btnClose.SetFocus

ExitMethod:
    Exit Sub
HandleErrors:
    MsgBox Err.Description, vbCritical, "Error " & Nz(Err.Number, "")
    Resume ExitMethod
End Sub
