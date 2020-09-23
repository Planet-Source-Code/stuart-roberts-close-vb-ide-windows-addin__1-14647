VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My Add In"
   ClientHeight    =   3195
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'# Look at the readme.txt for description on how to register this DLL so it starts with VB
'# If you compile this and save a DLL VB will auto register that DLL to auto-start from the
'# DLL's location

Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private Sub CancelButton_Click()
On Error Resume Next
    Me.Hide
End Sub

Private Sub OKButton_Click()
On Error Resume Next
    Dim tpaWindows() As VBIDE.vbext_WindowType
    
    ReDim tpaWindows(2)
    tpaWindows(1) = vbext_wt_CodeWindow '# Windows like this one
    tpaWindows(2) = vbext_wt_Designer '# Windows with forms on them
'# There are around 15 or so different window types available, look and see
    
    CloseWindows tpaWindows()
End Sub

Public Function CloseWindow(tpWindow As VBIDE.vbext_WindowType) As Boolean
'# Closes all windows in IDE matching the type supplied
On Error GoTo HandleError
    Dim nLoop As Integer
    
    With VBInstance
        nLoop = 1
        While nLoop < .Windows.Count
            If .Windows.Item(nLoop).Type = tpWindow Then
                .Windows.Item(nLoop).Close
            Else
                nLoop = nLoop + 1
            End If
        Wend
    End With
    
    CloseWindow = True
    Exit Function
    
HandleError:
    CloseWindow = False
End Function

Public Function CloseWindows(tpWindow() As VBIDE.vbext_WindowType) As Boolean
'# Closes all windows in IDE matching the types supplied in tpWindow array
On Error GoTo HandleError
    Dim nLoop As Integer, nTypes As Integer
    
    With VBInstance
        nLoop = 1
        While nLoop < .Windows.Count
            For nTypes = 0 To UBound(tpWindow())
                If .Windows.Item(nLoop).Type = tpWindow(nTypes) Then
                    .Windows.Item(nLoop).Close
                    nLoop = nLoop - 1
                    Exit For
                End If
            Next
            nLoop = nLoop + 1
        Wend
    End With
    
    CloseWindows = True
    Exit Function
    
HandleError:
    CloseWindows = False
End Function
