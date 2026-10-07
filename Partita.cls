VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Partita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Giornata As Long
Public Casa As String
Public Ospite As String
Public GoalCasa As Long
Public GoalOspite As Long
Public IdxCasa As Long
Public IdxOspite As Long

Private Sub Class_Initialize()
    GoalCasa = -1
    GoalOspite = -1
    IdxCasa = -1
    IdxOspite = -1
End Sub

Public Function SegnoReale() As String
    If GoalCasa > GoalOspite Then
        SegnoReale = "1"
    ElseIf GoalCasa = GoalOspite Then
        SegnoReale = "X"
    Else
        SegnoReale = "2"
    End If
End Function
