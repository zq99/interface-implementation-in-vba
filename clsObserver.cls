VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsObserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IObserver

Private mstrData As String

Private Sub Class_Initialize()
    mstrData = "{Empty}"
End Sub

Private Sub IObserver_update(ByVal strData As String)
    mstrData = strData
End Sub

Public Sub update(ByVal strData As String)
    Call IObserver_update(strData)
End Sub

Public Function getMessageStatus()
    getMessageStatus = mstrData
End Function
