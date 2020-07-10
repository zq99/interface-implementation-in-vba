VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub addObserver(ByVal observer As IObserver)
End Sub

Public Sub removeObserver(ByVal observer As IObserver)
End Sub

Public Sub notifyObservers()
End Sub

Public Function countObservers() As Integer
End Function
