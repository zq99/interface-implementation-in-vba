VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISubject

Private mcollObservers   As Collection
Private mstrData         As String

Private Sub Class_Initialize()
    Set mcollObservers = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcollObservers = Nothing
End Sub

Private Sub ISubject_addObserver(ByVal observer As IObserver)
    mcollObservers.Add observer
    Debug.Print ">>New Observer added"
End Sub

Private Sub ISubject_removeObserver(ByVal observer As IObserver)
    Dim iCnt As Integer
    For iCnt = 1 To mcollObservers.Count
        If ObjPtr(mcollObservers(iCnt)) = ObjPtr(observer) Then
            Debug.Print ">>Observer @" & iCnt & " has been removed"
            mcollObservers.Remove iCnt
            Exit Sub
        End If
    Next iCnt
End Sub

Private Sub ISubject_notifyObservers()
    If mcollObservers.Count > 0 Then
        Dim iCnt As Integer
        For iCnt = 1 To mcollObservers.Count
            Debug.Print ">>Notifying Observer of message"
            mcollObservers(iCnt).update (mstrData)
        Next iCnt
    End If
End Sub

Private Function ISubject_countObservers() As Integer
    ISubject_countObservers = mcollObservers.Count
End Function

Public Sub addObserver(ByVal observer As IObserver)
    Call ISubject_addObserver(observer)
End Sub

Public Sub removeObserver(ByVal observer As IObserver)
    Call ISubject_removeObserver(observer)
End Sub

Public Function countObservers() As Integer
    countObservers = ISubject_countObservers
End Function

Public Function SendMessage(ByVal sData As String)
    mstrData = sData
    ISubject_notifyObservers
End Function
