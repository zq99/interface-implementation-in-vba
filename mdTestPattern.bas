Attribute VB_Name = "mdTestPattern"
Option Explicit

Public Sub TestObserverPattern()

    Dim oOb1        As New clsObserver
    Dim oOb2        As New clsObserver
    Dim oSubject    As New clsSubject
    
    Debug.Print "1# Current number of observers is " & oSubject.countObservers
    oSubject.addObserver oOb1
    oSubject.addObserver oOb2
    Debug.Print "2# Current number of observers is " & oSubject.countObservers
    
    Debug.Print "3# Message Status of First object = " & oOb1.getMessageStatus
    Debug.Print "4# Message Status of Second object = " & oOb2.getMessageStatus
    oSubject.SendMessage "Hello"
    Debug.Print "5# Message Status of First object = " & oOb1.getMessageStatus
    Debug.Print "6# Message Status of Second object = " & oOb2.getMessageStatus
    
    oSubject.removeObserver oOb1
    Debug.Print "7# Current number of observers is " & oSubject.countObservers
    oSubject.SendMessage "Goodbye"
    Debug.Print "8# Message Status of First object = " & oOb1.getMessageStatus
    Debug.Print "9# Message Status of Second object = " & oOb2.getMessageStatus
    
    
    Set oOb1 = Nothing
    Set oOb2 = Nothing
    Set oSubject = Nothing

End Sub
