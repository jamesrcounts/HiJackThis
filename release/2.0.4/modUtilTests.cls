VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "modUtilTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub NormalizePath_Expands_EnvVars()
    Dim actual As String
    actual = modUtils.NormalizePath("%WINDIR%\Foo")
    
    Assert.That actual, Iz.EqualTo("C:\WINDOWS\Foo"), "Expand environment variables in a path"
End Sub

Public Sub NormalizePath_Unknown_Token()
    Dim actual As String
    actual = modUtils.NormalizePath("%A%Foo")
    
    Assert.That actual, Iz.EqualTo("%A%Foo"), "Normalize Unknown Token"
End Sub

Public Sub NormalizePath_Empty_Token()
    On Error GoTo ErrorHandler
    Dim actual As String
    actual = modUtils.NormalizePath("% %Foo")
    
ErrorHandler:
    Assert.That Err.Number, Iz.EqualTo(5), "Normalize empty token"
End Sub

Public Sub NormalizePath_With_Adjacent_Percent_Symbols()
    Dim actual As String
    actual = modUtils.NormalizePath("%%Foo")
    
    Assert.That actual, Iz.EqualTo("%%Foo"), "Normalize Adjacent %% in string"
End Sub

Public Sub NormalizePath_With_Unmatched_Percent_Symbol()
    Dim actual As String
    actual = modUtils.NormalizePath("%Foo")
    
    Assert.That actual, Iz.EqualTo("%Foo"), "Normalize Unmatched % in string"
End Sub

Public Sub NormalizePath_When_Already_Normal()
    Dim actual As String
    actual = modUtils.NormalizePath("Foo")
    
    Assert.That actual, Iz.EqualTo("Foo"), "Already Normal"
End Sub

