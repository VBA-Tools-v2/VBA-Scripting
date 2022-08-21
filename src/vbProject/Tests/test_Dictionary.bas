Attribute VB_Name = "test_Dictionary"
''
' VBA-Git Annotations
' https://github.com/VBA-Tools-v2/VBA-Git | https://radiuscore.co.nz
'
' @developmentmodule
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' RubberDuck Annotations
' https://rubberduckvba.com/ | https://github.com/rubberduck-vba/Rubberduck/
'
'@testmodule
'@folder Tests
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Private Module

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Type TTest
    Assert As Object                    ' Rubberduck.AssertClass
    Fakes As Object                     ' Rubberduck.FakesProvider
    ScrDictionary As Object             ' Scripting.Dictionary
    VbaDictionary As Dictionary         ' VBA.Dictionary
End Type

Private This As TTest

' ============================================= '
' Test Methods
' ============================================= '

' --------------------------------------------- '
' Speed Tests
' --------------------------------------------- '

'@testmethod Dictionary.SpeedTest
Private Sub speedtest_Add()
    Dim test_Temp As String
    Dim test_Long As Long
    Dim test_StartTime As Date
    Dim test_FinishTime As Date
    Dim test_VbaMS As Double
    Dim test_ScrMs As Double
    
    test_StartTime = VBA.Date + CDate(VBA.Timer / 86400)
    For test_Long = 1 To 10000
        This.ScrDictionary.Add test_Long, test_Long
    Next test_Long
    test_FinishTime = VBA.Date + CDate(VBA.Timer / 86400)
    test_ScrMs = VBA.Round((test_FinishTime - test_StartTime) * 86400 * 1000, 4)
    
    test_StartTime = VBA.Date + CDate(VBA.Timer / 86400)
    For test_Long = 1 To 10000
        This.VbaDictionary.Add test_Long, test_Long
    Next test_Long
    test_FinishTime = VBA.Date + CDate(VBA.Timer / 86400)
    test_VbaMS = VBA.Round((test_FinishTime - test_StartTime) * 86400 * 1000, 4)
    
    This.Assert.Inconclusive "SCR=" & test_ScrMs & "ms | VBA=" & test_VbaMS & "ms | " & VBA.IIf(test_VbaMS > test_ScrMs, "Scripting", "VBA") & " is " & VBA.Round(VBA.IIf(test_VbaMS > test_ScrMs, test_VbaMS / test_ScrMs, test_ScrMs / test_VbaMS), 4) & " times faster."
End Sub

' ============================================= '
' Initialize & Terminate Methods
' ============================================= '

'@TestInitialize
Private Sub TestInitialize()
    ' This method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    ' This method runs after every test in the module.
End Sub

'@ModuleInitialize
Private Sub ModuleInitialize()
    With This
        Set .Assert = CreateObject("Rubberduck.AssertClass")
        Set .Fakes = CreateObject("Rubberduck.FakesProvider")
        Set .ScrDictionary = CreateObject("Scripting.Dictionary")
        Set .VbaDictionary = New Dictionary
    End With
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    With This
        Set .Assert = Nothing
        Set .Fakes = Nothing
        Set .ScrDictionary = Nothing
        Set .VbaDictionary = Nothing
    End With
End Sub
