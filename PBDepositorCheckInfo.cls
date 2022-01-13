VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBDepositorCheckInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public result As String
Public resultMessage As String
Public accountNumber As String
Public bankCode As String
Public accountName As String
Public identityNumType As String
Public identityNum As String
Public checkDate As String

Public Sub fromDictionary(dic As Variant)
    If IsNull(dic.Item("result")) = False Then
        result = dic.Item("result")
    Else
        result = ""
    End If
    
    If IsNull(dic.Item("resultMessage")) = False Then
        resultMessage = dic.Item("resultMessage")
    Else
        resultMessage = ""
    End If
    
    If IsNull(dic.Item("accountNumber")) = False Then
        accountNumber = dic.Item("accountNumber")
    Else
        accountNumber = ""
    End If
    
    If IsNull(dic.Item("bankCode")) = False Then
        bankCode = dic.Item("bankCode")
    Else
        bankCode = ""
    End If
    
    If IsNull(dic.Item("accountName")) = False Then
        accountName = dic.Item("accountName")
    Else
        accountName = ""
    End If
    
        If IsNull(dic.Item("identityNumType")) = False Then
        identityNumType = dic.Item("identityNumType")
    Else
        identityNumType = ""
    End If
    
        If IsNull(dic.Item("identityNum")) = False Then
        identityNum = dic.Item("identityNum")
    Else
        identityNum = ""
    End If
    
    If IsNull(dic.Item("checkDate")) = False Then
        checkDate = dic.Item("checkDate")
    Else
        checkDate = ""
    End If
End Sub


