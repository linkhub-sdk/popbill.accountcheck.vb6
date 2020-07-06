VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBAccountCheckInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public resultCode As String
Public resultMessage As String
Public accountNumber As String
Public bankCode As String
Public accountName As String
Public checkDate As String

Public Sub fromDictionary(dic As Variant)
    If IsNull(dic.Item("resultCode")) = False Then
        resultCode = dic.Item("resultCode")
    Else
        resultCode = ""
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
    
    If IsNull(dic.Item("checkDate")) = False Then
        checkDate = dic.Item("checkDate")
    Else
        checkDate = ""
    End If
End Sub

