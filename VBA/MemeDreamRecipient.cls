VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MemeDreamRecipient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class MemeDreamRecipient

Private pName As String
Private pVotes As Scripting.Dictionary


Public Property Get Name() As String
Name = pName
End Property
Public Property Let Name(Value As String)
pName = Value
End Property

Public Property Get Votes() As Scripting.Dictionary
Set Votes = pVotes
End Property
Public Property Let Votes(Value As Scripting.Dictionary)
Set pVotes = Value
End Property


Public Function Create(oRecipient As Outlook.Recipient)
    Me.Votes = CreateObject("Scripting.Dictionary")
    Me.Name = oRecipient.Name
End Function

Public Function Initialize(sName As String)
    
    Me.Votes = CreateObject("Scripting.Dictionary")
    Me.Name = sName
    
End Function
