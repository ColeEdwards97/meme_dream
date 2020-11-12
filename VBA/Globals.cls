VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Globals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Property Get Version() As String
Version = "4.0"
End Property

Property Get ArchiveDir() As String
ArchiveDir = "C:\Users\EDWARDCO\Desktop\briefcase\rotational program\meme dream\_archive_\"
End Property

Property Get TemplateDir() As String
TemplateDir = "C:\Users\EDWARDCO\Desktop\briefcase\rotational program\meme dream\"
End Property

Property Get TemplateName() As String
TemplateName = "MemeDream.oft"
End Property

Property Get IncomingFolderName() As String
IncomingFolderName = "Meme Dream"
End Property

Property Get OutgoingFolderName() As String
OutgoingFolderName = "Meme Dream"
End Property

Property Get TrackerDir() As String
TrackerDir = "C:\Users\EDWARDCO\Desktop\briefcase\rotational program\meme dream\"
End Property

Property Get TrackerName() As String
TrackerName = "tracker.xlsx"
End Property

Property Get ReportDir() As String
ReportDir = "C:\Users\EDWARDCO\Desktop\briefcase\rotational program\meme dream\"
End Property

Property Get ReportName() As String
ReportName = "performance report.xlsx"
End Property

Property Get VotingOptions() As String
VotingOptions = "10 - MEMETACULAR;9 - lol EPIC;8 - big Chortle;7 - nice!;6 - good effort;5 - no opinion;4 - ehh;3 - nah;2 - didn't like;1 - BOO!;0 - ABSOLUTELY ABISMAL >:(;i don't get it"
End Property
