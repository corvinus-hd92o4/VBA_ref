Attribute VB_Name = "Module1"
Public Sub saveAttachtoDisk(itm As Outlook.MailItem)
    
    Dim objAtt As Outlook.Attachment
    Dim saveFolder As String
        saveFolder = "C:\Users\extboda\Documents\"
    Dim dateFormat As String
        dateFormat = Format(itm.ReceivedTime, "yyyy.mm.dd.hh.mm.ss")
    
    
    Dim olFileType As String
    Dim xExcelApp As Excel.Application
    
    'If itm.Attachments.Count = 0 Then
        'Call NoAtt
        'Exit Sub
        
   ' Else'
   'For Each objAtt In itm.Attachments
           ' objAtt.SaveAsFile saveFolder & "\" & "BSH_Osszevonas_" & dateFormat & "_GW_Nyers.xlsx"
           'Next
   ' End If
    
    
    If itm.Attachments.Count > 0 Then
    If itm.Attachments.Count = 1 Then
            For Each objAtt In itm.Attachments
            olFileType = LCase$(Right$(objAtt.FileName, 4))
            
                    If olFileType = "xlsx" Then
                        
                        objAtt.SaveAsFile saveFolder & "\" & "BSH_Osszevonas_" & dateFormat & "_GW_Nyers.xlsx"
                        
                        Set xExcelApp = CreateObject("Excel.Application")
                        Set Workbook = Workbooks.Open("C:\Users\extboda\Documents\" & "BSH_Osszevonas_" & dateFormat & "_GW_Nyers.xlsx")
                        Call EAKER
                        xExcelApp.DisplayAlerts = False
                        Workbook.Close SaveChanges:=True
                        xExcelApp.DisplayAlerts = False
                        xExcelApp.Quit
                    
                    Else
                    Call NoAtt
                    Exit Sub
                    End If
            Next
            
    End If
    
        If itm.Attachments.Count > 1 Then
        For Each objAtt In itm.Attachments
                olFileType = LCase$(Right$(objAtt.FileName, 4))
                
                        If olFileType = "xlsx" Then
                            
                            objAtt.SaveAsFile saveFolder & "\" & "BSH_Osszevonas_" & dateFormat & "_GW_Nyers.xlsx"
                            'Dim xExcelApp As Excel.Application
                            Set xExcelApp = CreateObject("Excel.Application")
                            Set Workbook = Workbooks.Open("C:\Users\extboda\Documents\" & "BSH_Osszevonas_" & dateFormat & "_GW_Nyers.xlsx")
                            Call EAKER
                            xExcelApp.DisplayAlerts = False
                            Workbook.Close SaveChanges:=True
                            xExcelApp.DisplayAlerts = False
                            xExcelApp.Quit
                        
                        End If
                Next
                
        
        End If
    Else
    Call NoAtt
    Exit Sub
    End If

    

        
End Sub

Sub EAKER()

Dim myFile As String

'j As Integer,
'n As Integer,

Dim temp1 As String
Dim k As Integer
Dim i As Integer
Dim LR As Integer
Dim LC As Integer
Dim l As Integer
Dim aux_str As String
Dim all_str As Variant

'l az új munkalapon a sor száma, ahova a jó sorokat másolni fogjuk
l = 1




'új munkalap letrehozasa, ahova masolni fogunk
Sheets.Add.Name = "Sorted"

Sheets(2).Select

'D szerinti ABC sorrend
With ActiveSheet.Sort
    .SortFields.Add Key:=Range("D1"), Order:=xlAscending
    .SetRange Range("A1", Range("H1").End(xlDown))
    .Header = xlYes
    .Apply
End With

 
LR = Cells(Rows.Count, 1).End(xlUp).Row
LC = Cells(1, Columns.Count).End(xlToLeft).Column

'referencia szam nagyobb-e mint 10, ha igen, atmasolom
   For k = 2 To LR
       For i = 1 To LC
           If i = 6 Then
                If Len(Cells(k, i).Value) > 10 Then
                    Rows(k).Copy Destination:=Worksheets("Sorted").Rows(l)
                    l = l + 1
                End If
           End If
       Next i
   Next k
   

         
Sheets("Sorted").Select
   

LR = Cells(Rows.Count, 1).End(xlUp).Row
LC = Cells(1, Columns.Count).End(xlToLeft).Column
    For k = 1 To LR
        For i = 1 To LC
        If i = 4 Then
        'csak az elso 3 betut tartom meg
            Cells(k, i).Value = Left(Cells(k, i), 3)
            
        End If
        
        If i = 6 Then
        '/ek kicserelese , re
            Cells(k, i).Value = Replace(Cells(k, i), "/", ",")
        End If
        Next i
    Next k
    
    
aux_str = ""

LR = Cells(Rows.Count, 1).End(xlUp).Row
LC = Cells(1, Columns.Count).End(xlToLeft).Column

'sorszámozás

    For k = 1 To LR
        For i = 1 To LC
        If i = 4 Then
            If aux_str = "" Then
             aux_str = Cells(k, i).Value
             Cells(k, i).Offset(0, 3).Value = 1
        Else
            If aux_str = Cells(k, i).Value Then
                Cells(k, i).Offset(0, 3).Value = Cells(k, i).Offset(-1, 3).Value + 1
            Else
                aux_str = Cells(k, i).Value
                Cells(k, i).Offset(0, 3).Value = 1
            End If
        
        End If
        End If
        Next i
    Next k
    

all_str = ""





LR = Cells(Rows.Count, 1).End(xlUp).Row
LC = Cells(1, Columns.Count).End(xlToLeft).Column


If Not IsEmpty(Cells(1, 1)) Then

myFile = "E:" & "\" & "OSSZEVONAS_" & Format(Now(), "YYYYMMDDHHMMSS" & ".txt")

Open myFile For Output As #1

    For k = 1 To LR
        'leveszem az utolso pontot a datumrol
        temp1 = Left(Cells(k, 1), Len(Cells(k, 1)) - 1)
        
        all_str = all_str & temp1 & "," & Cells(k, 7).Value & "," & Cells(k, 6).Value & "|"
    Next k
    
'enterek kiszedese
all_str = Replace(Replace(all_str, Chr(10), ""), Chr(13), "")
all_str = Replace(Replace(all_str, Chr(32), ""), Chr(13), "")



'uccso karakter levagasa,szokozok
all_str = Left(all_str, Len(all_str) - 1)

    
Print #1, all_str;


 
 
Close #1
End If


'MsgBox ("Done")

End Sub
Sub NoAtt()
 Dim OutlookApp As Outlook.Application
  Dim OutlookMail As Outlook.MailItem

  Set OutlookApp = New Outlook.Application
  Set OutlookMail = OutlookApp.CreateItem(olMailItem)
  
  With OutlookMail
    .BodyFormat = olFormatHTML
    .Display
    .HTMLBody = "Kedves Címzett!" & "<br>" & "<br>" & "Egy olyan leveled érkezett, amibõl hiányzik a szükséges csatolmány. Kérlek járj utána."
    '.HTMLBody
    'last .HTMLBody includes signature from the outlook.
    '<br> includes line breaks b/w two lines
    .To = "mihaly.boda-ext@bshg.com"
    .CC = ""
    .BCC = ""
    .Subject = "EAKER NO ATTACHMENT"
    .Send
  End With

 End Sub
