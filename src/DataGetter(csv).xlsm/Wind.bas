Attribute VB_Name = "Wind"
Option Explicit

Private Tok As String
Private cnt As Long
Private Lat As Double, Lon As Double, Node As String
Private MaxData As Long


Sub Wind()
    
    MaxData = Range("E2").End(xlDown).Row
    cnt = 2
    Call Looper

End Sub


Private Sub Looper()
    
    
    Node = Cells(cnt, 5).Value
    Lat = Cells(cnt, 6).Value
    Lon = Cells(cnt, 7).Value

    Call WindDataGetter


End Sub

Private Sub WindDataGetter()

    '**
    '* URLとパラメータの設定
    '
    '* api/data/wind?&lat=56&lon=-3&date_from=2014-01-01&date_to=2014-02-28&capacity=1&dataset=merra2&height=100&turbine=Vestas+V80+2000&format=json
    Dim date_from As String, date_to As String, Dataset As String, Capacity As Double, Height As Double, turbine As String, aggregat As String
    Dim year As Long
    Dim Url As String, Par As String
    Dim param As Object
    
    Tok = Range("J2").Value
        
    year = Sheet1.Cells(2, 2).Value
    date_from = CStr(year) + "-01-01"
    date_to = CStr(year) + "-12-31"
    
    Dataset = CStr(Sheet1.Cells(3, 2).Value)
    Capacity = CStr(Sheet1.Cells(4, 2).Value)
    Height = CStr(Sheet1.Cells(5, 2).Value)
    turbine = CStr(Sheet1.Cells(6, 2).Value)
    aggregat = Sheet1.Cells(8, 2).Value

    Url = "https://www.renewables.ninja/api/data/wind?"
    
    Par = "lat=" & Lat & "&lon=" & Lon & "&date_from=" & date_from & "&date_to=" & date_to & _
          "&dataset=" & Dataset & "&capacity=" & Capacity & "&height=" & Height & _
          "&turbine=" & turbine & "&format=csv&header=false"
    
    If (aggregat <> "hour") Then Par = Par & "&mean=" & aggregat
   
    

    '**
    '* APIからデータを取得
    '* @Input:URL,param
    '* @Output：機器の出力(温度・時間など)
    Dim WindData As Variant
    WindData = KickAPI(Url, Par, param)
    
    
End Sub
Public Function KickAPI( _
    ByVal Url As String, _
    ByVal Par As String, _
    Optional ByVal param As Object) As Variant
    
    Dim datFile As String
    Dim httpObject As Object, textchecker As String
    Set httpObject = CreateObject("MSXML2.XMLHTTP")

    With httpObject
    
        .Open "GET", Url & Par, True
        .setRequestHeader "Authorization", "Token " & Tok
        .send (Par)
        

        ' wait until data has been downloaded
        Do While 1
            If .readyState = 4 Then Exit Do
            DoEvents
        Loop
              
        ' check if we were successful
        If .Status = 200 Then
            
            textchecker = .ResponseText
            datFile = ActiveWorkbook.Path & "\wind\" & Node & "-" & CStr(Lat) & "-" & CStr(Lon) & ".csv"
            Debug.Print Node

            Open datFile For Output As #1
                Print #1, textchecker
            Close #1
'           Debug.Print textchecker
            Debug.Print .Status
            
            If cnt < MaxData Then
                cnt = cnt + 1
                Call Looper
            Else
                End
            End If
        Else
            
            Debug.Print "Waiting", Node, Now()
            Application.Wait [Now() + "00:02:00"]
            Call Looper

        End If
        
    End With

End Function
