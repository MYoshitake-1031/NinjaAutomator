Attribute VB_Name = "PV"
Option Explicit

Private Tok As String
Private cnt As Long
Private Lat As Double, Lon As Double, Node As String
Private MaxData As Long


Sub PV()
    
    MaxData = Range("E2").End(xlDown).Row
    cnt = 2
    
    Do
        Node = Cells(cnt, 5).Value
        Lat = Cells(cnt, 6).Value
        Lon = Cells(cnt, 7).Value
        
        If Len(Node) <> 0 And Len(Lat) <> 0 And Len(Lon) <> 0 Then
                    
            '**
            '* URLとパラメータの設定
            '
            '* api/data/wind?&lat=56&lon=-3&date_from=2014-01-01&date_to=2014-02-28&capacity=1&dataset=merra2&height=100&turbine=Vestas+V80+2000&format=json
            Dim date_from As String, date_to As String, Dataset As String, Capacity As Double, loss As Double, track As Double, tilt As Double, azimuth As Double, aggregat As String
            Dim year As Long
            Dim Url As String, Par As String
            Dim param As Object
            
            Tok = Range("J2").Value
            
            year = Cells(2, 2).Value
            date_from = CStr(year) + "-01-01"
            date_to = CStr(year) + "-12-31"
            
            Dataset = CStr(Sheet2.Cells(3, 2).Value)
            Capacity = CStr(Sheet2.Cells(4, 2).Value)
            loss = Sheet2.Cells(5, 2).Value
            track = CStr(Sheet2.Cells(6, 2).Value)
            tilt = CStr(Sheet2.Cells(7, 2).Value)
            azimuth = CStr(Sheet2.Cells(8, 2).Value)
            
            aggregat = Sheet2.Cells(10, 2).Value
            
            Url = "https://www.renewables.ninja/api/data/pv?"
            
            Par = "lat=" & Lat & "&lon=" & Lon & "&date_from=" & date_from & "&date_to=" & date_to & _
                "&dataset=" & Dataset & "&capacity=" & Capacity & "&system_loss=" & loss / 100 & _
                "&tracking=" & track & "&tilt=" & tilt & "&azim=" & azimuth & "&format=csv&header=false"
            
            If (aggregat <> "hour") Then Par = Par & "&mean=" & aggregat
            
            
            '**
            '* APIからデータを取得
            '* @Input:URL,param
            '*
            Dim datFile As String
            Dim httpObject As Object, text As String
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
            
                    text = .responseText
                    datFile = ActiveWorkbook.Path & "\pv\" & Node & "-" & CStr(Lat) & "-" & CStr(Lon) & ".csv"
    
                    Open datFile For Output As #1
                        Print #1, text
                    Close #1
                    Debug.Print "Moving", Node, cnt
                    cnt = cnt + 1
                ElseIf .Status = 429 Then
                    Debug.Print "Waiting", Node, cnt, Now()
                    Application.Wait [Now() + "00:02:00"]
                Else
                    MsgBox "Error: " & .Status & .statusText
                    Exit Do
                End If
            End With
        Else
            MsgBox "Error: Cannot read NODE or LAT/LON"
            Exit Do
        End If
    Loop

End Sub

