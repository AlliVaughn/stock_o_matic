'Final HW 
Sub stock_counter_looper()


    'Easy option:
    'Add run on each sheet capability
    'Added ws. in the places I needed to specify that we mean this specific worksheet.
    'First, Set variable for Worksheet and Wb and initialize
    Dim wb As Workbook
    Dim ws As Worksheet

    For Each ws In Worksheets
    'Found this nifty initialization here: http://codevba.com/excel/activate_worksheet.htm#.XUHbq5NKjyw
    'crossing fingers that it works.  Either way, it sounds cool!
    ' * Worksheet: ACTIVATE!*
    ' ps It works!
    ws.Activate

    'Easy Option Get the Ticker and Total volume to add up and print in two columns on a sheet
    'set variable types
    Dim ticker As String
    Dim total_volume As Double
    Dim Summary_Table_Row As Integer
   
    'Initialize just in case! (As , Manuel says!  :-) )
    'Init these variables as anything but global did not work for me.
    total_volume = 0
    'store the location of each row
    Summary_Table_Row = 2
    'determine the last row & Needed to scope this as ws.Cells, which makes sense
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row 
    'For Each ws in worksheets << This position did not work. 
    'loop through rows 
        For i = 2 To LastRow
        
            'Orginially had this here...No Bueno. Needs to be inside the Conditional
            ' ticker = Cells(i, 1).Value

            'checking if the value of cells are different, like in the State credit card table exercise,  <> is NOT equal in VBS
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                'setting the value of that ticker
                ticker = Cells(i, 1).Value

                'Setting the total_volume per ticker - i,7 is Column G's Value per row,
                'since we are adding each together for that total
                total_volume = total_volume + Cells(i, 7).Value
                
                'These two will show up next to one another bc they are J & K on the table as 
                'J: Value of Ticker Name  
                'K: Value of total volume
                'Adding ws. scope here.  It'll probably not work. UPDATE:  It DID!
                ws.Range("J" & Summary_Table_Row).Value = ticker
                ws.Range("K" & Summary_Table_Row).Value = total_volume
                'Make it neat & pretty
                ws.Range("J1").Value = "Ticker"
                ws.Range("k1").Value = "Total Volume"
                ws.Range("K:K").NumberFormat = "$#,##0.00"
                ' I left this bc I used it to make sure it was the ws I was on. 
                ' It's informative anyhow.
                ws.Range("M2").Value = "This is stock data for Year " + ws.Name
            
                'Build it: increment that summary stuff by 1   
                Summary_Table_Row = Summary_Table_Row + 1
                
                'we now need to reset the Total so the next one can go!
                'Otherwise, it's wrong.
                total_volume = 0

            Else
                'if the conditions above were not met just do the calculation
                total_volume = total_volume + Cells(i, 7).Value
            End If
        Next i
    'THIS bit me in the you-know-what bc I forgot to close it at first.
    'Kinda need to end and move along to the next sheet! 
    '***When's Python again?*** I'm about to tag VBS range as style ="Hate"
    Next ws
End Sub
