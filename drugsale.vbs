'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim rst, sql, cnt

Set rst = CreateObject("ADODB.Recordset")
sql = " SELECT top 100 DrugSaleName, PatientName, Qty, (Qty * UnitCost) AS TotalSales FROM DrugSale"
sql = sql & " JOIN Patient ON DrugSale.PatientID=Patient.PatientID"
sql = sql & " JOIN DrugSaleItems ON DrugSale.DrugSaleID=DrugSaleItems.DrugSaleID"

rst.open qrypro.FltQry(sql), conn, 3, 4

With response
    .write "<style>"
    .write "    * {"
    .write "    font-family: 'Roboto', sans-serif;"
    .write "    font-weight: 300;"
    .write "    font-size: 25px;"
    .write "    }"
    .write "    .head, .detail {"
    .write "    border-collapse: collapse;"
    .write "    }"
    .write "    .borDered {"
    .write "    border: solid 1px black;"
    .write "    border-radius: 10px;"
    .write "    }"
    .write "    .head {"
    .write "    background-color: skyblue;"
    .write "    font-size: 20px;"
    .write "    }"
    .write "</style>"
    
    .write "<h3>Total Drug Sales</h3>"

    .write "<table class = 'borDered';>"

    .write "<tr>"
        .write "<th class = 'head';>#</th>"
        .write "<th class = 'head';>Name of Patient</th>"
        .write "<th class = 'head';>Name of Drug</th>"
        .write "<th class = 'head';>Quantity Sold</th>"
        .write "<th class = 'head';>Total Sales</th>"
    .write "</tr>"
    
    cnt = 0
    
    If rst.RecordCount > 0 Then
        rst.movefirst
            Do While Not rst.EOF
            cnt = cnt + 1
    .write "<tr>"
        .write "<td class = 'detail';>" & (cnt) & "</td>"
        .write "<td class = 'detail';>" & rst.fields("PatientName") & "</td>"
        .write "<td class = 'detail';>" & rst.fields("DrugSaleName") & "</td>"
        .write "<td class = 'detail';>" & rst.fields("Qty") & "</td>"
        .write "<td class = 'detail';>" & rst.fields("TotalSales") & "</td>"
    .write "</tr>"
        rst.moveNext
            Loop
    End If
    
rst.Close
    .write "</table>"
End With
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'The purpose of this code is to get PatientNAme, DrugSaleNAme, Quantity of Drug and the TotalSales(Quantity of drug * Unit Price) 
'from DrugSale Table, DrugSaleItems Table and the Patient Table.

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>

