'This loops through HistoTrac databses to fetch the same query for each record

Sub HistTracConnt()

    'Initializes variables
    Set cnn = New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim ConnectionString As String
    Dim StrQuery1 As String
    Dim StrQuery2 As String
    Dim StrQuery3 As String
    Dim StrQuery4 As String
    Dim StrQuery5 As String
    Dim StrQuery6 As String
    Dim StrQuery7 As String
    Dim StrQuery8 As String
    Dim StrQuery9 As String
    Dim StrQuery10 As String
    Dim StrQuery11 As String
    Dim StrQuery12 As String
    Dim StrQuery13 As String
    Dim StrQuery14 As String
    
    'Define cells to integrate into query
       
    'MRN = Sheet1.Cells(2, 2)
    'DATE1 = Sheet1.Cells(1, 2)
    'DATE2 = Sheet1.Cells(1, 4)
    'TESTCODE = Sheet1.Cells(2, 2)
    'PTCATEGORY = Sheet1.Cells(2, 4)
    
    'clear worksheet before import
    Worksheets("Sheet1").Range("C2:O1000000").Clear
       
        
     'Defining Path for Connection
     ConnectionString = "Driver={SQL Server Native Client 11.0};server=BSWTILHISTP01;database=HistoTracPROD;Trusted_Connection=yes;"
     'ConnectionString = "driver={SQL Server};server=BHDATILHIST01;database=HistoTracPROD;SECURITY=sspi;"

    'Opens connection to the database
    cnn.Open ConnectionString
    'Timeout error in seconds for executing the entire query; this will run for 15 minutes before VBA timesout, but your database might timeout before this value
    cnn.CommandTimeout = 30
           
    For X = 1 To 986
    
    MRN = Cells(X + 1, 1)
    'FNAME = Cells(X + 1, 2)
    ROWC = Cells(Rows.Count, 3).End(xlUp).Row
    OUTPUTRW = ROWC + 1
    OUTPUTCLM = 3
    'SAMPIDN = SAMPID * 1
    
     Set rst = Nothing
     
     
    'This is your actual MS SQL query that you need to run; you should check this query first using a more robust SQL editor (such as HeidiSQL) to ensure your query is valid
    StrQuery1 = "SELECT        Patient.HospitalID, Patient.lastnm, Patient.firstnm, Patient.categoryCd, Sample.SampleDt, Sample.SampleNbr, Test.TestDt, Test.TestTypeCd, Test.SpecificityTxt, Test.PRAResultCd, Test.ReportableCommentsTxt"
    StrQuery2 = StrQuery1 & " FROM            Patient INNER JOIN"
    StrQuery3 = StrQuery2 & " Sample ON Patient.PatientId = Sample.PatientId INNER JOIN"
    StrQuery4 = StrQuery3 & " Test ON Sample.SampleID = Test.SampleId"
    'StrQuery5 = StrQuery4 & " "
    'StrQuery6 = StrQuery5 & " "
    StrQuery5 = StrQuery4 & " WHERE        (((Patient.HospitalID = N'" & MRN & "')) AND (Test.TestTypeCd LIKE '%DSA') )"
    'StrQuery7 = StrQuery6 & " WHERE        ((Patient.lastnm = N'" & LNAME & "')) AND (Test.TestTypeCd = 'HLATY')  AND (Patient.firstnm = N'" & FNAME & "')"
    
    
    'Performs the actual query
    rst.Open StrQuery5, cnn
    
    'Dumps all the results from the StrQuery into cell A2 of the first sheet in the active workbook
    Sheets("Sheet1").Cells(OUTPUTRW, OUTPUTCLM).CopyFromRecordset rst
    
          
        
    Next X
    

End Sub

