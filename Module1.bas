Attribute VB_Name = "Module1"
Option Explicit
Sub FirstSemester()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    With ws.Range("K4:S9")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Naming the cells
    ws.Range("K5").Value = "Total Score"
    ws.Range("K6").Value = "Percentage"
    ws.Range("K7").Value = "Grade"
    ws.Range("K8").Value = "12-Point"
    ws.Range("K9").Value = "GPA"
    ws.Range("S4").Value = "Total Score"
    
    ' Naming the Columns
    ws.Range("L4").Value = "Classroom Dynamics"
    ws.Range("M4").Value = "Human Development"
    ws.Range("N4").Value = "Use of Computers"
    ws.Range("O4").Value = "Introduction to Developmental Disabilities"
    ws.Range("P4").Value = "Behavior Managemnt I"
    ws.Range("Q4").Value = "College Writing Skills"
    ws.Range("R4").Value = "Professional Ethics"
    
    
    ' 1. Calculate Total Scores (Raw Points Earned)
    ws.Range("L5").Formula = "=SUM(B5:B42)"
    ws.Range("M5").Formula = "=SUM(C5:C42)"
    ws.Range("N5").Formula = "=SUM(D5:D42)"
    ws.Range("O5").Formula = "=SUM(E5:E42)"
    ws.Range("P5").Formula = "=SUM(F5:F42)"
    ws.Range("Q5").Formula = "=SUM(G5:G42)"
    ws.Range("R5").Formula = "=SUM(H5:H42)"
    
    
    ws.Range("B43").Formula = "=SUM(B5:B42)"
    ws.Range("C43").Formula = "=SUM(C5:C42)"
    ws.Range("D43").Formula = "=SUM(D5:D42)"
    ws.Range("E43").Formula = "=SUM(E5:E42)"
    ws.Range("F43").Formula = "=SUM(F5:F42)"
    ws.Range("G43").Formula = "=SUM(G5:G42)"
    ws.Range("H43").Formula = "=SUM(H5:H42)"
    
    ' 2. Calculate Percentages (Points Earned / Points Possible)
    ws.Range("L6").Formula = "=L5/100"
    ws.Range("M6").Formula = "=M5/100"
    ws.Range("N6").Formula = "=N5/6500"
    ws.Range("O6").Formula = "=O5/100"
    ws.Range("P6").Formula = "=P5/100"
    ws.Range("Q6").Formula = "=Q5/100"
    ws.Range("R6").Formula = "=R5/100"
    
    ws.Range("B44").Formula = "=B43/100"
    ws.Range("C44").Formula = "=C43/100"
    ws.Range("D44").Formula = "=D43/6500"
    ws.Range("E44").Formula = "=E43/100"
    ws.Range("F44").Formula = "=F43/100"
    ws.Range("G44").Formula = "=G43/100"
    ws.Range("H44").Formula = "=H43/100"

    ' 3. Grade Formula
    Dim gradeFormula As String
    gradeFormula = "=IF(L6>=90%,""A+"",IF(L6>=85%,""A"",IF(L6>=80%,""A-"",IF(L6>=77%,""B+"",IF(L6>=73%,""B"",IF(L6>=70%,""B-"",IF(L6>=67%,""C+"",IF(L6>=63%,""C"",IF(L6>=60%,""C-"",IF(L6>=57%,""D+"",IF(L6>=53%,""D"",IF(L6>=50%,""D-"",""F""))))))))))))"

    ws.Range("L7").Formula = gradeFormula
    ws.Range("M7").Formula = Replace(gradeFormula, "L6", "M6")
    ws.Range("N7").Formula = Replace(gradeFormula, "L6", "N6")
    ws.Range("O7").Formula = Replace(gradeFormula, "L6", "O6")
    ws.Range("P7").Formula = Replace(gradeFormula, "L6", "P6")
    ws.Range("Q7").Formula = Replace(gradeFormula, "L6", "Q6")
    ws.Range("R7").Formula = Replace(gradeFormula, "L6", "R6")
    
    ' 4. Point Formula
    Dim PointFormula As String
    PointFormula = "=IF(L6>=90%,12,IF(L6>=85%,11,IF(L6>=80%,10,IF(L6>=77%,9,IF(L6>=73%,8,IF(L6>=70%,7,IF(L6>=67%,6,IF(L6>=63%,5,IF(L6>=60%,4,IF(L6>=57%,3,IF(L6>=53%,2,IF(L6>=50%,1,0))))))))))))"

    ws.Range("L8").Formula = PointFormula
    ws.Range("M8").Formula = Replace(PointFormula, "L6", "M6")
    ws.Range("N8").Formula = Replace(PointFormula, "L6", "N6")
    ws.Range("O8").Formula = Replace(PointFormula, "L6", "O6")
    ws.Range("P8").Formula = Replace(PointFormula, "L6", "P6")
    ws.Range("Q8").Formula = Replace(PointFormula, "L6", "Q6")
    ws.Range("R8").Formula = Replace(PointFormula, "L6", "R6")
    
    ' 4. GPA Formula
    Dim gpaFormula As String
    gpaFormula = "=IF(L6>=90%,4,IF(L6>=85%,3.9,IF(L6>=80%,3.7,IF(L6>=77%,3.3,IF(L6>=73%,3,IF(L6>=70%,2.7,IF(L6>=67%,2.3,IF(L6>=63%,2,IF(L6>=60%,1.7,IF(L6>=57%,1.3,IF(L6>=53%,1,IF(L6>=50%,0.7,0))))))))))))"

    ws.Range("L9").Formula = gpaFormula
    ws.Range("M9").Formula = Replace(gpaFormula, "L6", "M6")
    ws.Range("N9").Formula = Replace(gpaFormula, "L6", "N6")
    ws.Range("O9").Formula = Replace(gpaFormula, "L6", "O6")
    ws.Range("P9").Formula = Replace(gpaFormula, "L6", "P6")
    ws.Range("Q9").Formula = Replace(gpaFormula, "L6", "Q6")
    ws.Range("R9").Formula = Replace(gpaFormula, "L6", "R6")
    
    Dim totalpoint As String
    totalpoint = "=SUM(L8:R8)"
    ws.Range("S8").Formula = totalpoint
    
    Dim totalgpa As String
    totalgpa = "=SUM(L9:R9)"
    ws.Range("S9").Formula = totalgpa
    
End Sub

Sub SecondSemester()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    With ws.Range("K54:R59")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Naming the cells
    ws.Range("K55").Value = "Total Score"
    ws.Range("K56").Value = "Percentage"
    ws.Range("K57").Value = "Grade"
    ws.Range("K58").Value = "12-Point"
    ws.Range("K59").Value = "GPA"
    ws.Range("R54").Value = "Total Score"
    
    ' Naming the Columns
    ws.Range("L54").Value = "Group Dynamics"
    ws.Range("M54").Value = "Program Development"
    ws.Range("N54").Value = "Abnormal Psychology"
    ws.Range("O54").Value = "Behavior Management II"
    ws.Range("P54").Value = "Social Inclusion Strategies – QAM, CFSA"
    ws.Range("Q54").Value = "Placement 1"
    
    ' 1. Calculate Total Scores (Raw Points Earned)
    ws.Range("L55").Formula = "=SUM(B55:B75)"
    ws.Range("M55").Formula = "=SUM(C55:C75)"
    ws.Range("N55").Formula = "=SUM(D55:D75)"
    ws.Range("O55").Formula = "=SUM(E55:E75)"
    ws.Range("P55").Formula = "=SUM(F55:F75)"
    ws.Range("Q55").Formula = "=SUM(G55:G75)"
   
    
    
    ws.Range("B76").Formula = "=SUM(B55:B75)"
    ws.Range("C76").Formula = "=SUM(C55:C75)"
    ws.Range("D76").Formula = "=SUM(D55:D75)"
    ws.Range("E76").Formula = "=SUM(E55:E75)"
    ws.Range("F76").Formula = "=SUM(F55:F75)"
    ws.Range("G76").Formula = "=SUM(G55:G75)"
    
    
    ' 2. Calculate Percentages (Points Earned / Points Possible)
    ws.Range("L56").Formula = "=L55/100"
    ws.Range("M56").Formula = "=M55/100"
    ws.Range("N56").Formula = "=N55/100"
    ws.Range("O56").Formula = "=O55/100"
    ws.Range("P56").Formula = "=P55/100"
    ws.Range("Q56").Formula = "=Q55/100"
    
    
    ws.Range("B77").Formula = "=B76/100"
    ws.Range("C77").Formula = "=C76/100"
    ws.Range("D77").Formula = "=D76/100"
    ws.Range("E77").Formula = "=E76/100"
    ws.Range("F77").Formula = "=F76/100"
    ws.Range("G77").Formula = "=G76/100"
    

    ' 3. Grade Formula
    Dim gradeFormula As String
    gradeFormula = "=IF(L56>=90%,""A+"",IF(L56>=85%,""A"",IF(L56>=80%,""A-"",IF(L56>=77%,""B+"",IF(L56>=73%,""B"",IF(L56>=70%,""B-"",IF(L56>=67%,""C+"",IF(L56>=63%,""C"",IF(L56>=60%,""C-"",IF(L56>=57%,""D+"",IF(L56>=53%,""D"",IF(L56>=50%,""D-"",""F""))))))))))))"

    ws.Range("L57").Formula = gradeFormula
    ws.Range("M57").Formula = Replace(gradeFormula, "L56", "M56")
    ws.Range("N57").Formula = Replace(gradeFormula, "L56", "N56")
    ws.Range("O57").Formula = Replace(gradeFormula, "L56", "O56")
    ws.Range("P57").Formula = Replace(gradeFormula, "L56", "P56")
    ws.Range("Q57").Formula = Replace(gradeFormula, "L56", "Q56")
    
    
    ' 4. Point Formula
    Dim PointFormula As String
    PointFormula = "=IF(L56>=90%,12,IF(L56>=85%,11,IF(L56>=80%,10,IF(L56>=77%,9,IF(L56>=73%,8,IF(L56>=70%,7,IF(L56>=67%,6,IF(L56>=63%,5,IF(L56>=60%,4,IF(L56>=57%,3,IF(L56>=53%,2,IF(L56>=50%,1,0))))))))))))"

    ws.Range("L58").Formula = PointFormula
    ws.Range("M58").Formula = Replace(PointFormula, "L56", "M56")
    ws.Range("N58").Formula = Replace(PointFormula, "L56", "N56")
    ws.Range("O58").Formula = Replace(PointFormula, "L56", "O56")
    ws.Range("P58").Formula = Replace(PointFormula, "L56", "P56")
    ws.Range("Q58").Formula = Replace(PointFormula, "L56", "Q56")
    
    
    ' 4. GPA Formula
    Dim gpaFormula As String
    gpaFormula = "=IF(L56>=90%,4,IF(L56>=85%,3.9,IF(L56>=80%,3.7,IF(L56>=77%,3.3,IF(L56>=73%,3,IF(L56>=70%,2.7,IF(L56>=67%,2.3,IF(L56>=63%,2,IF(L56>=60%,1.7,IF(L56>=57%,1.3,IF(L56>=53%,1,IF(L56>=50%,0.7,0))))))))))))"

    ws.Range("L59").Formula = gpaFormula
    ws.Range("M59").Formula = Replace(gpaFormula, "L56", "M56")
    ws.Range("N59").Formula = Replace(gpaFormula, "L56", "N56")
    ws.Range("O59").Formula = Replace(gpaFormula, "L56", "O56")
    ws.Range("P59").Formula = Replace(gpaFormula, "L56", "P56")
    ws.Range("Q59").Formula = Replace(gpaFormula, "L56", "Q56")
    
    
    Dim totalpoint As String
    totalpoint = "=SUM(L58:Q58)"
    ws.Range("R58").Formula = totalpoint
    
    Dim totalgpa As String
    totalgpa = "=SUM(L59:Q59)"
    ws.Range("R59").Formula = totalgpa
    
End Sub


Sub ThirdSemester()

    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    With ws.Range("K85:R90")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Naming the cells
    ws.Range("K86").Value = "Total Score"
    ws.Range("K87").Value = "Percentage"
    ws.Range("K88").Value = "Grade"
    ws.Range("K89").Value = "12-Point"
    ws.Range("K90").Value = "GPA"
    ws.Range("R85").Value = "Total Score"
    
    ' Naming the Columns
    ws.Range("L85").Value = "Alternative Comm & Augmentative Devices"
    ws.Range("M85").Value = "Medication & Pharmacology"
    ws.Range("N85").Value = "Dual Diagnosis & Other Complex Needs"
    ws.Range("O85").Value = "Behavior Management III"
    ws.Range("P85").Value = "Placement"
    ws.Range("Q85").Value = "Human Sexuality"
    
    ' 1. Calculate Total Scores (Raw Points Earned)
    ws.Range("L86").Formula = "=SUM(B86:B112)"
    ws.Range("M86").Formula = "=SUM(C86:C112)"
    ws.Range("N86").Formula = "=SUM(D86:D112)"
    ws.Range("O86").Formula = "=SUM(E86:E112)"
    ws.Range("P86").Formula = "=SUM(F86:F112)"
    ws.Range("Q86").Formula = "=SUM(G86:G112)"
   
    
    
    ws.Range("B113").Formula = "=SUM(B86:B112)"
    ws.Range("C113").Formula = "=SUM(C86:c112)"
    ws.Range("D113").Formula = "=SUM(D86:D112)"
    ws.Range("E113").Formula = "=SUM(E86:E112)"
    ws.Range("F113").Formula = "=SUM(F86:F112)"
    ws.Range("G113").Formula = "=SUM(G86:G112)"
    
    
    ' 2. Calculate Percentages (Points Earned / Points Possible)
    ws.Range("L87").Formula = "=L86/100"
    ws.Range("M87").Formula = "=M86/100"
    ws.Range("N87").Formula = "=N86/100"
    ws.Range("O87").Formula = "=O86/100"
    ws.Range("P87").Formula = "=P86/100"
    ws.Range("Q87").Formula = "=Q86/100"
    
    
    ws.Range("B114").Formula = "=B113/100"
    ws.Range("C114").Formula = "=C113/100"
    ws.Range("D114").Formula = "=D113/100"
    ws.Range("E114").Formula = "=E113/100"
    ws.Range("F114").Formula = "=F113/100"
    ws.Range("G114").Formula = "=G113/100"
    

    ' 3. Grade Formula
    Dim gradeFormula As String
    gradeFormula = "=IF(L87>=90%,""A+"",IF(L87>=85%,""A"",IF(L87>=80%,""A-"",IF(L87>=77%,""B+"",IF(L87>=73%,""B"",IF(L87>=70%,""B-"",IF(L87>=67%,""C+"",IF(L87>=63%,""C"",IF(L87>=60%,""C-"",IF(L87>=57%,""D+"",IF(L87>=53%,""D"",IF(L87>=50%,""D-"",""F""))))))))))))"

    ws.Range("L88").Formula = gradeFormula
    ws.Range("M88").Formula = Replace(gradeFormula, "L87", "M87")
    ws.Range("N88").Formula = Replace(gradeFormula, "L87", "N87")
    ws.Range("O88").Formula = Replace(gradeFormula, "L87", "O87")
    ws.Range("P88").Formula = Replace(gradeFormula, "L87", "P87")
    ws.Range("Q88").Formula = Replace(gradeFormula, "L87", "Q87")
    
    
    ' 4. Point Formula
    Dim PointFormula As String
    PointFormula = "=IF(L87>=90%,12,IF(L87>=85%,11,IF(L87>=80%,10,IF(L87>=77%,9,IF(L87>=73%,8,IF(L87>=70%,7,IF(L87>=67%,6,IF(L87>=63%,5,IF(L87>=60%,4,IF(L87>=57%,3,IF(L87>=53%,2,IF(L87>=50%,1,0))))))))))))"

    ws.Range("L89").Formula = PointFormula
    ws.Range("M89").Formula = Replace(PointFormula, "L87", "M87")
    ws.Range("N89").Formula = Replace(PointFormula, "L87", "N87")
    ws.Range("O89").Formula = Replace(PointFormula, "L87", "O87")
    ws.Range("P89").Formula = Replace(PointFormula, "L87", "P87")
    ws.Range("Q89").Formula = Replace(PointFormula, "L87", "Q87")
    
    
    ' 4. GPA Formula
    Dim gpaFormula As String
    gpaFormula = "=IF(L87>=90%,4,IF(L87>=85%,3.9,IF(L87>=80%,3.7,IF(L87>=77%,3.3,IF(L87>=73%,3,IF(L87>=70%,2.7,IF(L87>=67%,2.3,IF(L87>=63%,2,IF(L87>=60%,1.7,IF(L87>=57%,1.3,IF(L87>=53%,1,IF(L87>=50%,0.7,0))))))))))))"

    ws.Range("L90").Formula = gpaFormula
    ws.Range("M90").Formula = Replace(gpaFormula, "L87", "M87")
    ws.Range("N90").Formula = Replace(gpaFormula, "L87", "N87")
    ws.Range("O90").Formula = Replace(gpaFormula, "L87", "O87")
    ws.Range("P90").Formula = Replace(gpaFormula, "L87", "P87")
    ws.Range("Q90").Formula = Replace(gpaFormula, "L87", "Q87")
    
    
    Dim totalpoint As String
    totalpoint = "=SUM(L89:Q89)"
    ws.Range("R89").Formula = totalpoint
    
    Dim totalgpa As String
    totalgpa = "=SUM(L90:Q90)"
    ws.Range("R90").Formula = totalgpa

End Sub



Sub FinalSemester()

    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    With ws.Range("K122:Q127")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Naming the cells
    ws.Range("K123").Value = "Total Score"
    ws.Range("K124").Value = "Percentage"
    ws.Range("K125").Value = "Grade"
    ws.Range("K126").Value = "12-Point"
    ws.Range("K127").Value = "GPA"
    ws.Range("Q122").Value = "Total Score"
    
    ' Naming the Columns
    ws.Range("L122").Value = "Case Study and Field Seminar"
    ws.Range("M122").Value = "Lab"
    ws.Range("N122").Value = "Core Competencies & Interviewing Skills"
    ws.Range("O122").Value = "Behavior Management IV"
    ws.Range("P122").Value = "Placement 3"
    
    ' 1. Calculate Total Scores (Raw Points Earned)
    ws.Range("L123").Formula = "=SUM(B123:B142)"
    ws.Range("M123").Formula = "=SUM(C123:C142)"
    ws.Range("N123").Formula = "=SUM(D123:D142)"
    ws.Range("O123").Formula = "=SUM(E123:E142)"
    ws.Range("P123").Formula = "=SUM(F123:F142)"
   
    
    
    ws.Range("B143").Formula = "=SUM(B123:B142)"
    ws.Range("C143").Formula = "=SUM(C123:C142)"
    ws.Range("D143").Formula = "=SUM(D123:D142)"
    ws.Range("E143").Formula = "=SUM(E123:E142)"
    ws.Range("F143").Formula = "=SUM(E123:E142)"
    
    
    ' 2. Calculate Percentages (Points Earned / Points Possible)
    ws.Range("L124").Formula = "=L123/100"
    ws.Range("M124").Formula = "=M123/100"
    ws.Range("N124").Formula = "=N123/100"
    ws.Range("O124").Formula = "=O123/100"
    ws.Range("P124").Formula = "=P123/100"
    
    
    ws.Range("B144").Formula = "=B143/100"
    ws.Range("C144").Formula = "=C143/100"
    ws.Range("D144").Formula = "=D143/100"
    ws.Range("E144").Formula = "=E143/100"
    ws.Range("F144").Formula = "=F143/100"
    

    ' 3. Grade Formula
    Dim gradeFormula As String
    gradeFormula = "=IF(L124>=90%,""A+"",IF(L124>=85%,""A"",IF(L124>=80%,""A-"",IF(L124>=77%,""B+"",IF(L124>=73%,""B"",IF(L124>=70%,""B-"",IF(L124>=67%,""C+"",IF(L124>=63%,""C"",IF(L124>=60%,""C-"",IF(L124>=57%,""D+"",IF(L124>=53%,""D"",IF(L124>=50%,""D-"",""F""))))))))))))"

    ws.Range("L125").Formula = gradeFormula
    ws.Range("M125").Formula = Replace(gradeFormula, "L124", "M124")
    ws.Range("N125").Formula = Replace(gradeFormula, "L124", "N124")
    ws.Range("O125").Formula = Replace(gradeFormula, "L124", "O124")
    ws.Range("P125").Formula = Replace(gradeFormula, "L124", "P124")
    
    
    ' 4. Point Formula
    Dim PointFormula As String
    PointFormula = "=IF(L124>=90%,12,IF(L124>=85%,11,IF(L124>=80%,10,IF(L124>=77%,9,IF(L124>=73%,8,IF(L124>=70%,7,IF(L124>=67%,6,IF(L124>=63%,5,IF(L124>=60%,4,IF(L124>=57%,3,IF(L124>=53%,2,IF(L124>=50%,1,0))))))))))))"

    ws.Range("L126").Formula = PointFormula
    ws.Range("M126").Formula = Replace(PointFormula, "L124", "M124")
    ws.Range("N126").Formula = Replace(PointFormula, "L124", "N124")
    ws.Range("O126").Formula = Replace(PointFormula, "L124", "O124")
    ws.Range("P126").Formula = Replace(PointFormula, "L124", "P124")
    
    
    ' 4. GPA Formula
    Dim gpaFormula As String
    gpaFormula = "=IF(L124>=90%,4,IF(L124>=85%,3.9,IF(L124>=80%,3.7,IF(L124>=77%,3.3,IF(L124>=73%,3,IF(L124>=70%,2.7,IF(L124>=67%,2.3,IF(L124>=63%,2,IF(L124>=60%,1.7,IF(L124>=57%,1.3,IF(L124>=53%,1,IF(L124>=50%,0.7,0))))))))))))"

    ws.Range("L127").Formula = gpaFormula
    ws.Range("M127").Formula = Replace(gpaFormula, "L124", "M124")
    ws.Range("N127").Formula = Replace(gpaFormula, "L124", "N124")
    ws.Range("O127").Formula = Replace(gpaFormula, "L124", "O124")
    ws.Range("P127").Formula = Replace(gpaFormula, "L124", "P124")
 
    
    Dim totalpoint As String
    totalpoint = "=SUM(L126:P126)"
    ws.Range("Q126").Formula = totalpoint
    
    Dim totalgpa As String
    totalgpa = "=SUM(L127:P127)"
    ws.Range("Q127").Formula = totalgpa

End Sub

Sub CGPA_Calculator()
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    With ws.Range("H145:I151")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 255)
        .HorizontalAlignment = xlCenter
    End With
    
        ' Naming the cells
    ws.Range("H146").Value = "Total Points"
    ws.Range("H147").Value = "GPA"
    ws.Range("H149").Value = "Total Number of Courses "
    ws.Range("H150").Value = "Cummulative - GPA/23"
    ws.Range("H151").Value = "Final GPA"
    
    Dim cumulativepoint As String
    cumulativepoint = "=SUM(Q126,R89,R58,S8)"
    ws.Range("I146").Formula = cumulativepoint
    
    Dim cumulativegpa As String
    cumulativegpa = "=SUM(Q127,R90,R59,S9)"
    ws.Range("I147").Formula = cumulativegpa
    ws.Range("I149").Value = 23
    ws.Range("I150").Formula = "=I147/I149"
    ws.Range("I151").Formula = "=ROUND(I150,1)"
    
End Sub


Sub AddSemesterTable(doc As Object, semesterTitle As String, courseData As Variant)
    Dim tbl As Object
    Dim r As Long, i As Integer
    Dim titleRange As Object
    
    ' Insert the title at the end of the document
    Set titleRange = doc.Content
    titleRange.Collapse Direction:=0 ' Collapse to end
    titleRange.InsertParagraphAfter
    titleRange.InsertAfter semesterTitle & vbCrLf
    titleRange.Font.Bold = True
    titleRange.Font.Size = 11 ' Optional: adjust size

    ' Move range to end again to insert table
    titleRange.Collapse Direction:=0
    Set tbl = doc.Tables.Add(Range:=titleRange, NumRows:=UBound(courseData) + 2, NumColumns:=3)
    
    ' Headers
    tbl.Cell(1, 1).Range.Text = "COURSE CODE"
    tbl.Cell(1, 2).Range.Text = "COURSE DESCRIPTION"
    tbl.Cell(1, 3).Range.Text = "GRADE"
    
    ' Bold headers
    For i = 1 To 3
        tbl.Cell(1, i).Range.Bold = True
        tbl.Cell(1, i).Range.ParagraphFormat.Alignment = 0 ' Left align
    Next i
    
    ' Fill in course data
    For r = 0 To UBound(courseData)
        tbl.Cell(r + 2, 1).Range.Text = courseData(r)(0)
        tbl.Cell(r + 2, 2).Range.Text = courseData(r)(1)
        tbl.Cell(r + 2, 3).Range.Text = courseData(r)(2)
    Next r
    
    ' Set font size for entire table
    tbl.Range.Font.Size = 9 ' Change to desired font size

    tbl.Rows.Alignment = 1 ' Align table left
    tbl.Range.ParagraphFormat.SpaceAfter = 1
    tbl.AllowAutoFit = True
    tbl.Columns(1).PreferredWidth = 150
    tbl.Columns(2).PreferredWidth = 300
    tbl.Columns(3).PreferredWidth = 60
End Sub



Sub AddGradeTable(doc As Object, semesterTitle As String, GradeData As Variant)
    Dim tbl As Object
    Dim r As Long, i As Integer
    
    With doc.Content
        .InsertAfter vbCrLf & semesterTitle & vbCrLf
        Set tbl = doc.Tables.Add(Range:=.Paragraphs.Last.Range, NumRows:=UBound(GradeData) + 1, NumColumns:=2)
    End With
    
    
    ' Fill in course data
    For r = 0 To UBound(GradeData)
        tbl.Cell(r + 2, 1).Range.Text = GradeData(r)(0)
        tbl.Cell(r + 2, 2).Range.Text = GradeData(r)(1)
    Next r
    
    ' Set font size for entire table
    tbl.Range.Font.Size = 9 ' Change to desired font size
    tbl.Range.Font.Name = "Georgia"

    tbl.Rows.Alignment = 0 ' Align table left
    tbl.Range.ParagraphFormat.SpaceAfter = 1
    tbl.AllowAutoFit = True
    tbl.Columns(1).PreferredWidth = 150
    tbl.Columns(2).PreferredWidth = 80
End Sub

Sub AddGPATable(doc As Object, semesterTitle As String, courseData As Variant)
    Dim tbl As Object
    Dim r As Long, i As Integer
    Dim titleRange As Object
    
    ' Insert the title at the end of the document
    Set titleRange = doc.Content
    titleRange.Collapse Direction:=0 ' Collapse to end
    titleRange.InsertParagraphAfter
    titleRange.InsertAfter semesterTitle & vbCrLf
    titleRange.Font.Bold = True
    titleRange.Font.Size = 10 ' Optional: adjust size

    ' Move range to end again to insert table
    titleRange.Collapse Direction:=0
    Set tbl = doc.Tables.Add(Range:=titleRange, NumRows:=UBound(courseData) + 2, NumColumns:=4)
    
    ' Headers
    tbl.Cell(1, 1).Range.Text = "Earned Hours"
    tbl.Cell(1, 2).Range.Text = "GPA Hours"
    tbl.Cell(1, 3).Range.Text = "TOTAL POINTS"
    tbl.Cell(1, 4).Range.Text = "GPA"
    
    ' Bold headers
    For i = 1 To 4
        tbl.Cell(1, i).Range.Bold = True
        tbl.Cell(1, i).Range.ParagraphFormat.Alignment = 0 ' Left align
    Next i
    
    ' Fill in course data
    For r = 0 To UBound(courseData)
        tbl.Cell(r + 2, 1).Range.Text = courseData(r)(0)
        tbl.Cell(r + 2, 2).Range.Text = courseData(r)(1)
        tbl.Cell(r + 2, 3).Range.Text = courseData(r)(2)
        tbl.Cell(r + 2, 4).Range.Text = courseData(r)(3)
    Next r
    
    ' Set font size for entire table
    tbl.Range.Font.Size = 9 ' Change to desired font size

    tbl.Rows.Alignment = 1 ' Align table left
    tbl.Range.ParagraphFormat.SpaceAfter = 1
    tbl.AllowAutoFit = True
    tbl.Columns(1).PreferredWidth = 150
    tbl.Columns(2).PreferredWidth = 150
    tbl.Columns(3).PreferredWidth = 150
    tbl.Columns(4).PreferredWidth = 150
End Sub




Sub GenerateTranscriptToWord()
    Dim wdApp As Object, wdDoc As Object
    Dim pws As Worksheet
    Dim studentName As String, studentID As String
    Dim enrolmentDate As String, address As String, issueDate As String
    Dim awardDate As String, attendancePeriod As String
    Dim totalPoints As String, gpa As String
    
    Set pws = Worksheets(ActiveSheet.Index)

    ' Get data
    studentName = pws.Range("B147").Value
    studentID = pws.Range("B148").Value
    enrolmentDate = pws.Range("B149").Value
    address = pws.Range("B150").Value
    issueDate = pws.Range("B151").Value
    attendancePeriod = pws.Range("B152").Value
    awardDate = pws.Range("B153").Value
    totalPoints = pws.Range("I146").Value
    gpa = pws.Range("I151").Value

    ' Open Word
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0
    wdApp.Visible = True

    ' Create document
    Set wdDoc = wdApp.Documents.Add
    Dim logoShape As Shape
    Set logoShape = Worksheets("Shape").Shapes("Picture 1") ' Change to the actual shape name

    With wdDoc
        ' Insert a 2-column table for logo and header
        Dim hdrTbl As Object
        Set hdrTbl = .Tables.Add(.Range, 1, 2)
        hdrTbl.Rows.Alignment = 1 ' Left
        hdrTbl.Borders.Enable = False
        hdrTbl.Columns(1).Width = 100
        hdrTbl.Columns(2).Width = 400
        
        ' Insert logo
        logoShape.Copy
        hdrTbl.Cell(1, 1).Range.Paste
        
        ' Optional: Resize the pasted image
        With hdrTbl.Cell(1, 1).Range.InlineShapes(1)
            .LockAspectRatio = msoTrue
            .Width = 100 ' Adjust as needed
        End With
        
        ' Header text
        With hdrTbl.Cell(1, 2).Range
            .Text = "" ' Clear existing content first
        
            ' College Name
            .InsertAfter "Behaviorprise College of Business & Health Studies" & vbCrLf
            With .Paragraphs(1).Range
                .Font.Name = "Calibri"
                .Font.Size = 16
                .Font.Bold = True
                .ParagraphFormat.SpaceAfter = 1
            End With
        
            ' Address
            .InsertAfter "Office of the Campus Administrator, 800 Petrolia Road, Unit 16, North York, Ontario. M3J 3K4" & vbCrLf
            With .Paragraphs(2).Range
                .Font.Name = "Calibri"
                .Font.Size = 11
                .Font.Bold = False
                .ParagraphFormat.SpaceAfter = 2
            End With
        
            ' Divider line
            .InsertAfter "_______________________________________________________________" & vbCrLf
            With .Paragraphs(3).Range
                .Font.Size = 12
                .Font.Name = "Courier New"
                .Font.Bold = True
                .ParagraphFormat.SpaceAfter = 10
            End With
        
            ' Subtitle
            .InsertAfter "Student Academic Unofficial Transcript" & vbCrLf
            With .Paragraphs(4).Range
                .Font.Name = "Lucida Calligraphy"
                .Font.Size = 12
                .Font.Bold = True
                .ParagraphFormat.SpaceAfter = 10
            End With
        
            ' Student info
            .InsertAfter "Name:  " & studentName & vbTab & "     " & vbTab & vbTab & vbTab & "Student ID: " & studentID
            With .Paragraphs(5).Range
                .Font.Name = "Calibri"
                .Font.Size = 12
                .Font.Bold = True
            End With
        End With
        
        ' Student info
        Dim infoText As String
        infoText = "Enrolment Date: " & pws.Range("B149").Value & vbCrLf & _
            "Student Address: " & pws.Range("B150").Value & vbCrLf & _
            "Date of Issue: " & pws.Range("B151").Value & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Course Level: Diploma" & vbCrLf & _
            "Current Program: Developmental Service Worker – Behavioral" & vbTab & _
            "Period of attendance: " & pws.Range("B152").Value & vbCrLf & _
            "Academic Award(s) obtained:" & vbTab & "Date of Award: " & pws.Range("B153").Value & vbCrLf
        
        With .Paragraphs.Add.Range
            .Text = infoText
            .Font.Name = "Calibri"
            .Font.Size = 9
            .Font.Bold = True
        End With
        ' Add section headers
        With .Paragraphs.Add.Range
            .Text = "TRANSFER CREDIT FROM OTHER ACADEMIC INSTITUTION(S):   NONE" & vbCrLf & _
                    "INSTITUTION CREDIT:"
            .Font.Name = "Calibri"
            .Font.Size = 10
            .Font.Bold = True
        End With
    End With

    ' Add tables
    Call AddSemesterTable(wdDoc, "SEMESTER 1", Array( _
        Array("DSW 001", "Introduction to Developmental Disabilities", pws.Range("O7").Value), _
        Array("DSW 002", "Human Development", pws.Range("M7").Value), _
        Array("DSW 003", "Behavior Management I", pws.Range("P7").Value), _
        Array("GEN 004", "College Writing Skills", pws.Range("Q7").Value), _
        Array("GEN 005", "Classroom Dynamics", pws.Range("L7").Value), _
        Array("GEN 006", "Professional Ethics", pws.Range("R7").Value), _
        Array("GEN 007", "Use of Computers – Microsoft Office Suite", pws.Range("N7").Value) _
    ))

    ' Repeat for other semesters...
    ' SEMESTER 2
    Call AddSemesterTable(wdDoc, "SEMESTER 2", Array( _
        Array("DSW 009", "Abnormal Psychology", pws.Range("N57").Value), _
        Array("DSW 010", "Behavior Management II", pws.Range("O57").Value), _
        Array("DSW 011", "Social Inclusion Strategies – QAM, CFSA", pws.Range("P57").Value), _
        Array("DSW 012", "Group Dynamics", pws.Range("L57").Value), _
        Array("DSW 013", "Program Development – Person Centered & Group", pws.Range("M57").Value), _
        Array("DSW 014", "Field Placement I", pws.Range("Q57").Value)))
        
    With wdDoc.Content
        .InsertAfter vbCrLf & "*********************************END OF PAGE, NO ENTRIES BELOW THIS LINE***********************" & vbCrLf & vbCrLf
        .Font.Size = 10
    End With

    ' SEMESTER 3
    Call AddSemesterTable(wdDoc, "SEMESTER 3", Array( _
        Array("DSW 201", "Alternative Comm & Augmentative Devices", pws.Range("L88").Value), _
        Array("DSW 202", "Behavior Management III", pws.Range("O88").Value), _
        Array("DSW 203", "Medication and Pharmacology", pws.Range("M88").Value), _
        Array("DSW 204", "Dual Diagnosis & Other Complex Needs", pws.Range("N88").Value), _
        Array("DSW 205", "Human Sexuality", pws.Range("Q88").Value), _
        Array("DSW 206", "Field Placement II", pws.Range("P88").Value) _
    ))

    ' SEMESTER 4
    Call AddSemesterTable(wdDoc, "SEMESTER 4", Array( _
        Array("DSW 251", "Behavior Management IV", pws.Range("O125").Value), _
        Array("DSW 252", "Lab", pws.Range("M125").Value), _
        Array("DSW 253", "Core Competencies & Interviewing Skills", pws.Range("N125").Value), _
        Array("DSW 254", "Field Placement III", pws.Range("P125").Value), _
        Array("DSW 255", "Case Study and Field Seminar", pws.Range("L125").Value) _
    ))
    
    ' GPA Table
    
    Call AddGPATable(wdDoc, "TRANSCRIPT TOTALS", Array( _
        Array("57", "57", pws.Range("I146").Value, pws.Range("I151").Value)))

    ' Add footer info
    With wdDoc.Content
        .InsertAfter "INC - Incomplete" & vbCrLf
        .InsertAfter "AB - Absent" & vbCrLf
        .InsertAfter "NA – Not Applicable" & vbCrLf
        .InsertAfter "Transcript valid only if bearing original signature of the Campus Administrator with the official college seal." & vbCrLf
        .InsertAfter "*********************************END OF TRANSCRIPT***********************"
        .Font.Size = 10
    End With
    
       Call AddGradeTable(wdDoc, "Grading System:", Array( _
        Array("A+", "90–100"), _
        Array("A", "85–89"), _
        Array("A-", "80–84"), _
        Array("B+", "77–79"), _
        Array("B", "73–76"), _
        Array("B-", "70–72"), _
        Array("C+", "67–69"), _
        Array("C", "63–66"), _
        Array("C-", "60–62"), _
        Array("D", "53–56"), _
        Array("D-", "50–52"), _
        Array("E, F", "0–49") _
    ))

    MsgBox "Transcript Generated successfully!"
End Sub


