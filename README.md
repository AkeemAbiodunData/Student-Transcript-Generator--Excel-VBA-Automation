# Student-Transcript-Generator--Excel-VBA-Automation

## Project Overview
This project automates the generation of a student’s academic transcript for the all semester using Microsoft Excel and VBA. Designed for educational institutions and administrative staff, the macro formats the transcript layout, computes grades, and presents results in a standardized and visually coherent format. The automation improves accuracy, saves time, and ensures consistency across student records.

The solution is implemented using Visual Basic for Applications (VBA) within Excel, enabling seamless integration with existing grade sheets and dynamic customization.

## Dataset / Input
The macro expects structured student performance data in an Excel worksheet, including:

Subject Scores (e.g., Classroom Dynamics, Human Development)

Grading Criteria (Raw scores out of 100)

Defined Output Range (K4:S9 for result metrics and subject layout)

## Subjects Included:
Classroom Dynamics

Human Development

Use of Computers

Introduction to Developmental Disabilities

Behavior Management I

College Writing Skills

Professional Ethics

Automation Techniques

## Data Formatting & Layout:
Applies uniform font and cell styling to the results section

Adds headers for grading categories: Total Score, Percentage, Grade, 12-Point Scale, GPA

Centrally aligns content for clarity

## Grade Computation:
Calculates Total Score from subject scores

Computes Percentage based on the number of subjects

Assigns Letter Grades using score thresholds

Converts grades to a 12-Point Scale

Computes GPA from 12-point equivalents

## VBA Logic:
Uses Excel cell ranges and formulas embedded via VBA

Implements logic with If-Else conditions to map scores to grades

Employs reusable code structure for scalability to additional semesters

## Key Functionalities
📊 Automatic Grade Calculation

🧮 GPA Computation Based on Configurable Logic

🖋️ Predefined Layout Styling

✅ Error-Free Processing of Scores

📄 Ready for Export and Printing

How to Use
Open the Excel workbook where student scores are stored.

Press ALT + F11 to launch the VBA Editor.

Go to File > Import File... and select the Module1.bas file.

Return to Excel and press ALT + F8 to open the macro menu.

Select FirstSemester and click Run.

⚠️ Ensure the worksheet layout matches expected input format (columns L to R contain subject scores).

Key Outputs
Total Score: Sum of all subject scores

Percentage: Average score as a percentage

Grade: A-F grade assigned based on percentage

12-Point Equivalent: Numeric GPA scale

Cumulative GPA: Final calculated GPA

Visualizations
The macro transforms raw data into a clearly structured table featuring academic metrics. The result can be exported or printed as an official transcript for each student.

screenshots of the code
![Image](https://github.com/user-attachments/assets/ce7fa8b6-d82b-4e8b-812f-92706494acaa)

![Image](https://github.com/user-attachments/assets/e869cc21-54a3-46c1-afe8-5e140d29cad2)

![Image](https://github.com/user-attachments/assets/6708f615-2019-42f6-961b-e78a0fe0c00f)

![Image](https://github.com/user-attachments/assets/184e3b1c-3d6b-4ad6-a1cf-1cdf3ba3b52c)


Conclusion
This Excel VBA solution simplifies the process of academic transcript generation. By automating grading and formatting, institutions can maintain consistent reporting standards while minimizing manual work. It serves as a robust tool for academic administration, especially in institutions that handle large student volumes.
