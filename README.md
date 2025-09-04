# Excel Data Analysis Function

## Table of Contents
- [Overview](#overview)
- [Project Objectives](#project-objectives)
- [1. Time Functions](#1-time-functions)
  - [Data used](#data-used)
  - [Description](#description)
  - [Key Questions](#key-questions)
  - [Formulas Used](#formulas-used)
  - [Key Insights](#key-insights)
- [2. Time Functions (Student Dataset)](#2-time-functions-student-dataset)
  - [Data used](#data-used-1)
  - [Description](#description-1)
  - [Key Questions and Answers](#key-questions-and-answers)
  - [Formulas Used](#formulas-used-1)
  - [Key Insights](#key-insights-1)
- [3. Grading System](#3-grading-system)
  - [Data used](#data-used-2)
  - [Description](#description-2)
  - [Formulas Used](#formulas-used-2)
  - [Key Questions and Answers](#key-questions-and-answers-1)
  - [Key Insights](#key-insights-2)
- [Conclusion](#conclusion)

---

## Overview
This project demonstrates the use of Excel formulas and functions to perform data analysis and automation tasks.  
It is divided into three main sections:
1. Time functions  
2. Time function student (Student Dataset)  
3. Grading system  

The project highlights how Excel can be used to manipulate dates, extract insights from datasets, and automate grading with formulas.

---

## Project Objectives
This project demonstrates how Excel functions can be applied to automate and simplify complex tasks such as calculating time differences, extracting date components, and building dynamic grading systems. By leveraging formulas instead of manual input, the project highlights how efficiency, accuracy, and consistency can be achieved in handling real-world data challenges.

---

## 1. Time Functions

### Data used
 - <a href="https://github.com/ElizabethEboigbe/Time_and_Grading_Analysis_In_Excel/blob/main/Time%20Function.xlsx">Dataset</a>
 
### Description
This dataset demonstrates the use of Excel time and date functions to calculate and manipulate values such as differences between dates, extracting parts of dates, and determining weekdays or workdays.

### Key Questions
- How many complete years and months exist between two dates?  
- What is the remaining days difference when years/months are ignored?  
- What weekday does the given date fall on?  
- How can a month number be converted to its name?  
- How is a sheet kept current with today’s date/time automatically?  
- How many complete years exist between the start date and the end date?  
- What is the total month difference between the start date and end date?  
- Calculate the difference between two dates in years, months, and days.  

### Formulas Used
Reference cells include:  
- J11 = 01/01/2000 (Start Date)  
- M11 = 10/07/2023 (End Date)  
- A3 = 2 (month number)  
- E2 = Current date (today)  
- E3 = Current date and time (now)  

1. =TEXT(A3*29,"MMMM") – Convert a month number in A3 into a full month name
• A3 – contains the numeric value representing the month (e.g., 2 for February).
• 29 – multiplies the month number by 29 to generate a valid Excel date serial for that month.
• MMMM – format the date into the full month name.

2. `=TODAY()` – Current date (recalculates daily).
3. `=NOW()` – Current date and time.
4. `=TIME(10,15,0)` – Creates a time value for 10:15:00 (H:M:S).
5. `=YEAR(E2)` – Year from a date in E2.
6. `=MONTH(E2)` – Month from E2.
7. `=DAY(E2)` – Day of month from E2.
8. `=EDATE(E2,-10)` – Date 10 months earlier than E2 (back date by 10 months).
9. `=EDATE(E3,-6)` – Back date by 6 months (use positive for future).
10. `=EOMONTH(E2,7)` – Last day of 7 months after E2.
11. `=DATEDIF(J11,M11,"Y")` – Returns the total number of complete years between the two dates.
12. `=DATEDIF(J11,M11,"M")` – Returns the total number of complete months between the two dates.
13. `=DATEDIF(J11,M11,"MD")` – Returns the total number of days difference, ignoring both month and years.
14. `=DATEDIF(J11,M11,"YM")` – Returns the difference in months, ignoring years.
15. `=DATEDIF(J11,M11,"YD")` – Returns the difference in days, ignoring years.
16. `=DATEDIF(J11,TODAY(),"Y")` – Returns the current age (in years) based on the given start date and today’s date.
17. `=WORKDAY(J11,10)` – Returns the date after 10 working days from the start date.
18. `=WEEKDAY(J11)` – Returns the numeric representation of the weekday (1–7) for the given date.

### Key Insights
- `TEXT()` is a way to turn a month number into a month name.  
- `TODAY()` / `NOW()` update automatically to current time.  
- `YEAR()` extracts the year from a date.  
- `MONTH()` extracts the month from a date.  
- `EDATE` and `EOMONTH` are essential for rolling periods and month-end reporting.  
- `DATEDIF(J11,M11,"Y")` shows the year difference between start date and end date to be 23 years.  
- `DATEDIF(J11,M11,"M")` calculated the total months difference to 282 months.  
- `DATEDIF(J11,M11,"MD")` calculates the remaining days after ignoring years and months to be 9 days.  
- `DATEDIF(J11,M11,"YM")` calculates the difference in months ignoring the years to be 6 months.  
- `DATEDIF(J11,M11,"YD")` calculates the difference in days ignoring years to be 191 days.  
- `DATEDIF(J11,TODAY(),"Y")` can calculate age when linked to date of birth.  
- `WORKDAY(J11,10)` is useful in business contexts to project deadlines excluding weekends.  
- `WEEKDAY(J11)` helps determine which day of the week a given date falls on.  

---

## 2. Time Functions (Student Dataset)

### Data used
- <a href="https://github.com/ElizabethEboigbe/Time_and_Grading_Analysis_In_Excel/blob/main/Time%20Function-Student.xlsx">Dataset</a>
 
### Description
This dataset contains detailed student records with columns such as:  
Student number, surname, title, first name, other initials, hall, user id, tutor, option, date of birth, departure date.  

Additional derived columns were created using Excel formulas to extract and calculate important details such as:  
- Birth day, birth year, birth month  
- Adding 6 months to date of birth  
- Month difference between two columns (date of birth and departure date)  
- Adjusting departure dates with 8 months  
- Weekday for date of birth  

All additional columns were computed dynamically with formulas to ensure accuracy and efficiency.  

### Key Questions and Answers
1) Who are the 15 youngest students?  
Using the date column, I applied sorting and filters to identify the 15 youngest students.
 
<img width="340" height="328" alt="Screenshot 2025-09-04 051920" src="https://github.com/user-attachments/assets/483b8102-3a46-4492-96a3-1bc6aa5ef62b" />

3) What are the names of students born in the month of April?  
Using the date of birth column, I clicked on the filter drop-down and clicked on date filter and from the drop-down I clicked on “all date in the period” and selected April.
 
<img width="800" height="543" alt="Screenshot 2025-09-04 052223" src="https://github.com/user-attachments/assets/de1a95ec-55f1-43f2-88a2-ff515eb8e88c" />

5) How many of tutor Robinson’s students live in private accommodation?  
Using the tutor’s column filter, I selected “Robinson” and also filtered the hall (accommodation) to private. This showed the exact number to be 6.
 
<img width="602" height="144" alt="Screenshot 2025-09-04 052334" src="https://github.com/user-attachments/assets/93dcb2aa-8889-4a4a-92fd-f930c18056f7" />

7) Show the 17 oldest students.  
Using the filter on the date of birth column to sort in ascending order, I was able to get the 17 oldest students.  

<img width="801" height="358" alt="Screenshot 2025-09-04 052614" src="https://github.com/user-attachments/assets/f372652a-3cfe-4e34-bc16-0392c5462f85" />

8) Display all the students that are not living in private accommodation.  
From the hall column filter, I unchecked “private” and got the students not living in private accommodation.  
  
<img width="438" height="543" alt="Screenshot 2025-09-04 052730" src="https://github.com/user-attachments/assets/0fc56b16-c152-4541-a77a-0e3ef819b3a6" />

9) How many blank spaces are there in “other_initials” column?  
Using the count blank formula, there are 114 blank spaces in other_initials.  

10) Who are the students without middle name as initial?  
Using the other_initial filter, I was able to get the names of students without middle names.  
      
<img width="349" height="546" alt="Screenshot 2025-09-04 052832" src="https://github.com/user-attachments/assets/79dc4b10-d59b-4bef-8151-6c71f366dfdf" />

11) Who are the students with middle name as initial?  
The other_initial column filter was used to get the names of students with middle names.  
   
<img width="344" height="542" alt="Screenshot 2025-09-04 052923" src="https://github.com/user-attachments/assets/5453b3b7-960c-4432-8f72-d44a5a913761" />

12) Show all the St Patrick’s students that live in Bridges accommodation.  
The hall and tutor column filter was used to identify St Patrick’s students that live in Bridges accommodation. One student. 
   
<img width="441" height="46" alt="Screenshot 2025-09-04 053046" src="https://github.com/user-attachments/assets/6ca8d8f2-def4-4282-ace2-6d3bf3db2686" />

### Formulas Used
Reference cells: J2 (date of birth), K2 (departure date), M (month)  

`=DAY(J2)` – Extract the day from the Date of Birth.
`=YEAR(J2)` – Extract the year from the Date of Birth.
`=MONTH(J2)` – Extract the month from the Date of Birth.
`=EDATE(J2,6)` – Add 6 months to the Date of Birth.
`=DATEDIF(J2,K2,"M")` – Calculate month difference between two dates.
`=EDATE(K2,-8)` – Backdate departure date with 8 months.
`=WEEKDAY(J2)` – Return the weekday for the Date of Birth.

### Key Insights
- Dates can be dynamically transformed to derive age, month, or day values without manual entry.  
- `EDATE` and `DATEDIF` functions simplify calculations involving differences between dates or future projections.  
- Automating fields like Birth Year or Weekday ensures consistency and reduces human error in large datasets.  
- This approach demonstrates how raw student records can be enriched with additional, useful time-based insights using only formulas.  

---

## 3. Grading System

### Data used

- <a href="https://github.com/ElizabethEboigbe/Time_and_Grading_Analysis_In_Excel/blob/main/Grading%20System.xlsx">Dataset<a/>
 
### Description
The grading system dataset contained student names, marks, and grades, where Excel formulas such as `LOOKUP`, `IF`, and `COUNT` were applied to assign grades, determine qualification status, and summarize results.

### Formulas Used
`=VLOOKUP()` – To automatically assign grades based on the marks obtained by students.
`=LOOKUP(C11,$B$4:$B$7,$C$4:$C$7)`

`=IF(E38>30,"qualify for tomorrow python test","not qualify for tomorrow python test")` – To determine whether a student was Qualified or Not Qualified based on their score.

`=COUNTIF(F38:F49,"qualify for tomorrow python test")` – To count how many students were qualified and how many were not.

`=LEFT()`, `=RIGHT()`, `=MID()`, and `=LEN()` – To split student full names into First Name, Last Name, and Number.

### Key Questions and Answers
1. What grades correspond to each student’s score?  
Using Excel formula (LOOKUP).
 
<img width="382" height="405" alt="Screenshot 2025-09-04 053244" src="https://github.com/user-attachments/assets/d5be44ca-9483-4274-aeea-d3746bda1727" />

3. How many students are qualified versus not qualified?  
Using formula (IF), 6 students were not qualified, while 6 students were qualified.
  
 <img width="963" height="251" alt="Screenshot 2025-09-04 053405" src="https://github.com/user-attachments/assets/6f6cf500-046a-4bfe-b0b4-0ed5d7b3e7a0" />

5. Which students achieved the highest and lowest grades?  
- Highest: **Ebuwa Edowaye** with 95 marks.  
- Lowest: **Ogbes Okpebholo** with 40 marks.  

### Key Insights
- **Automatic Grading** – Using VLOOKUP, grades were efficiently matched to marks without manual input, reducing errors and saving time.  
- **Qualification Tracking** – The IF function clearly distinguished qualified students from those who were not, making the dataset actionable for decision-making.  
- **Aggregated Results** – The COUNTIF formula provided a quick overview of total qualified versus unqualified students.  
- **Data Cleaning & Transformation** – Splitting full names into separate components demonstrated how text formulas in Excel can prepare messy data for structured analysis.  

---

## Conclusion
This project demonstrates how Excel functions can be applied to automate calculations, analyze datasets, and simplify decision-making processes. By working with time-based functions, student datasets, and grading systems, the analysis highlights how formulas such as `DATEDIF`, `TEXT`, `LOOKUP`, and `IF` can transform raw data into actionable insights. These approaches not only improve accuracy but also reduce manual effort, showcasing Excel’s role as a powerful tool for real-world problem solving in data analysis.
