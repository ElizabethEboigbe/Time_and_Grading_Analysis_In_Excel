# Excel Data Analysis Function

## Table of Contents
- [Overview](#overview)
- [Project Objectives](#project-objectives)
- [1. Time Functions](#1-time-functions)
  - [Data Used](#data-used)
  - [Description](#description)
  - [Key Questions](#key-questions)
  - [Formulas Used](#formulas-used)
  - [Key Insights](#key-insights)
- [2. Time Functions (Student Dataset)](#2-time-functions-student-dataset)
  - [Data Used](#data-used-1)
  - [Description](#description-1)
  - [Key Questions and Answers](#key-questions-and-answers)
  - [Formulas Used](#formulas-used-1)
  - [Key Insights](#key-insights-1)
- [3. Grading System](#3-grading-system)
  - [Data Used](#data-used-2)
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

### Data Used
[file link for dataset]

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
- `J11 = 01/01/2000 (Start Date)`  
- `M11 = 10/07/2023 (End Date)`  
- `A3 = 2 (month number)`  
- `E2 = Current date (today)`  
- `E3 = Current date and time (now)`  

```excel
=TEXT(A3*29,"MMMM")
=TODAY()
=NOW()
=TIME(10,15,0)
=YEAR(E2)
=MONTH(E2)
=DAY(E2)
=EDATE(E2,-10)
=EDATE(E3,-6)
=EOMONTH(E2,7)
=DATEDIF(J11,M11,"Y")
=DATEDIF(J11,M11,"M")
=DATEDIF(J11,M11,"MD")
=DATEDIF(J11,M11,"YM")
=DATEDIF(J11,M11,"YD")
=DATEDIF(J11,TODAY(),"Y")
=WORKDAY(J11,10)
=WEEKDAY(J11)







# Excel Data Analysis Function

## Table of Contents
- [Overview](#overview)
- [Project Objectives](#project-objectives)
- [1. Time Functions](#1-time-functions)
  - [Data Used](#data-used)
  - [Description](#description)
  - [Key Questions](#key-questions)
  - [Formulas Used](#formulas-used)
  - [Key Insights](#key-insights)
- [2. Time Functions (Student Dataset)](#2-time-functions-student-dataset)
  - [Data Used](#data-used-1)
  - [Description](#description-1)
  - [Key Questions and Answers](#key-questions-and-answers)
  - [Formulas Used](#formulas-used-1)
  - [Key Insights](#key-insights-1)
- [3. Grading System](#3-grading-system)
  - [Data Used](#data-used-2)
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

### Data Used
[file link for dataset]

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
- `J11 = 01/01/2000 (Start Date)`  
- `M11 = 10/07/2023 (End Date)`  
- `A3 = 2 (month number)`  
- `E2 = Current date (today)`  
- `E3 = Current date and time (now)`  

```excel
=TEXT(A3*29,"MMMM")    // Convert a month number into a full month name
=TODAY()               // Current date (recalculates daily)
=NOW()                 // Current date and time
=TIME(10,15,0)         // Creates a time value for 10:15:00 (H:M:S)
=YEAR(E2)              // Year from a date in E2
=MONTH(E2)             // Month from E2
=DAY(E2)               // Day of month from E2
=EDATE(E2,-10)         // Date 10 months earlier than E2
=EDATE(E3,-6)          // Back date by 6 months (positive for future)
=EOMONTH(E2,7)         // Last day of 7 months after E2
=DATEDIF(J11,M11,"Y")  // Number of complete years
=DATEDIF(J11,M11,"M")  // Number of complete months
=DATEDIF(J11,M11,"MD") // Days difference, ignoring months and years
=DATEDIF(J11,M11,"YM") // Months difference, ignoring years
=DATEDIF(J11,M11,"YD") // Days difference, ignoring years
=DATEDIF(J11,TODAY(),"Y")  // Current age (in years)
=WORKDAY(J11,10)       // Date after 10 working days
=WEEKDAY(J11)          // Numeric representation of weekday (1-7)

