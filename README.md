# CRM Finance Loan Tracking Excel Based Client Feedback and Loan Management System

This project explores **Excel VBA Macros** to automate **CRM and loan tracking** for financial institutions. It streamlines **client feedback logging, loan monitoring, and data retrieval** using VBA scripts and **VLOOKUP**. By automating **data entry and tracking**, it enhances efficiency, reduces errors, and optimizes Excel as a CRM tool for small businesses.

---

## **Project Overview**
Many small and medium-sized businesses use Excel for **Customer Relationship Management (CRM)** due to its accessibility and ease of use. However, Excel lacks built-in automation features found in advanced CRM platforms like **Salesforce, HubSpot, and Zoho**. This project leverages **VBA Macros** to introduce **automated tracking, logging, and retrieval** of client data in an Excel-based CRM system.

---

## **Features & Capabilities**
- **Automated Client Feedback Logging** – VBA scripts capture and timestamp customer interactions.
- **Loan Tracking & Account Management** – Automates loan data updates and debt monitoring.
- **VLOOKUP for Quick Data Retrieval** – Enables easy searching of past client notes.
- **Customizable CRM Framework** – Adaptable for different industries, including **finance, sales, and portfolio management**. -**Error Handling & Event-Driven Automation** – Prevents data loss and enhances system stability.  

---

## **Technical Implementation**
### 1. **VBA Macros for Event-Driven Logging**
- **Automatically logs client feedback** when changes occur in specified columns.  
- **Stores previous interactions** for quick reference.  

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LogSheet") ' Ensure this sheet exists

    ' Check if the change was made in Column M
    If Not Intersect(Target, Me.Range("M:M")) Is Nothing Then
        Application.EnableEvents = False ' Prevent infinite loops
        On Error GoTo ErrorHandler ' Error handling

        ' Insert a new row at Row 2 (pushing old entries down)
        ws.Rows(2).Insert Shift:=xlDown

        ' Log details in LogSheet
        ws.Cells(2, 1).Value = Now ' Timestamp (Column A)
        ws.Cells(2, 2).Value = Me.Cells(Target.Row, 2).Value ' Client (Column B)
        ws.Cells(2, 3).Value = Target.Value ' Note (Column C)
    End If

ExitHandler:
    Application.EnableEvents = True ' Re-enable events
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub
```

---

## **Excel Functions for CRM Optimization** 

- **VLOOKUP and Index-Match** for Quick Data Retrieval  
- **IF Statements** for Automated Loan Status Updates  
- **Conditional Formatting** for Flagging Overdue Loans  

---

## **Use Cases**

- **Portfolio Management** – Monitor financial transactions and debt repayment.  
- **Sales Pipeline Management** – Track customer interactions in an automated Excel CRM.  
- **Client Feedback Tracking** – Maintain a structured client database with timestamps.  
- **Organizational Records Management** – Store and retrieve historical CRM data efficiently.  
- **Budget Management** – Automate financial projections using VBA and Excel formulas.  

---

## **Future Enhancements**

- **Integration with Google Sheets** – Sync data between Excel and cloud platforms.  
- **Power BI Dashboard** – Visualize CRM and loan data using Power BI for deeper insights.  
- **User-Friendly Interface** – Add form-based input for non-technical users.  
- **Advanced Reporting & Analytics** – Implement trend analysis using VBA and Excel formulas.

---

## **Contact & Contributions**  
Feel free to explore and contribute. If you have any suggestions, reach out or submit a pull request.  

- **Email**: [jordan.c.l.wright@gmail.com](mailto:jordan.c.l.wright@gmail.com)  

---

### **Author:** Jordan  
[GitHub Profile](https://github.com/JordanConallLuthaisWright)



