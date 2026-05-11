# 🏥 Hospital Analysis Dashboard

A Power BI project that provides a comprehensive analysis of hospital operations including patient records, doctor performance, billing, medicine tracking, and appointment management.

---

## 📊 Dashboard Preview

### Patient Dashboard
![Patient Dashboard]("C:\Users\Shrikant\OneDrive\Pictures\Screenshots\Screenshot 2026-05-12 000632.png")

### Doctor Dashboard
![Doctor Dashboard](screenshots/doctor_dashboard.png)

---

## 🛠️ Tools Used

- **Power BI Desktop** — Dashboard creation and data visualization
- **Microsoft Excel** — Data source files
- **Power Query (M Language)** — Data cleaning and transformation
- **DAX (Data Analysis Expressions)** — Measures and calculated columns

---

## 📁 Project Structure

```
Hospital-Analysis-Dashboard/
│
├── Hospital_Dashboard.pbix        ← Main Power BI file
│
├── data/
│   ├── patient_info.xlsx          ← Patient & doctor details
│   ├── appointments.xlsx          ← Appointment records
│   ├── bills.xlsx                 ← Billing & charges
│   ├── medicine_patient.xlsx      ← Medicines given to patients
│   ├── medical_stock_info.xlsx    ← Medicine inventory
│   └── patient_tests.xlsx         ← Lab test records
│
├── screenshots/
│   ├── patient_dashboard.png
│   └── doctor_dashboard.png
│
└── README.md
```

---

## 🗂️ Data Sources

| File | Table Name in Power BI | Description |
|---|---|---|
| `patient_info.xlsx` | `Patient_info` | Patient demographics, admission/discharge dates, bed, room, department, doctor info |
| `appointments.xlsx` | `Appoinment` | Appointment records with date, time, diagnosis, fees, discount, doctor ID |
| `bills.xlsx` | `Bills` | Billing breakdown by charge type (Surgery, Test, Room, Medicine, Discount) |
| `medicine_patient.xlsx` | `Medicine_patient` | Medicines dispensed to each patient with quantity and date |
| `medical_stock_info.xlsx` | `Medical_stocks` | Medicine inventory — batch number, category, cost price, unit price, stock qty, expiry date |
| `patient_tests.xlsx` | `Patient_test` | Lab test results with amount, result (Normal/Abnormal), payment method, discount |

---

## 🧹 Data Preparation Steps

### 1. Data Loading
- Loaded all 6 Excel files into Power BI using **Get Data → Excel**
- Each file was loaded as a separate query and renamed to match the model table names

### 2. Data Cleaning (Power Query)
- **Patient_info** — Removed redundant merged columns; kept only patient-specific columns
- **Appoinment** — Created a combined `Appoinment_date_time` column by merging `appointment_date` and `appointment_time`; converted Excel serial number in `upload_timestamp` to proper date
- **Medicine_patient** — Converted `date` column from Excel serial number format to proper date using `Date.FromOADate()`; converted `upload_timestamp` similarly
- **Patient_test** — Replaced null values in `discount` column with `0`; converted `upload_timestamp` to date format
- **Bills** — Filtered out rows where `Value = 0` to remove empty/unused charge entries
- **Medical_stocks** — Verified `expiry_date` and `manufacture_date` columns were correctly typed as Date

### 3. Calendar Table
- Created a dedicated **`Calender`** table using DAX:

```DAX
Calender = CALENDAR(MIN(Medicine_patient[date]), MAX(Medicine_patient[date]))
```

Added columns:
```DAX
Calender = CALENDAR(DATE(2013, 1, 1), DATE(2025, 5, 31))
Day = FORMAT(Calender[Date], "Ddd")
Month = FORMAT(Calender[Date], "MMM")
Month_index = MONTH(Calender[Date])
Month_year = FORMAT(Calender[Date], "MMM-YYYY")
Only_date = FORMAT([date],"DD-MMM-YY")
```

### 4. Additional Tables
- **`Measure`** — Dedicated table to store all DAX measures (keeps model organized)
- **`Estimated_patient_`** — What-if parameter table used for the Commission Calculator sliders on the Doctor Dashboard

---

## 🔗 Data Model — Table Relationships

> Exactly as built in Power BI Model View

| From Table | From Column | To Table | To Column | Cardinality |
|---|---|---|---|---|
| `Patient_info` | `patient_id` | `Bills` | `patient_id` | One-to-Many |
| `Patient_info` | `patient_id` | `Appoinment` | `patient_id` | One-to-Many |
| `Patient_info` | `patient_id` | `Medicine_patient` | `patient_id` | One-to-Many |
| `Patient_info` | `patient_id` | `Patient_test` | `patient_id` | One-to-Many |
| `Medicine_patient` | `medicine_id` | `Medical_stocks` | `medicine_id` | Many-to-One |
| `Medicine_patient` | `date` | `Calender` | `Date` | Many-to-One |
| `Appoinment` | `appointment_id` | `Bills` | `bill_id` | One-to-One |

**`Patient_info` is the central table** — all patient-related tables connect through `patient_id`.

---

## 📐 DAX Measures

All measures are stored in the dedicated **`Measure`** table.

---

### 💰 Total Bill Amount
Calculates the **net bill amount after subtracting discounts** from the total charges.
```DAX
Total_bill_amount =
VAR discount = CALCULATE(SUM(Bills[Value]), Bills[Charge_Type] = "discount")
VAR totalamount = SUM(Bills[Value])
RETURN totalamount - discount
```
> Uses two variables — first isolates the discount rows, then subtracts from the gross total so the displayed bill is always the actual amount the patient owes.

---

### 🩺 Doctor Fees
Filters the `Bills` table to return only rows where `Charge_Type = "fees"` — isolating the doctor consultation fee from other charge types like tests or medicine.
```DAX
Doctor_fees = CALCULATE(SUM(Bills[Value]), Bills[Charge_Type] = "fees")
```

---

### 📊 Doctor Commission Rate
A fixed constant measure set at **10%**, used as the base commission rate for all doctors.
```DAX
Doc_Comission_rate = 0.10
```

---

### 💵 Doctor Commission Amount
Multiplies the **net bill amount** by the **commission rate** to calculate how much commission a doctor earns from a patient's total bill.
```DAX
Doc_comission_amt = [Total_bill_amount] * [Doc_Comission_rate]
```
> Depends on `Total_bill_amount` (already discount-adjusted) so the commission is always calculated on the correct net figure.

---

### 🧮 Estimated Commission (What-If Calculator)
Powers the **Commission Calculator** on the Doctor Dashboard. Uses two **Numeric Range Parameters** created via Power BI's *New Parameter* feature to simulate potential commission earnings interactively.

**Parameters created:**

| Parameter | Min | Max | Increment | Purpose |
|---|---|---|---|---|
| `Estimate_Comission_rate` | 0 | 100 | 1 | Commission % slider (e.g. 0% to 100%) |
| `Estimate_patient_spending` | 100 | 100,000 | 1 | Patient bill amount slider (₹100 to ₹1,00,000) |

> Both parameters were created using **Modeling → New Parameter → Numeric Range** in Power BI. This automatically generates a disconnected table and a default measure for each parameter, which are then referenced in the `Estimated_comission` measure below.

```DAX
Estimated_comission =
VAR Patient_amount = MAX(Estimate_patient_spending[Estimate_patient_spending])
VAR Comissinon_rate = MAX(Estimate_Comission_rate[Estimate_Comission_rate])
VAR Comission_percentage = DIVIDE(Comissinon_rate, 100)
RETURN Patient_amount * Comission_percentage
```
> `MAX()` extracts the currently selected slider value from each parameter table. `DIVIDE()` safely converts the commission percentage (e.g. 23) into a decimal (0.23) to calculate the final estimated commission amount.

---

### 💊 Medicine Total Sales
Sums the total quantity of all medicines dispensed across all patients.
```DAX
medicine_total_sales = SUM(Medicine_patient[qty])
```

---

## 📋 Dashboard Features

### 🧑‍⚕️ Page 1 — Patient Dashboard
- **Patient Profile Card** — Photo, name, ID, assigned doctor, diagnosis, Admitted/Discharged status
- **Patient Details Panel** — Gender, age, weight, blood group, address
- **Doctor Details Panel** — Department, qualification, specialization, availability, experience
- **Medicine Calendar Matrix** — Medicine quantity per day-of-week across months (powered by `Calender` table)
- **KPI Cards** — Total bill amount, total medicines given, admission date, discharge date
- **Charges Bar Chart** — Breakdown: Test, Fees, Medicine, Discount, Other
- **Top Medicines Bar Chart** — Most given medicines to the selected patient

### 👨‍⚕️ Page 2 — Doctor Dashboard
- **Doctor Profile Card** — Photo, name, ID, salary, department, qualification, specialization
- **Availability Badge** — Available / On Leave
- **Appointments Panel** — Past and upcoming appointments with time, date, patient name, status icons
- **Commission Calculator** — Two what-if sliders (`Estimate_patient_spending` + `Estimate_Comission_rate`) to simulate earnings
- **KPI Cards** — Total billing, commission earned, patient fees, commission rate %
- **Patient Data Table** — Photo, name, status, gender, age for all patients under selected doctor
- **Feedback** — Latest patient satisfaction comment

---

## 🔗 How to Open This Project

1. Download and install **[Power BI Desktop](https://powerbi.microsoft.com/desktop/)** (free)
2. Clone or download this repository
3. Open `Hospital_Dashboard.pbix` in Power BI Desktop
4. If you see data source errors, go to:
   **Home → Transform Data → Data Source Settings → Change Source**
   and point each file to your local `data/` folder
5. Click **Refresh** to reload all data

---

## 👤 Author

**[Shrikant Mangalam]**  
[www.linkedin.com/in/shrikant-mangalam-75148126a] | [shrikantmangalam2002@gmail.com]

---

## 📌 Notes

- All patient and doctor data is **fictional/sample data** for educational purposes only
- No real personal health information is included
- The table name `Calender` (with this spelling) matches the actual table name used in the Power BI file
