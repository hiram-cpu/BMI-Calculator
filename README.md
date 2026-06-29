# 📊 VBA & Excel IMC Calculator with Dynamic Lookup

## 📖 Description
This project combines the power of **VBA (Visual Basic for Applications)** automation with native **Microsoft Excel** functions to create an efficient system for recording and diagnosing Body Mass Index (BMI).

The system uses a macro to ensure clean and validated data capture (weight and height), injects the data into the spreadsheet, and leverages Excel's formula engine (`VLOOKUP`) to process the index and automatically classify the user's nutritional status against standard health ranges.

---

## ⚙️ System Architecture
The project operates under a three-layer functional architecture:

### 1. Input and Integrity Layer (VBA)
* **Guided Capture:** Uses `InputBox` pop-ups to request height (in meters) and weight from the user.
* **Data Validation:** The script uses `IsNumeric` to verify that the entered values are strictly quantitative before interacting with the workbook, preventing run-time errors.
* **Precise Injection:** Validated data is transformed into double-precision floating-point numbers (`CDbl`) and inserted directly into the assigned cells (`M17` for height and `M18` for weight).

### 2. Processing Layer (Excel Formulas)
* **Metric Calculation:** The spreadsheet synchronously processes the BMI using the standard mathematical formula:
  $$\text{BMI} = \frac{\text{Weight (kg)}}{\text{Height (m)}^2}$$

### 3. Classification Layer (`VLOOKUP`)
* **Dynamic Indexing:** Once the numerical BMI value is obtained, the `VLOOKUP` function is used with **approximate match (`TRUE` / `1`)**.
* **Reference Matrix:** The value is contrasted against a matrix table parameterized with official WHO ranges (Underweight, Normal, Overweight, Obese), instantly returning the corresponding text diagnosis.

---

## 🚀 Key Features
* **Zero Data Corruption:** Pre-validation in VBA prevents text or special characters from being entered into the calculation cells.
* **Functional Separation:** VBA handles workflow automation exclusively, while Excel manages the mathematical model and logical searches, optimizing file performance.
* **Clean Interface:** Native notifications (`MsgBox`) in `vbInformation` or `vbCritical` format provide feedback based on process success.

---

## 🛠️ Technologies and Functions Used
* **Language:** VBA (Visual Basic for Applications)
* **Key VBA Functions:** `InputBox`, `IsNumeric`, `CDbl`, `Range.Value`, `MsgBox`
* **Key Excel Functions:** `VLOOKUP` (Range lookup with approximate match), Power (`^`).

---

## 📂 Repository Structure
* `IMC.xlsm`: Macro-enabled Excel workbook containing the visual interface, search formulas, and classification tables.
* `BMI.bas`: Exported source code module containing the `Height_and_Weight` subroutine for review or direct import.
