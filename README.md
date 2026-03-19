## Project Structure

- `DeleteSheets.bas` → Deletes previously generated sheets to allow regeneration
- `HideShow.bas` → Handles hiding and showing sheets properly
- `Hyperlinks.bas` → Creates navigation links between worksheets
- `IndexTable.bas` → Generates a summary table with sellers and their respective totals and subtotals
- `Recorded.bas` → Contains a recorded macro used for report formatting
- `Report.bas` → Main macro responsible for dynamically generating reports

---

## How to Use

1. Open Excel and press `ALT + F11` to access the VBA editor  
2. Import all `.bas` files into the project  
3. Prepare the required sheets:
   - `Principal` → list of sellers/categories (column A)
   - `Data` → source dataset used for report generation  
4. Ensure the data in `Data` starts from row 1 with headers  
5. Run the `reportation` macro  

---

## Demo

https://youtube.com/SEU-LINK-AQUI

---

## Technologies

- VBA (Visual Basic for Applications)
- Microsoft Excel

---

## Improvements

- Refactor to remove `.Select` usage
- Improve performance
- Add error handling
