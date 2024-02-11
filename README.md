# Technical Document: ASP.NET Core Excel Spreadsheet Manipulation

## 1. Introduction

This ASP.NET Core application provides a user-friendly interface for uploading Excel files, viewing their contents, and performing operations such as updating, adding, and clearing rows.

### Technologies Used:
- **ASP.NET Core MVC:** Provides the framework for building web applications.
- **OfficeOpenXml library:** Allows manipulation of Excel files within the application.

## 2. Code Structure

### 2.1 HomeController

#### 2.1.1 Actions

```csharp
public class HomeController : Controller
{
    // Renders the main view
    public ActionResult Index() { }

    // Handles file upload and displays Excel data
    [HttpPost]
    public ActionResult Upload(IFormFile file) { }

    // Updates the entire Excel data
    [HttpPost]
    public ActionResult UpdateExcel(List<List<string>> excelData) { }

    // Adds a new row to the Excel data
    [HttpPost]
    public ActionResult AddRow(List<string> newRow) { }

    // Clears a specific row in the Excel data
    [HttpPost]
    public ActionResult ClearRow(int rowIndex) { }
}
```

### 2.2 Excel Operations

#### 2.2.1 ReadExcel

```csharp
// Reads the contents of the Excel file
private List<List<string>> ReadExcel(string filePath) { }
```

This method reads an Excel file from the specified file path and returns a List of Lists containing the Excel data.

#### 2.2.2 UpdateExcel

```csharp
// Updates the entire Excel data with the provided data
public ActionResult UpdateExcel(List<List<string>> excelData) { }
```

This action updates the entire Excel data with the provided data and renders the "Index" view with the updated data.

#### 2.2.3 AddRow

```csharp
// Adds a new row to the Excel data
[HttpPost]
public ActionResult AddRow(List<string> newRow) { }
```

This action adds a new row to the Excel data, appends it to the existing data, and renders the "Index" view with the added row.

#### 2.2.4 ClearRow

```csharp
// Clears a specific row in the Excel data
[HttpPost]
public ActionResult ClearRow(int rowIndex) { }
```

This action clears the values in a specific row of the Excel data, replaces them with empty strings, and renders the "Index" view with the cleared row.

## 3. View (Index.cshtml)

### 3.1 File Upload

```html
@using (Html.BeginForm("Upload", "Home", FormMethod.Post, new { enctype = "multipart/form-data" }))
{
    <div>
        <label for="file">Choose Excel File:</label>
        <input type="file" name="file" id="file" />
        <input type="submit" value="Upload" />
    </div>
}
```

This form allows users to upload an Excel file.

### 3.2 Excel Data Display

```html
@if (ViewBag.ExcelData != null)
{
    <table border="1" cellpadding="5">
        @for (var i = 0; i < ViewBag.ExcelData.Count; i++)
        {
            <tr>
                @for (var j = 0; j < ViewBag.ExcelData[i].Count; j++)
                {
                    <td>
                        <input type="hidden" name="excelData[@i][@j]" value="@ViewBag.ExcelData[i][j]" />
                        <input type="text" name="UpdateExcelcell[@i][@j]" value="@ViewBag.ExcelData[i][j]" />
                    </td>
                }
                <td>
                    <button type="submit" name="clearRowIndex" value="@i">Clear</button>
                </td>
            </tr>
        }
    </table>
}
```

This code snippet displays the Excel data in a table format. Each cell is editable with input fields, and there's a "Clear" button for each row.

### 3.3 Form Actions

```html
<form method="post" action="@Url.Action("UpdateExcel", "Home")">
    <!-- Form actions for adding, clearing, and saving changes -->
    <button type="submit" name="addRow" value="true">Add Row</button>
    <button type="submit">Save Changes</button>
</form>
```

These buttons allow users to add a new row, clear a row, and save changes to the Excel data, respectively.

## 4. Scenarios and Use Cases

### 4.1 File Upload

- **Scenario:** User uploads an Excel file.
- **Expected Outcome:** Excel data is displayed in the table.

### 4.2 Update Excel Data

- **Scenario:** User edits cell values and clicks "

Save Changes."
- **Expected Outcome:** Entire Excel data is updated with the modified values.

### 4.3 Add Row

- **Scenario:** User enters values in a new row and clicks "Add Row."
- **Expected Outcome:** A new row is added to the Excel data.

### 4.4 Clear Row

- **Scenario:** User clicks "Clear" on a specific row.
- **Expected Outcome:** Values in the selected row are cleared.

## 5. Conclusion

This ASP.NET Core application offers a robust solution for Excel file manipulation with a user-friendly interface. The code structure is modular and well-documented, making it easy to understand, maintain, and extend for future enhancements.
