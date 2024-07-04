using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Windows.Media;
using System.Collections.Generic;
using System.Windows.Controls;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        private List<ComboBox> excelColumnComboBoxes;
        private List<ComboBox> wordColumnComboBoxes;

        public MainWindow()
        {
            InitializeComponent();

            excelColumnComboBoxes = new List<ComboBox>
            {
                excelColumn1, excelColumn2, excelColumn3, excelColumn4, excelColumn5,
                excelColumn6, excelColumn7, excelColumn8, excelColumn9, excelColumn10
            };
            wordColumnComboBoxes = new List<ComboBox>
            {
                wordColumn1, wordColumn2, wordColumn3, wordColumn4, wordColumn5,
                wordColumn6, wordColumn7, wordColumn8, wordColumn9, wordColumn10
            };

            foreach (var comboBox in excelColumnComboBoxes.Concat(wordColumnComboBoxes))
            {
                comboBox.IsEnabled = false;
            }

            numberOfColumnsComboBox.SelectedIndex = 9;
        }

        private void SelectWordFile_Click(object sender, RoutedEventArgs e)
        {
            string wordFilePath = GetSelectedFilePath("Word files (*.docx)|*.docx|All files (*.*)|*.*");

            if (wordFilePath != null)
            {
                wordFilePathTextBox.Text = wordFilePath;
                MessageBox.Show("Word файл найден!");
                LoadWordColumns(wordFilePath);
            }
        }

        private void SelectExcelFile_Click(object sender, RoutedEventArgs e)
        {
            string excelFilePath = GetSelectedFilePath("Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*");

            if (excelFilePath != null)
            {
                excelFilePathTextBox.Text = excelFilePath;
                MessageBox.Show("Excel файл найден!");

                sheetComboBox.Items.Clear();

                try
                {
                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
                    {
                        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                        foreach (Sheet sheet in workbookPart.Workbook.Sheets)
                        {
                            sheetComboBox.Items.Add(sheet.Name);
                        }

                        sheetComboBox.SelectedIndex = 0;
                        LoadColumnsForSelectedSheet();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при открытии Excel файла: {ex.Message}");
                }
            }
        }

        private void sheetComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadColumnsForSelectedSheet();
        }

        private void LoadColumnsForSelectedSheet()
        {
            string excelFilePath = excelFilePathTextBox.Text;

            if (string.IsNullOrEmpty(excelFilePath) || sheetComboBox.SelectedIndex == -1)
            {
                return;
            }

            string selectedSheetName = sheetComboBox.SelectedItem.ToString();

            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    Sheet selectedSheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == selectedSheetName);
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(selectedSheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                    if (sheetData != null)
                    {
                        foreach (var comboBox in excelColumnComboBoxes)
                        {
                            comboBox.Items.Clear();
                        }

                        Row firstRow = sheetData.Elements<Row>().FirstOrDefault();
                        if (firstRow != null)
                        {
                            foreach (Cell cell in firstRow.Elements<Cell>())
                            {
                                string columnName = GetColumnNameFromCellReference(cell.CellReference.Value);
                                foreach (var comboBox in excelColumnComboBoxes)
                                {
                                    comboBox.Items.Add(columnName);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных листа Excel: {ex.Message}");
            }
        }

        private void LoadWordColumns(string wordFilePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.Commercial;

                // If you use EPPlus in a noncommercial context
                // according to the Polyform Noncommercial license:
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordFilePath, false))
                {
                    MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                    if (mainPart == null)
                        throw new Exception("Невозможно найти основную часть документа");

                    List<string> columnNames = new List<string>();
                    Table table = mainPart.Document.Body.Elements<Table>().FirstOrDefault();

                    if (table != null)
                    {
                        TableRow firstRow = table.Elements<TableRow>().FirstOrDefault();
                        if (firstRow != null)
                        {
                            int colIndex = 1;
                            foreach (TableCell cell in firstRow.Elements<TableCell>())
                            {
                                columnNames.Add($"Column {colIndex}");
                                colIndex++;
                            }

                            foreach (var comboBox in wordColumnComboBoxes)
                            {
                                comboBox.Items.Clear();
                                foreach (var columnName in columnNames)
                                {
                                    comboBox.Items.Add(columnName);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных таблицы Word: {ex.Message}");
            }
        }

        private void numberOfColumnsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int selectedColumns = numberOfColumnsComboBox.SelectedIndex + 1;

            for (int i = 0; i < excelColumnComboBoxes.Count; i++)
            {
                if (i < selectedColumns)
                {
                    excelColumnComboBoxes[i].IsEnabled = true;
                    wordColumnComboBoxes[i].IsEnabled = true;
                }
                else
                {
                    excelColumnComboBoxes[i].IsEnabled = false;
                    wordColumnComboBoxes[i].IsEnabled = false;
                }
            }
        }
        //Рабочая1
        /*private void CopyData_Click(object sender, RoutedEventArgs e)
        {
            string excelFilePath = excelFilePathTextBox.Text;
            string wordFilePath = wordFilePathTextBox.Text;

            if (string.IsNullOrEmpty(excelFilePath) || string.IsNullOrEmpty(wordFilePath))
            {
                MessageBox.Show("Выберите файлы Excel и Word.");
                return;
            }

            int numberOfColumns = numberOfColumnsComboBox.SelectedIndex + 1; // +1 потому что индексы начинаются с нуля

            List<int> excelColumnIndexes = new List<int>();
            List<int> wordColumnIndexes = new List<int>();

            // Получаем выбранные индексы колонок из ComboBox'ов Excel и Word
            for (int i = 0; i < numberOfColumns; i++)
            {
                ComboBox excelComboBox = FindComboBoxByName("excelColumn" + (i + 1));
                ComboBox wordComboBox = FindComboBoxByName("wordColumn" + (i + 1));

                if (excelComboBox != null && wordComboBox != null)
                {
                    int excelIndex = excelComboBox.SelectedIndex;
                    int wordIndex = wordComboBox.SelectedIndex;

                    if (excelIndex == -1 || wordIndex == -1)
                    {
                        MessageBox.Show("Выберите корректные столбцы для Excel и Word.");
                        return;
                    }

                    excelColumnIndexes.Add(excelIndex);
                    wordColumnIndexes.Add(wordIndex);
                }
            }

            try
            {
                int maxRowCount = GetMaxRowCount(excelFilePath);
                
                for (int rowIndex = 1; rowIndex <= maxRowCount; rowIndex++)
                {
                    foreach (var pair in excelColumnIndexes.Zip(wordColumnIndexes, (excelIndex, wordIndex) => new { ExcelIndex = excelIndex, WordIndex = wordIndex }))
                    {
                        int excelIndex = pair.ExcelIndex;
                        int wordIndex = pair.WordIndex;

                        string excelValue = GetCellValue(excelFilePath, rowIndex, excelIndex);
                        SetCellValue(wordFilePath, rowIndex, wordIndex, excelValue);
                    }
                }

                MessageBox.Show("Данные успешно скопированы из Excel в Word.");
            }
            catch (Exception ex)
            {

                MessageBox.Show($"Ошибка при записи данных из Excel в Word файл: {ex.Message}");
            }
        }
*/
        //Рабочая 2
        /*private void CopyData_Click(object sender, RoutedEventArgs e)
        {
            string excelFilePath = excelFilePathTextBox.Text;
            string wordFilePath = wordFilePathTextBox.Text;

            if (string.IsNullOrEmpty(excelFilePath) || string.IsNullOrEmpty(wordFilePath))
            {
                MessageBox.Show("Выберите файлы Excel и Word.");
                return;
            }

            int numberOfColumns = numberOfColumnsComboBox.SelectedIndex + 1; // +1 потому что индексы начинаются с нуля

            List<int> excelColumnIndexes = new List<int>();
            List<int> wordColumnIndexes = new List<int>();

            // Получаем выбранные индексы колонок из ComboBox'ов Excel и Word
            for (int i = 0; i < numberOfColumns; i++)
            {
                ComboBox excelComboBox = FindComboBoxByName("excelColumn" + (i + 1));
                ComboBox wordComboBox = FindComboBoxByName("wordColumn" + (i + 1));

                if (excelComboBox != null && wordComboBox != null)
                {
                    int excelIndex = excelComboBox.SelectedIndex;
                    int wordIndex = wordComboBox.SelectedIndex;

                    if (excelIndex == -1 || wordIndex == -1)
                    {
                        MessageBox.Show("Выберите корректные столбцы для Excel и Word.");
                        return;
                    }

                    excelColumnIndexes.Add(excelIndex + 1); // +1 потому что EPPlus использует индексацию с 1
                    wordColumnIndexes.Add(wordIndex + 1); // +1 потому что Interop использует индексацию с 1
                }
            }

            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

                    Word.Application wordApp = new Word.Application();
                    wordApp.Visible = false;

                    try
                    {
                        Word.Document doc = wordApp.Documents.Open(wordFilePath);
                        Word.Table table = doc.Tables[1];

                        int excelRowCount = worksheet.Dimension.End.Row;
                        int wordRowCount = table.Rows.Count;

                        // Добавление недостающих строк в таблицу Word
                        for (int i = wordRowCount + 1; i <= excelRowCount; i++)
                        {
                            table.Rows.Add();
                        }

                        for (int rowIndex = 2; rowIndex <= excelRowCount; rowIndex++)
                        {
                            for (int columnIndex = 0; columnIndex < excelColumnIndexes.Count; columnIndex++)
                            {
                                string cellValue = worksheet.Cells[rowIndex, excelColumnIndexes[columnIndex]].Value?.ToString();
                                table.Cell(rowIndex, wordColumnIndexes[columnIndex]).Range.Text = cellValue;
                            }
                        }

                        doc.Save();
                        doc.Close();
                        MessageBox.Show("Данные успешно скопированы из Excel в Word.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при записи данных из Excel в Word файл: {ex.Message}");
                    }
                    finally
                    {
                        wordApp.Quit();
                        Marshal.ReleaseComObject(wordApp);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии Excel файла: {ex.Message}");
            }
        }

        private ComboBox FindComboBoxByName(string name)
        {
            return this.FindName(name) as ComboBox;
        }*/
        private void CopyData_Click(object sender, RoutedEventArgs e)
        {
            string excelFilePath = excelFilePathTextBox.Text;
            string wordFilePath = wordFilePathTextBox.Text;

            if (string.IsNullOrEmpty(excelFilePath) || string.IsNullOrEmpty(wordFilePath))
            {
                MessageBox.Show("Выберите файлы Excel и Word.");
                return;
            }

            int numberOfColumns = numberOfColumnsComboBox.SelectedIndex + 1; // +1 потому что индексы начинаются с нуля

            List<int> excelColumnIndexes = new List<int>();
            List<int> wordColumnIndexes = new List<int>();

            // Получаем выбранные индексы колонок из ComboBox'ов Excel и Word
            for (int i = 0; i < numberOfColumns; i++)
            {
                ComboBox excelComboBox = FindComboBoxByName("excelColumn" + (i + 1));
                ComboBox wordComboBox = FindComboBoxByName("wordColumn" + (i + 1));

                if (excelComboBox != null && wordComboBox != null)
                {
                    int excelIndex = excelComboBox.SelectedIndex;
                    int wordIndex = wordComboBox.SelectedIndex;

                    if (excelIndex == -1 || wordIndex == -1)
                    {
                        MessageBox.Show("Выберите корректные столбцы для Excel и Word.");
                        return;
                    }

                    excelColumnIndexes.Add(excelIndex + 1); // +1 потому что EPPlus использует индексацию с 1
                    wordColumnIndexes.Add(wordIndex + 1); // +1 потому что Interop использует индексацию с 1
                }
            }

            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

                    Word.Application wordApp = new Word.Application();
                    wordApp.Visible = false;

                    try
                    {
                        Word.Document doc = wordApp.Documents.Open(wordFilePath);
                        Word.Table table = doc.Tables[1];

                        int excelRowCount = worksheet.Dimension.End.Row;
                        int wordRowCount = table.Rows.Count;

                        // Добавление недостающих строк в таблицу Word
                        for (int i = wordRowCount + 1; i <= excelRowCount; i++)
                        {
                            table.Rows.Add();
                        }

                        for (int rowIndex = 2; rowIndex <= excelRowCount; rowIndex++)
                        {
                            for (int columnIndex = 0; columnIndex < excelColumnIndexes.Count; columnIndex++)
                            {
                                string cellValue = worksheet.Cells[rowIndex, excelColumnIndexes[columnIndex]].Value?.ToString();
                                table.Cell(rowIndex, wordColumnIndexes[columnIndex]).Range.Text = cellValue;

                                // Устанавливаем границы для ячейки
                                Word.Cell cell = table.Cell(rowIndex, wordColumnIndexes[columnIndex]);
                                cell.Borders.Enable = 1; // 1 для отображения всех границ
                            }
                        }

                        doc.Save();
                        doc.Close();
                        MessageBox.Show("Данные успешно скопированы из Excel в Word.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при записи данных из Excel в Word файл: {ex.Message}");
                    }
                    finally
                    {
                        wordApp.Quit();
                        Marshal.ReleaseComObject(wordApp);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии Excel файла: {ex.Message}");
            }
        }

        private ComboBox FindComboBoxByName(string name)
        {
            return this.FindName(name) as ComboBox;
        }


        // Вспомогательный метод для поиска ComboBox по имени РАБОЧИЙ
        /*private ComboBox FindComboBoxByName(string name)
        {
            var fieldInfo = GetType().GetField(name, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            if (fieldInfo != null)
            {
                return (ComboBox)fieldInfo.GetValue(this);
            }
            return null;
        }*/

        private string GetCellValue(string excelFilePath, int rowIndex, int columnIndex)
        {
            // Создаем объект приложения Excel
            Excel.Application excelApp = new Excel.Application();
           
            // Открываем книгу Excel
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            // Получаем доступ к первому листу (индексация начинается с 1)
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            // Считываем значение ячейки
            string cellValue = ((Excel.Range)worksheet.Cells[rowIndex, columnIndex + 1]).Value?.ToString(); // columnIndex + 1 потому что индексация в Excel начинается с 1
           
            // Закрываем книгу Excel
            workbook.Close(false);
            // Закрываем приложение Excel
            excelApp.Quit();

            // Освобождаем ресурсы
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);
            
            return cellValue;
        }

        private void SetCellValue(string wordFilePath, int rowIndex, int columnIndex, string cellValue)
        {
            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordFilePath, true))
                {
                    Table table = wordDoc.MainDocumentPart.Document.Body.Elements<Table>().FirstOrDefault();
                    TableRow row = table.Elements<TableRow>().ElementAt(rowIndex - 1); // rowIndex - 1 потому что индексация начинается с 0
                    TableCell cell = row.Elements<TableCell>().ElementAt(columnIndex);

                    // Очищаем содержимое ячейки
                    cell.RemoveAllChildren<Paragraph>();

                    // Создаем новый параграф с текстом из Excel
                    Paragraph para = new Paragraph(new Run(new Text(cellValue)));

                    // Добавляем параграф в ячейку
                    cell.Append(para);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при установке значения ячейки: {ex.Message}");
            }
        }

        private int GetMaxRowCount(string excelFilePath)
        {
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    Sheet selectedSheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == sheetComboBox.SelectedItem.ToString());
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(selectedSheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                    if (sheetData != null)
                    {
                        return sheetData.Elements<Row>().Count();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при определении количества строк в Excel: {ex.Message}");
            }

            return 0;
        }

        private string GetSelectedFilePath(string filter)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = filter;
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (openFileDialog.ShowDialog() == true)
            {
                return openFileDialog.FileName;
            }

            return null;
        }

        // Вспомогательный метод для получения имени столбца из ссылки на ячейку Excel (например, "A1" -> "A")
        private string GetColumnNameFromCellReference(string cellReference)
        {
            // Поиск первой цифры в строке
            int firstDigitIndex = cellReference.IndexOfAny("0123456789".ToCharArray());

            // Вырезаем подстроку до первой цифры
            return cellReference.Substring(0, firstDigitIndex);
        }
    }
}
