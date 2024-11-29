using ClosedXML.Excel;
using System;
using System.Windows.Forms;

namespace vacation
{
    public partial class Form1 : Form
    {
        private string leaveStatus;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            string employeeName = txtEmployeeName.Text;
            string employeePosition = txtEmployeePosition.Text;
            DateTime startDate = dtpStartDate.Value;
            DateTime endDate = dtpEndDate.Value;

            if (!string.IsNullOrEmpty(employeeName) && !string.IsNullOrEmpty(employeePosition) && startDate <= endDate)
            {
                leaveStatus = "Затверджено";
                lblStatus.Text = "Заявку на відпустку відправлено!";
                lblStatus.ForeColor = System.Drawing.Color.Green;
            }
            else
            {
                leaveStatus = "Відхилено"; // статус1
                lblStatus.Text = "Помилка! Перевірте введені дані.";
                lblStatus.ForeColor = System.Drawing.Color.Red;
            }
        }

        private void btnDeny_Click(object sender, EventArgs e)
        {
            leaveStatus = "Відхилено"; // статус2
            lblStatus.Text = "Відмовлено у відпустці.";
            lblStatus.ForeColor = System.Drawing.Color.Red;
        }

        private void btnSaveToExcel_Click(object sender, EventArgs e)
        {
            string employeeName = txtEmployeeName.Text;
            string employeePosition = txtEmployeePosition.Text;
            DateTime startDate = dtpStartDate.Value;
            DateTime endDate = dtpEndDate.Value;

            // Ексель
            string filePath = @"C:\Users\Gera\Desktop\Курсова робота\vacation\Excel\VacationRequests.xlsx";

            XLWorkbook workbook;
            IXLWorksheet worksheet;

            // перевіка на існування файлу
            if (System.IO.File.Exists(filePath))
            {
                // відкрити
                workbook = new XLWorkbook(filePath);
                worksheet = workbook.Worksheet("Заявки на відпустку");
            }
            else
            {
                // новий файл
                workbook = new XLWorkbook();
                worksheet = workbook.Worksheets.Add("Заявки на відпустку");
                worksheet.Cell(1, 1).Value = "ПІБ співробітника";
                worksheet.Cell(1, 2).Value = "Посада співробітника";
                worksheet.Cell(1, 3).Value = "Дата початку";
                worksheet.Cell(1, 4).Value = "Дата закінчення";
                worksheet.Cell(1, 5).Value = "Статус";
            }

            int lastRow = worksheet.LastRowUsed().RowNumber() + 1;

            worksheet.Cell(lastRow, 1).Value = employeeName;
            worksheet.Cell(lastRow, 2).Value = employeePosition;
            worksheet.Cell(lastRow, 3).Value = startDate.ToShortDateString();
            worksheet.Cell(lastRow, 4).Value = endDate.ToShortDateString();
            worksheet.Cell(lastRow, 5).Value = leaveStatus;

            // Збереження файлу
            workbook.SaveAs(filePath);

            lblStatus.Text = "Дані збережено в Excel!";
            lblStatus.ForeColor = System.Drawing.Color.Green;
        }

    }
}
