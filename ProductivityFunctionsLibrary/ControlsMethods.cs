using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;

namespace ProductivityFunctionsLibrary
{

    /// <summary>
    /// This class is contains all the functions that been  attched to Controls.
    /// </summary>
    public static class ControlsMethods
    {

        #region Private Helper Methods

        private static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.Message, "Error");
            }
            finally
            {
                GC.Collect();
            }
        }

        #endregion

        #region PictureBox Methods

        /// <summary>
        /// Load Image to picture box from file that user choises.
        /// </summary>
        /// <param name="pictureBox">The picture box to load image on.</param>
        /// <param name="filter">The filter to be used to get type of images. example ("PNG Files|*.png|BMP Files|*.bmp")</param>
        /// <param name="fileDialogTitle">The title for the dialog appears for the user.</param>
        public static void LoadImageFromFile(this PictureBox pictureBox, string filter, string fileDialogTitle)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = filter;
            openFileDialog.Title = fileDialogTitle;
            openFileDialog.FileName = "";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                pictureBox.Image = System.Drawing.Image.FromFile(openFileDialog.FileName);
            }

        }

        /// <summary>
        /// Load Image to picture box from a byte array.
        /// </summary>
        /// <param name="pictureBox">The picture box to load image on.</param>
        /// <param name="byteArray">The byte array to get image form.</param>
        public static void LoadImageFromByteArray(this PictureBox pictureBox, byte[] byteArray)
        {

            MemoryStream stream = new MemoryStream(byteArray);
            var img = System.Drawing.Image.FromStream(stream);
            stream.Close();
            pictureBox.Image = img;

        }

        /// <summary>
        /// Gets a byte array copy of the image inside picture box.
        /// </summary>
        /// <param name="pictureBox">The picture box to get byte image form.</param>
        /// <returns></returns>
        public static byte[] GetByteImage(this PictureBox pictureBox)
        {

            byte[] ImgByteArray;
            MemoryStream stream = new MemoryStream();
            pictureBox.Image.Save(stream, ImageFormat.Jpeg); //TODO- Fix gdi+ error
            ImgByteArray = stream.ToArray();
            stream.Close();

            return ImgByteArray;
        }

        /// <summary>
        /// Load Image to picture box from file that user choises.
        /// </summary>
        /// <param name="pictureBox">The picture box to save its image</param>
        /// <param name="filter">The filter to be used to get type of images. example ("PNG Files|*.png|BMP Files|*.bmp")</param>
        /// <param name="saveDialogTitle">The title for the save dialog appears for the user.</param>
        public static void SaveImageInFile(this PictureBox pictureBox, string filter, string saveDialogTitle)
        {

            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Title = saveDialogTitle;
            saveFileDialog.Filter = filter;
            saveFileDialog.FileName = "";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                var s = pictureBox.Size;
                var memoryImage = new Bitmap(s.Width, s.Height);
                var memoryGraphics = Graphics.FromImage(memoryImage);
                var screenPos = pictureBox.PointToScreen(new Point(0, 0));
                memoryGraphics.CopyFromScreen(screenPos.X, screenPos.Y, 0, 0, s);
                memoryImage.Save(saveFileDialog.FileName);
            }

        }

        #endregion

        #region DataGridView Methods

        /// <summary>
        /// Updates the data source of data grid view with disabling auto generate columns.
        /// </summary>
        /// <param name="gridView">The data grid view to update data source.</param>
        /// <param name="dataSource">The data source.</param>
        public static void UpdateDataSource(this DataGridView gridView, object dataSource)
        {
            gridView.AutoGenerateColumns = false;
            gridView.DataSource = null;
            gridView.DataSource = dataSource;
        }

        /// <summary>
        /// Export all the DataGridView data to a csv file.
        /// </summary>
        /// <param name="gridView">The grid view you want to export its contents.</param>
        /// <param name="saveDialogTitle">The title of the save dialog appears to user.</param>
        public static void ExportToCSVFile(this DataGridView gridView, string saveDialogTitle)
        {
            if (gridView.Rows.Count == 0) return;

            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Title = saveDialogTitle;
            saveFileDialog.Filter = "CSV Files|*.csv";
            saveFileDialog.FileName = "";
            bool fileError = false;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {

                if (File.Exists(saveFileDialog.FileName))
                {
                    try
                    {
                        File.Delete(saveFileDialog.FileName);
                    }
                    catch (IOException ex)
                    {
                        fileError = true;
                        MessageBox.Show("can't write the data to the disk." + ex.Message);
                    }
                }
                if (!fileError)
                {
                    try
                    {
                        int columnCount = gridView.Columns.Count;
                        string columnNames = "";
                        string[] outputCsv = new string[gridView.Rows.Count + 1];
                        for (int i = 0; i < columnCount; i++)
                        {
                            columnNames += gridView.Columns[i].HeaderText.ToString() + ",";
                        }
                        outputCsv[0] += columnNames;

                        for (int i = 1; (i - 1) < gridView.Rows.Count; i++)
                        {
                            for (int j = 0; j < columnCount; j++)
                            {
                                outputCsv[i] += gridView.Rows[i - 1].Cells[j].Value.ToString() + ",";
                            }
                        }

                        File.WriteAllLines(saveFileDialog.FileName, outputCsv, Encoding.UTF8);
                        MessageBox.Show("Data Exported Successfully !!!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error :" + ex.Message);
                    }
                }
            }//End of Save file dialog
            else
            {
                MessageBox.Show("No Record To Export !!!", "Info");
            }

        }

        public static void ExportToPDFFile(this DataGridView gridView, string saveDialogTitle)
        {

            if (gridView.Rows.Count == 0) return;

            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Title = saveDialogTitle;
            saveFileDialog.Filter = "CSV Files|*.csv";
            saveFileDialog.FileName = "";
            bool fileError = false;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (File.Exists(saveFileDialog.FileName))
                {
                    try
                    {
                        File.Delete(saveFileDialog.FileName);
                    }
                    catch (IOException ex)
                    {
                        fileError = true;
                        MessageBox.Show("It wasn't possible to write the data to the disk." + ex.Message);
                    }
                }
                if (!fileError)
                {
                    try
                    {
                        PdfPTable pdfTable = new PdfPTable(gridView.Columns.Count);
                        pdfTable.DefaultCell.Padding = 3;
                        pdfTable.WidthPercentage = 100;
                        pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;

                        foreach (DataGridViewColumn column in gridView.Columns)
                        {
                            PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                            pdfTable.AddCell(cell);
                        }

                        foreach (DataGridViewRow row in gridView.Rows)
                        {
                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                pdfTable.AddCell(cell.Value.ToString());
                            }
                        }

                        using (FileStream stream = new FileStream(saveFileDialog.FileName, FileMode.Create))
                        {
                            Document pdfDoc = new Document(PageSize.A4, 10f, 20f, 20f, 10f);
                            PdfWriter.GetInstance(pdfDoc, stream);
                            pdfDoc.Open();
                            pdfDoc.Add(pdfTable);
                            pdfDoc.Close();
                            stream.Close();
                        }

                        MessageBox.Show("Data Exported Successfully !!!", "Info");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error :" + ex.Message);
                    }
                }
            }
        }

        public static void ExportToExcelFile(this DataGridView gridView, string saveDialogTitle)
        {

            if (gridView.Rows.Count > 0)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel (.xlsx)|  *.xlsx";
                saveFileDialog.FileName = "Output.xlsx";
                bool fileError = false;
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(saveFileDialog.FileName))
                    {
                        try
                        {
                            File.Delete(saveFileDialog.FileName);
                        }
                        catch (IOException ex)
                        {
                            fileError = true;
                            MessageBox.Show("It wasn't possible to write the data to the disk." + ex.Message);
                        }
                    }
                    if (!fileError)
                    {
                        try
                        {
                            Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();
                            Microsoft.Office.Interop.Excel._Workbook workbook = XcelApp.Workbooks.Add(Type.Missing);
                            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

                            worksheet = workbook.Sheets["Sheet1"];
                            worksheet = workbook.ActiveSheet;
                            worksheet.Name = "Output";
                            worksheet.Application.ActiveWindow.SplitRow = 1;
                            worksheet.Application.ActiveWindow.FreezePanes = true;

                            for (int i = 1; i < gridView.Columns.Count + 1; i++)
                            {
                                worksheet.Cells[1, i] = gridView.Columns[i - 1].HeaderText;
                                worksheet.Cells[1, i].Font.NAME = "Calibri";
                                worksheet.Cells[1, i].Font.Bold = true;
                                worksheet.Cells[1, i].Interior.Color = Color.Wheat;
                                worksheet.Cells[1, i].Font.Size = 12;
                            }

                            for (int i = 0; i < gridView.Rows.Count; i++)
                            {
                                for (int j = 0; j < gridView.Columns.Count; j++)
                                {
                                    worksheet.Cells[i + 2, j + 1] = gridView.Rows[i].Cells[j].Value.ToString();
                                }
                            }

                            worksheet.Columns.AutoFit();
                            workbook.SaveAs(saveFileDialog.FileName);
                            XcelApp.Quit();

                            ReleaseObject(worksheet);
                            ReleaseObject(workbook);
                            ReleaseObject(XcelApp);

                            MessageBox.Show("Data Exported Successfully !!!", "Info");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error :" + ex.Message);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("No Record To Export !!!", "Info");
            }
        }

        #endregion

        #region ComboBox Methods

        /// <summary>
        /// Updates the combo box data source with value and display members.
        /// </summary>
        /// <param name="comboBox">The combo box.</param>
        /// <param name="dataSource">The Data source.</param>
        /// <param name="displayMember">The display member.</param>
        /// <param name="valueMember">The value memeber.</param>
        public static void UpdateDataSource(this ComboBox comboBox, object dataSource, string displayMember, string valueMember)
        {
            comboBox.DataSource = null;
            comboBox.DataSource = dataSource;
            comboBox.DisplayMember = displayMember;
            comboBox.ValueMember = valueMember;

            comboBox.SelectedIndex = -1;
        }

        #endregion


        #region Form Methods

        public static void ClearForm(this Form form)
        {

            foreach (Control parentControl in form.Controls)
            {

                if (parentControl.HasChildren)
                {
                    foreach (Control childControl in parentControl.Controls)
                    {
                        if (childControl.GetType().ToString().Contains("TextBox"))
                        {
                            ((TextBox)childControl).Text = "";
                        }
                        else if (childControl.GetType().ToString().Contains("DateTimePicker"))
                        {
                            ((DateTimePicker)childControl).Value = DateTime.Today.Date;
                        }
                        else if (childControl.GetType().ToString().Contains("ComboBox"))
                        {
                            ((ComboBox)childControl).SelectedIndex = -1;
                            ((ComboBox)childControl).Text = "";
                        }
                    }
                }
                else
                {
                    if (parentControl.GetType().ToString().Contains("TextBox"))
                    {
                        ((TextBox)parentControl).Text = "";
                    }
                    else if (parentControl.GetType().ToString().Contains("DateTimePicker"))
                    {
                        ((DateTimePicker)parentControl).Value = DateTime.Today.Date;
                    }
                    else if (parentControl.GetType().ToString().Contains("ComboBox"))
                    {
                        ((ComboBox)parentControl).SelectedIndex = -1;
                        ((ComboBox)parentControl).Text = "";
                    }
                }

            }

        }

        #endregion

    }

}



