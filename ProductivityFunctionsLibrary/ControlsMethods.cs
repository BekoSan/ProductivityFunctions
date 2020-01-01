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

        private static List<string> LoadAllFolders_(string parentDirectory)
        {

            List<string> output = new List<string>();

            if (Directory.Exists(parentDirectory))
            {
                var foldersList = Directory.GetDirectories(parentDirectory);
                if (foldersList.Length != 0)
                {
                    foreach (string directory in foldersList)
                    {
                        output.Add(directory);
                    }
                }
            }

            return output;

        }

        private static async Task<List<string>> LoadAllFolders(string parentDirectory)
        {

            List<string> output = await Task.Run(() => LoadAllFolders_(parentDirectory));
            return output;

        }

        private static List<string> LoadAllFiles_(string parentDirectory)
        {

            List<string> output = new List<string>();

            var filesList = Directory.GetFiles(parentDirectory);

            if (filesList.Length != 0)
            {
                foreach (string file in filesList)
                {
                    output.Add(file);
                }
            }

            return output;

        }

        private static async Task<List<string>> LoadAllFiles(string parentDirectory)
        {

            List<string> output = await Task.Run(() => LoadAllFiles_(parentDirectory));
            return output;

        }

        private static void FillIconsList(this List<string> files,List<Icon> iconsList)
        {
            Icon tempIcon;

            foreach (string file in files)
            {
                tempIcon = Icon.ExtractAssociatedIcon(file);
                iconsList.Add(tempIcon);
            }

        }

        private static async void FillIconsListAysnc(this List<string> files, List<Icon> iconsList)
        {
            Icon tempIcon;

            await Task.Run(() => {
                foreach (string file in files)
                {
                    tempIcon = Icon.ExtractAssociatedIcon(file);
                    iconsList.Add(tempIcon);
                }
            });

        }

        private static List<Icon> LoadAllFileIcons(string parentDirectory)
        {

            List<Icon> output = new List<Icon>();
            List<string> allFiles = LoadAllFiles_(parentDirectory);
            if (allFiles.Count > 0)
            {
                allFiles.FillIconsList(output);
            }

            return output;

        }

        private static async  Task<List<Icon>> LoadAllFileIconsAysnc(string parentDirectory)
        {

            List<Icon> output = new List<Icon>();
            List<string> allFiles = await LoadAllFiles(parentDirectory);
            if (allFiles.Count > 0)
            {
                await Task.Run(() => allFiles.FillIconsList(output));
            }

            return output;

        }

        //private static List<Icon> LoadAllFoldersIcons(string parentDirectory)
        //{

        //    List<Icon> output = new List<Icon>();
        //    List<string> allFolders = Directory.GetFiles(parentDirectory).ToList();
        //    Icon tempIcon;
        //    if (allFolders.Count > 0)
        //    {
        //        foreach (string file in allFolders)
        //        {
        //            tempIcon = Icon.ExtractAssociatedIcon(file);
        //            output.Add(tempIcon);
        //        }
        //    }

        //    return output;

        //}

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

        /*TODO- PictureBox Functions
         - GenerateQRCode(text);
         - GenerateBarcode(text);
        */

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

        #region DataGridViewComboBoxCell Methods

        /// <summary>
        /// Updates the combo box data source with value and display members.
        /// Important Note : Use this method in RowsAdded Event of the DataGridView.
        /// </summary>
        /// <param name="comboBox">The combo box.</param>
        /// <param name="dataSource">The Data source.</param>
        /// <param name="displayMember">The display member.</param>
        /// <param name="valueMember">The value memeber.</param>
        public static void UpdateDataSource(this DataGridViewComboBoxCell comboBox, object dataSource, string displayMember, string valueMember)
        {
            comboBox.DataSource = null;
            comboBox.DataSource = dataSource;
            comboBox.DisplayMember = displayMember;
            comboBox.ValueMember = valueMember;
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

        #region ImageList Methods

        /// <summary>
        /// Fills the ImageList.Images Property.
        /// </summary>
        /// <param name="imageList">The image list control to fill.</param>
        /// <param name="icons">The list of icons to be inserted in the list.</param>
        public static void FillImageList(this ImageList imageList, List<Icon> icons)
        {

            foreach (Icon icon in icons)
            {
                imageList.Images.Add(icon);
            }

        }

        #endregion

        #region ListView Methods

        /// <summary>
        /// Loads all the folders inside parent directory.
        /// </summary>
        /// <param name="listView">The list view to load folders on.</param>
        /// <param name="parentDirectory">Full path of the parent directory.</param>
        public static void LoadFoldersList(this ListView listView, string parentDirectory)
        {

            List<string> ItemsList = LoadAllFolders_(parentDirectory);

            if (ItemsList.Count > 0)
            {
                listView.Items.Clear();
                foreach (string item in ItemsList)
                {
                    listView.Items.Add(item);
                }
            }

        }

        /// <summary>
        /// Loads all the folders inside parent directory aysnc.
        /// </summary>
        /// <param name="listView">The list view to load folders on.</param>
        /// <param name="parentDirectory">Full path of the parent directory.</param>
        public static async void LoadFoldersListAysnc(this ListView listView, string parentDirectory)
        {

            List<string> listViewItems = await LoadAllFolders(parentDirectory);

            if (listViewItems.Count > 0)
            {
                listView.Items.Clear();
                foreach (string item in listViewItems)
                {
                    listView.Items.Add(item);
                }
            }

        }

        /// <summary>
        /// Load all files inside of parent Directory.
        /// </summary>
        /// <param name="listView">The list view to load files on</param>
        /// <param name="parentDirectory">Full path of the parent directory.</param>
        public static void LoadFilesList(this ListView listView, string parentDirectory)
        {

            List<string> ItemsList = LoadAllFiles_(parentDirectory);

            if (ItemsList.Count > 0)
            {
                listView.Items.Clear();
                foreach (string item in ItemsList)
                {
                    listView.Items.Add(item);
                }
            }

        }

        /// <summary>
        /// Load all files inside of parent Directory aysnc.
        /// </summary>
        /// <param name="listView">The list view to load files on</param>
        /// <param name="parentDirectory">Full path of the parent directory.</param>
        public static async void LoadFilesListAysnc(this ListView listView, string parentDirectory)
        {

            List<string> ItemsList = await LoadAllFiles(parentDirectory);

            if (ItemsList.Count > 0)
            {
                listView.Items.Clear();
                foreach (string item in ItemsList)
                {
                    listView.Items.Add(item);
                }
            }

        }

        /// <summary>
        /// Load all files inside of parent Directory.
        /// </summary>
        /// <param name="listView">The list view to load files on</param>
        /// <param name="parentDirectory">Full path of the parent directory.</param>
        public static void LoadFilesWithIcons(this ListView listView, string parentDirectory)
        {

            List<string> ItemsList = LoadAllFiles_(parentDirectory);
            ImageList imageList = new ImageList();
            imageList.FillImageList(LoadAllFileIcons(parentDirectory));

            if (ItemsList.Count > 0)
            {
                listView.Items.Clear();
                //int lastIndex = listView.Items.Count - 1;
                listView.LargeImageList = imageList;
                listView.SmallImageList = imageList;
                foreach (string item in ItemsList)
                {
                    listView.Items.Add(item);
                    listView.Items[listView.Items.Count - 1].ImageIndex = (listView.Items.Count - 1);
                }
            }

        }

        /// <summary>
        /// Load all files inside of parent directoryaysnc.
        /// </summary>
        /// <param name="listView">The list view to load files on</param>
        /// <param name="parentDirectory">Full path of the parent directory.</param>
        public static async void LoadFilesWithIconsAysnc(this ListView listView, string parentDirectory)
        {

            List<string> ItemsList = await LoadAllFiles(parentDirectory);
            ImageList imageList = new ImageList();
            imageList.FillImageList(await LoadAllFileIconsAysnc(parentDirectory));

            if (ItemsList.Count > 0)
            {
                listView.Items.Clear();
                listView.LargeImageList = imageList;
                listView.SmallImageList = imageList;
                foreach (string item in ItemsList)
                {
                    listView.Items.Add(item);
                    listView.Items[listView.Items.Count - 1].ImageIndex = (listView.Items.Count - 1);
                }
            }

        }

        /// <summary>
        /// Load all files and folders inside of parent Directory.
        /// </summary>
        /// <param name="listView">The list view to load files on</param>
        /// <param name="parentDirectory">Full path of the parent directory.</param>
        public static void LoadFilesAndFolders(this ListView listView, string parentDirectory)
        {

            List<string> ItemsList = LoadAllFiles_(parentDirectory);
            ItemsList.AddRange(LoadAllFolders_(parentDirectory));

            if (ItemsList.Count > 0)
            {
                listView.Items.Clear();
                foreach (string item in ItemsList)
                {
                    listView.Items.Add(item);
                }
            }

        }

        /// <summary>
        /// Load all files and folders inside of parent Directory aysnc.
        /// </summary>
        /// <param name="listView">The list view to load files on</param>
        /// <param name="parentDirectory">Full path of the parent directory.</param>
        public static async void LoadFilesAndFoldersAysnc(this ListView listView, string parentDirectory)
        {

            List<string> ItemsList = await LoadAllFiles(parentDirectory);
            ItemsList.AddRange(await LoadAllFolders(parentDirectory));

            if (ItemsList.Count > 0)
            {
                listView.Items.Clear();
                foreach (string item in ItemsList)
                {
                    listView.Items.Add(item);
                }
            }

        }

        /// <summary>
        /// Load all files and folders inside of parent Directory with thier icons.
        /// </summary>
        /// <param name="listView">The list view to load files on</param>
        /// <param name="parentDirectory">Full path of the parent directory.</param>
        public static void LoadFilesAndFoldersWithIcons(this ListView listView, string parentDirectory)
        {

            List<string> ItemsList = LoadAllFiles_(parentDirectory);
            ItemsList.AddRange(LoadAllFolders_(parentDirectory));
            ImageList imageList = new ImageList();
            imageList.FillImageList(LoadAllFileIcons(parentDirectory));

            if (ItemsList.Count > 0)
            {
                listView.Items.Clear();
                listView.LargeImageList = imageList;
                listView.SmallImageList = imageList;
                foreach (string item in ItemsList)
                {
                    listView.Items.Add(item);
                    listView.Items[listView.Items.Count - 1].ImageIndex = (listView.Items.Count - 1);
                }
            }

        }

        /// <summary>
        /// Load all files and folders inside of parent Directory with thier icons aysnc.
        /// </summary>
        /// <param name="listView">The list view to load files on</param>
        /// <param name="parentDirectory">Full path of the parent directory.</param>
        public static async void LoadFilesAndFoldersWithIconsAysnc(this ListView listView, string parentDirectory)
        {

            List<string> ItemsList = await LoadAllFiles(parentDirectory);
            ItemsList.AddRange(await LoadAllFolders(parentDirectory));
            ImageList imageList = new ImageList();
            imageList.FillImageList(await LoadAllFileIconsAysnc(parentDirectory));

            if (ItemsList.Count > 0)
            {
                listView.Items.Clear();
                listView.LargeImageList = imageList;
                listView.SmallImageList = imageList;
                foreach (string item in ItemsList)
                {
                    listView.Items.Add(item);
                    listView.Items[listView.Items.Count - 1].ImageIndex = (listView.Items.Count - 1);
                }
            }

        }

        /*TODO- ListView Functions 
   - LoadList(list<T>);
   - LoadList(list<T>,ImageList images);
   */

        #endregion

        #region TreeView Methods

        /// <summary>
        /// Loads all the folders inside parent directory.
        /// </summary>
        /// <param name="treeView">The tree view to load folders on.</param>
        /// <param name="parentDirectory">Full path of the parent directory.</param>
        public static void LoadFoldersList(this TreeView treeView, string parentDirectory)
        {

            List<string> ItemsList = LoadAllFolders_(parentDirectory);

            if (ItemsList.Count > 0)
            {
                treeView.Nodes.Clear();
                foreach (string item in ItemsList)
                {
                    treeView.Nodes.Add(item);
                }
            }

        }

        /// <summary>
        /// Loads all the folders inside parent directory aysnc.
        /// </summary>
        /// <param name="treeView">The tree view to load folders on.</param>
        /// <param name="parentDirectory">Full path of the parent directory.</param>
        public static async void LoadFoldersListAysnc(this TreeView treeView, string parentDirectory)
        {

            List<string> ItemsList = await  LoadAllFolders(parentDirectory);

            if (ItemsList.Count > 0)
            {
                treeView.Nodes.Clear();
                foreach (string item in ItemsList)
                {
                    treeView.Nodes.Add(item);
                }
            }

        }

        /// <summary>
        /// Loads all the files inside parent directory.
        /// </summary>
        /// <param name="treeView">The tree view to load files on.</param>
        /// <param name="parentDirectory">Full path of the parent directory.</param>
        public static void LoadFilesList(this TreeView treeView, string parentDirectory)
        {

            List<string> ItemsList = LoadAllFiles_(parentDirectory);

            if (ItemsList.Count > 0)
            {
                treeView.Nodes.Clear();
                foreach (string item in ItemsList)
                {
                    treeView.Nodes.Add(item);
                }
            }

        }

        /// <summary>
        /// Loads all the files inside parent directory aysnc.
        /// </summary>
        /// <param name="treeView">The tree view to load files on.</param>
        /// <param name="parentDirectory">Full path of the parent directory.</param>
        public static async void LoadFilesListAysnc(this TreeView treeView, string parentDirectory)
        {

            List<string> ItemsList = await LoadAllFiles(parentDirectory);

            if (ItemsList.Count > 0)
            {
                treeView.Nodes.Clear();
                foreach (string item in ItemsList)
                {
                    treeView.Nodes.Add(item);
                }
            }

        }

        /// <summary>
        /// Loads all the files and folders inside parent directory.
        /// </summary>
        /// <param name="treeView">The tree view to load files and folders on.</param>
        /// <param name="parentDirectory">Full path of the parent directory.</param>
        public static void LoadFilesAndFolders(this TreeView treeView, string parentDirectory)
        {

            List<string> ItemsList = LoadAllFiles_(parentDirectory);
            ItemsList.AddRange(LoadAllFolders_(parentDirectory));

            if (ItemsList.Count > 0)
            {
                treeView.Nodes.Clear();
                foreach (string item in ItemsList)
                {
                    treeView.Nodes.Add(item);
                }
            }

        }

        /// <summary>
        /// Loads all the files and folders inside parent directory aysnc.
        /// </summary>
        /// <param name="treeView">The tree view to load files and folders on.</param>
        /// <param name="parentDirectory">Full path of the parent directory.</param>
        public static async void LoadFilesAndFoldersAysnc(this TreeView treeView, string parentDirectory)
        {

            List<string> ItemsList = await LoadAllFiles(parentDirectory);
            ItemsList.AddRange(await LoadAllFolders(parentDirectory));

            if (ItemsList.Count > 0)
            {
                treeView.Nodes.Clear();
                foreach (string item in ItemsList)
                {
                    treeView.Nodes.Add(item);
                }
            }

        }

        #endregion

        #region RichTextBox Methods

        /// <summary>
        /// Load Text to rich text box from file that user choises.
        /// </summary>
        /// <param name="richTextBox">The rich text box to load file on.</param>
        /// <param name="filter">The filter to be used to get type of document. example ("Doc Files|*.doc|Txt Files|*.txt")</param>
        /// <param name="fileDialogTitle">The title for the dialog appears for the user.</param>
        public static void LoadFile(this RichTextBox richTextBox,string filter,string fileDialogTitle)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = filter;
            openFileDialog.Title = fileDialogTitle;
            openFileDialog.FileName = "";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBox.Text = File.ReadAllText(openFileDialog.FileName);
            }
        }

        #endregion

    }

}



