# ProductivityFunctions
![ProductivityLogo](https://github.com/BekoSan/ProductivityFunctions/blob/master/ProductivityLogo.png)

This is a list of extension methods that make it easy to work with comon WinForms Controls and classes.

# Release Notes
v1.0.0

PictureBox Methods :

- LoadImageFormFile() // Loads image form image file you select.
- LoadImageFormByteArray() // This will load image form byte array you passed into this method.
- GetByteImage() // Will return a byte array copy of image inside the PictureBox.
- SaveImageInFile() // Show SaveFileDialog to save image by it , you can specify the filter and title as parameters on this method.

DataGridView Methods:

- UpdateDataSource() // This method takes one parameter which is a datasource . Note :It disables the AutoGenerateColumns Property.
- ExportToCVSFile() // Show SaveFileDialog and you can set its title by passing title as parameter.
- ExportToExcelFile() // Same as the last one it takes title as parameter.
- ExportToPDFFile() // Exports DataGridView to pdf file.

ComboBox Methods :

- UpdateDataSource() // It takes 3 parameters 
*dataSource // the data source.
*displayMember // the display member for the ComboBox.
*valueMember // the value member for the ComboBox
Its sets SelectedIndex Property to -1

Form Methods :

- ClearForm() // clears all textboxes and datetimepickers and comboboxes inside the form.

Byte Array Methods :

-GetImageFromByteArray() // it returns a Image Object from the byte array.

Exception Class Methods :

- ExportExceptionData() // Its export important exception data to csv file.

v2.0.0
Added New Functions.

ListView Functions:
- LoadFoldersList() // Takes one parameter which is parentDirectory . its loads all the folders in parentDirectory to list view.
- LoadFoldersListAysnc() // Takes one parameter which is parentDirectory . its loads all the folders in parentDirectory to list view aysnc.
 - LoadFilesList() // Takes one parameter which is parentDirectory . its loads all the files in parentDirectory to list view.
 - LoadFoldersListAysnc() // Takes one parameter which is parentDirectory . its loads all the folders in parentDirectory to list view aysnc.
 - LoadFilesListWithIcons() // Takes one parameter which is parentDirectory . its loads all the files with there icons in parentDirectory to list view.
 - LoadFilesListWithIconsAysnc() // Takes one parameter which is parentDirectory . its loads all the files with there icons in parentDirectory to list view aysnc.
 
 DataGridViewComboBoxCell Methods : 
 - UpdateDataSource() // It takes 3 parameters 
*dataSource // the data source.
*displayMember // the display member for the DataGridViewComboBoxCell.
*valueMember // the value member for the DataGridViewComboBoxCell 
Important Note : Use this on in DataGridView.RowsAdded Event.

ImageList Methods :
- FillImageList() // Takes one parameter icons List List<Icon>. it fills the ImageList.Images Property.

## GetLatest Version From nuget.org

[Productivity Functions Nuget](https://www.nuget.org/packages/BekoSan.ProductivityFunctions/)
