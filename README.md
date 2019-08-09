# ProductivityFunctions

This is an extension  functions library for  PictureBox , DataGridView ,Form and Byte Array and Exceiption Class 

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
