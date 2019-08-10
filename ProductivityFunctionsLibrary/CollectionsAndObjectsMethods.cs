using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace ProductivityFunctionsLibrary
{

    /// <summary>
    /// This class is for methods that attached to collections and objects.
    /// </summary>
   public static class CollectionsAndObjectsMethods
    {

        /// <summary>
        /// Gets image form byte array.
        /// </summary>
        /// <param name="byteArray">The byte array.</param>
        public static Image GetImageFromByteArray(byte[] byteArray)
        {

            MemoryStream stream = new MemoryStream(byteArray);
            Bitmap bmp = new Bitmap(stream);
            var img = Image.FromStream(stream);

            stream.Close();
            return img;

        }

        /// <summary>
        /// Exports the exception data to csv file , its get the source and the message and stackTrace of the exciption.
        /// </summary>
        /// <param name="exception">The exception to export it's data.</param>
        public static void ExportExceptionData(this Exception exception)
        {

            List<string> Lines = new List<string>();
            Lines.Add($"{ exception.Source },{ exception.Message  },{ exception.StackTrace }");
            File.WriteAllLines("ExceptionsLog.csv", Lines);

        }

        /*String Object Functions
         - GetFileIcon(fileName); returns Image.
         - GetFileIcon(fileName); returns Icon.
         - GetFileIcon(fileName); returns byte array.
         */

        ///// <summary>
        ///// Gets the icon of the file.
        ///// Important Note : make sure that your string is a file path.
        ///// </summary>
        ///// <param name="fileName">Full path of the file.</param>
        ///// <returns></returns>
        //public static Bitmap GetFileIcon(this string fileName)
        //{

        //    if (fileName.Contains(@"\") != true) return null;
        //    Bitmap bitmap = Icon.ExtractAssociatedIcon(fileName).ToBitmap();
        //    return bitmap;

        //}

    }

}
