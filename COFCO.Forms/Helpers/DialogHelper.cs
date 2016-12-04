using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace COFCO.Forms.Helpers
{
    internal class DialogHelper
    {
        private static OpenFileDialog _fileDialog = new OpenFileDialog();
        private static FolderBrowserDialog _folderBrowserDialog = new FolderBrowserDialog();

        /// <summary>
        /// Opens file chooser dialog, and put path to Textbox
        /// </summary>
        /// <param name="textBox">textbox to which will be written file path</param>
        internal static string OpenChooseFileDialog(TextBox textBox)
        {
            var dialogResult = _fileDialog.ShowDialog();

            switch (dialogResult)
            {
                case DialogResult.OK:
                    textBox.Text = _fileDialog.FileName;
                    return _fileDialog.FileName;
            }

            _fileDialog.Reset();
            return string.Empty;
        }

        /// <summary>
        /// Opens folder chooser dialog, and put path to Textbox
        /// </summary>
        /// <param name="textBox">textbox to which will be written file path</param>
        internal static string OpenChooseFolderDialog(TextBox textBox)
        {
            var dialogResult = _folderBrowserDialog.ShowDialog();

            switch (dialogResult)
            {
                case DialogResult.OK:
                    textBox.Text = _folderBrowserDialog.SelectedPath;
                    return _folderBrowserDialog.SelectedPath;
            }

            _folderBrowserDialog.Reset();
            return string.Empty;
        }
    }
}
