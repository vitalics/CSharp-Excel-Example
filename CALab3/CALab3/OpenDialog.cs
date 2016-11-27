using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CALab3
{
    class OpenDialog
    {
        private static string[] _data;

        public static void FileChooser()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            string file = null;
            string text = null;

            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.

            if (result == DialogResult.OK) // Test result.
            {
                file = openFileDialog1.FileName;
                try
                {
                    text = File.ReadAllText(file);
                    _data = text.Split('\r', '\n');
                    _data = _data.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                    Console.WriteLine("Write filename");
                    string fileName = Console.ReadLine();
                    MyExcel myExcel = new MyExcel(_data, fileName);
                }
                catch (IOException)
                {
                }
            }
        }

        public static string FolderChooser()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            string path = null;
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                path = fbd.SelectedPath;
                Console.WriteLine(path);

            }
            return path;
        }

    }
}
