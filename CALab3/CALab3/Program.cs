using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CALab3
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            //OpenDialog.FileChooser();
            //OpenDialog.FolderChooser();
            OpenDialog.FileChooser();
        }
    }
}
