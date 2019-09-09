using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Threading.Tasks;

namespace ConsoleAppProjectServerSSEC
{
    [RunInstaller(true)]
    public partial class InstallerConsola : System.Configuration.Install.Installer
    {
        public InstallerConsola()
        {
            InitializeComponent();
        }
    }
}
