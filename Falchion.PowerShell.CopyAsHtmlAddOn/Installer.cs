using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;

namespace Falchion.PowerShell.CopyAsHtmlAddOn
{
    [System.ComponentModel.RunInstallerAttribute(true)]
    public class Installer : System.Configuration.Install.Installer
    {
        public override void Install(System.Collections.IDictionary stateSaver)
        {
            FileInfo assembly = new FileInfo(Context.Parameters["assemblypath"]);

            if (Environment.Is64BitOperatingSystem)
            {
                string x64Path = Environment.ExpandEnvironmentVariables("%windir%\\sysnative\\WindowsPowerShell\\v1.0\\");
                string x86Path = Environment.ExpandEnvironmentVariables("%windir%\\SYSWOW64\\WindowsPowerShell\\v1.0\\");

                AddScript(x64Path, assembly.DirectoryName);
                AddScript(x86Path, assembly.DirectoryName);
            }
            else
            {
                string x86Path = Environment.ExpandEnvironmentVariables("%windir%\\system32\\WindowsPowerShell\\v1.0\\");

                AddScript(x86Path, assembly.DirectoryName);
            }
            base.Install(stateSaver);

        }
        private void RemoveScript()
        {
            FileInfo assembly = new FileInfo(Context.Parameters["assemblypath"]);
            if (Environment.Is64BitOperatingSystem)
            {
                string x64Path = Environment.ExpandEnvironmentVariables("%windir%\\sysnative\\WindowsPowerShell\\v1.0\\");
                string x86Path = Environment.ExpandEnvironmentVariables("%windir%\\SYSWOW64\\WindowsPowerShell\\v1.0\\");

                RemoveScript(x64Path, assembly.DirectoryName);
                RemoveScript(x86Path, assembly.DirectoryName);
            }
            else
            {
                string x86Path = Environment.ExpandEnvironmentVariables("%windir%\\system32\\WindowsPowerShell\\v1.0\\");

                RemoveScript(x86Path, assembly.DirectoryName);
            }
        }
        private void RemoveScript(string powerShellPath, string assemblyPath)
        {
            string fileName = "Microsoft.PowerShellISE_profile.ps1";
            if (Directory.Exists(powerShellPath))
            {
                string[] profileText = null;
                fileName = Path.Combine(powerShellPath, fileName);
                if (File.Exists(fileName))
                {
                    profileText = File.ReadAllLines(fileName);
                }

                string lines = string.Empty;
                bool removed = false;
                if (profileText != null && profileText.Length > 0)
                {
                    foreach (string line in profileText)
                    {
                        if (!string.IsNullOrEmpty(line) && line.ToLower().Contains("falchion.powershell.copyashtmladdon.ps1"))
                        {
                            removed = true;
                            continue;
                        }
                        lines += line + "\r\n";
                    }
                }
                if (removed)
                    File.WriteAllText(fileName, lines);
            }
        }
        private void AddScript(string powerShellPath, string assemblyPath)
        {
            RemoveScript(powerShellPath, assemblyPath);

            string fileName = "Microsoft.PowerShellISE_profile.ps1";
            if (Directory.Exists(powerShellPath))
            {
                string profileText = null;
                fileName = Path.Combine(powerShellPath, fileName);
                if (File.Exists(fileName))
                {
                    profileText = File.ReadAllText(fileName);
                }
                if (string.IsNullOrEmpty(profileText) || !profileText.ToLower().Contains("falchion.powershell.copyashtmladdon.ps1"))
                {
                    string ps1File = Path.Combine(assemblyPath, "Falchion.PowerShell.CopyAsHtmlAddOn.ps1");
                    profileText = ".\"" + ps1File + "\"\r\n" + profileText;
                }
                File.WriteAllText(fileName, profileText);
            }
        }
        public override void Uninstall(System.Collections.IDictionary savedState)
        {
            RemoveScript();

            string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (string.IsNullOrEmpty(path)) return;

            foreach (string subDir in Directory.GetDirectories(path))
            {
                try { Directory.Delete(Path.Combine(path, subDir)); }
                catch { }
            }
            foreach (string file in Directory.GetFiles(path))
            {
                try { File.Delete(Path.Combine(path, file)); }
                catch { }
            }
            base.Uninstall(savedState);
        }
        protected override void OnAfterInstall(System.Collections.IDictionary savedState)
        {

            base.OnAfterInstall(savedState);
        }

    }
}
