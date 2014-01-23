// Copyright (c) 2014 Michal Borejszo michal@traal.eu
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.

using System;
using System.Diagnostics;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace UpdateChecker
{
    class UpdateChecker
    {
        public static void startAsync(string metaFileURL, string name, string root, bool silent)
        {
            _UpdateChecker uc = new _UpdateChecker();
            uc.rootname = root;
            uc.URL = metaFileURL;
            uc.silent = silent;
            uc.name = name;
            Thread t = new Thread(uc.start);
            t.Start();
        }

        public static void start(string metaFileURL, string name, string root, bool silent)
        {
            _UpdateChecker uc = new _UpdateChecker();
            uc.rootname = root;
            uc.URL = metaFileURL;
            uc.silent = silent;
            uc.name = name;
            uc.start();
        }
    }

    class _UpdateChecker
    {
        public string rootname { get; set; }
        public string name { get; set; }
        public string URL { get; set; }
        public bool silent { get; set; }

        public void start()
        {
            XmlDocument x = new XmlDocument();
            try
            {
                x.Load(this.URL);
                XmlNode root = x.DocumentElement;
                if (root.Name != rootname)
                {
                    throw new XmlException();
                }
                string version = x.SelectSingleNode("descendant::version").InnerText;
                string downloadURL = x.SelectSingleNode("descendant::URL").InnerText;
                if (version == null || downloadURL == null)
                {
                    throw new XmlException();
                }
                if (String.Compare(version, Assembly.GetExecutingAssembly().GetName().Version.ToString()) == 1)
                {
                    DialogResult r = MessageBox.Show("New version of " + name + " is available to download. Would you like to open download website?", "Update available", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    if (r == DialogResult.Yes)
                    {
                        Process.Start(downloadURL);
                    }
                }
                else if (!silent)
                {
                    MessageBox.Show("No new version version available.", "No update available", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                if (!silent)
                {
                    MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
