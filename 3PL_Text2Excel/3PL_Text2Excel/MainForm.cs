using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Drawing;

using System.Windows.Forms;


namespace _3PL_Text2Excel
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void sourceFolderButton_Click(object sender, EventArgs e)
        {
            sourceFolderTextBox.Text = SelectFolder();
        }

        private void outputFolderButton_Click(object sender, EventArgs e)
        {
            outputFolderTextBox.Text = SelectFolder();
        }

        private void startButton_Click(object sender, EventArgs e)
        {
            toolStripStatusLabel.Text = "";

            string sourceFolder = sourceFolderTextBox.Text.Trim();
            string outputFolder = outputFolderTextBox.Text.Trim();

            if (sourceFolder.Length == 0)
            {
                MessageBox.Show("Please select a source folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (outputFolder.Length == 0)
            {
                MessageBox.Show("Please select an output folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            FileHandler handler = new FileHandler(sourceFolder, outputFolder);
            handler.Process();

            toolStripStatusLabel.Text = "Done!";
        }

        private string SelectFolder()
        {
            FolderBrowserDialog folderbrowser = new FolderBrowserDialog();
            folderbrowser.RootFolder = Environment.SpecialFolder.MyComputer;
            folderbrowser.SelectedPath = @"C:\";
            folderbrowser.ShowNewFolderButton = true;

            if (folderbrowser.ShowDialog() == DialogResult.OK)
            {
                return folderbrowser.SelectedPath;
            }

            return String.Empty;
        }        
    }
}
