using System;
using System.IO;
using System.Windows.Forms;

namespace IfcToExcelWinForms
{
    public class MainForm : Form
    {
        public MainForm()
        {
            Text = "Connection JSON → Excel (IDEA StatiCa Parser)";
            Width = 520;
            Height = 180;

            var btn = new Button { Text = "Select JSON and Convert…", Width = 220, Height = 35, Left = 20, Top = 20 };
            var lbl = new Label { Text = "No file selected.", AutoSize = true, Left = 20, Top = 70 };

            btn.Click += (s, e) =>
            {
                using var ofd = new OpenFileDialog
                {
                    Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                    Title = "Select a connection JSON exported from IDEA StatiCa"
                };

                if (ofd.ShowDialog(this) != DialogResult.OK) return;

                lbl.Text = Path.GetFileName(ofd.FileName);

                using var sfd = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    Title = "Save output Excel",
                    FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + "_parsed.xlsx"
                };

                if (sfd.ShowDialog(this) != DialogResult.OK) return;

                try
                {
                    JsonToExcelConverter.Convert(ofd.FileName, sfd.FileName);
                    MessageBox.Show(this, "Done!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.ToString(), "Conversion failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            Controls.Add(btn);
            Controls.Add(lbl);
        }
    }
}