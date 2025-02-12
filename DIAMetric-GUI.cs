using System;
using System.Diagnostics;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace DIAMetricGUI {
    public class WinFormExample : Form {

        private Button   SetCWD;
	private TextBox  CWDText;
	private string   CWD = "";
	private GroupBox Checkboxes;
	private CheckBox DIANNBox;
	private CheckBox CCBox;
	private CheckBox DNBox;
	private Button   Launch;

        public WinFormExample() {
            DisplayGUI();
        }

        private void DisplayGUI() {
            this.Name = "DIAMetric GUI";
            this.Text = "DIAMetric GUI";
            this.Size = new Size(640,240);
            this.StartPosition = FormStartPosition.CenterScreen;

	    CWDText= new TextBox();
	    CWDText.Name = "CWDText";
	    CWDText.Text = "Directory containing mzMLs not set.";
	    CWDText.ReadOnly = true;
            SetCWD = new Button();
            SetCWD.Name = "SetCWDButton";
            SetCWD.Text = "Select working directory";
            SetCWD.Size = new Size(150, 30);
            SetCWD.Location = new Point(50,50);
            SetCWD.Click += new System.EventHandler(this.SetCWDClick);
	    Launch = new Button();
	    Launch.Name = "LaunchButton";
	    Launch.Text = "Launch DIAMetric";
	    Launch.Size = new Size(150, 30);
	    Launch.Location = new Point(50,250);
	    Launch.Click += new System.EventHandler(this.LaunchClick);
	    Checkboxes = new GroupBox();
	    DIANNBox = new CheckBox();
	    DIANNBox.Name = "DIANNBox";
	    DIANNBox.Text = "Incorporate DIA-NN results";
	    CCBox = new CheckBox();
	    CCBox.Name = "CCBox";
	    CCBox.Text = "(currently unused)";
	    DNBox = new CheckBox();
	    DNBox.Name = "DNBox";
	    DNBox.Text = "(currently unused)";
	    CWDText.Size = new Size(600, 30);
	    DIANNBox.Size = new Size(600, 30);
	    CCBox.Size = new Size(600, 30);
	    DNBox.Size = new Size(600, 30);
	    CWDText.Location = new Point(15, 40);
	    DIANNBox.Location = new Point(15, 70);
	    CCBox.Location = new Point(15, 100);
	    DNBox.Location = new Point(15, 130);

	    Checkboxes.Text = "Options";
	    Checkboxes.Controls.Add(CWDText);
	    Checkboxes.Controls.Add(DIANNBox);
	    Checkboxes.Controls.Add(CCBox);
	    Checkboxes.Controls.Add(DNBox);
	    SetCWD.Dock = DockStyle.Top;
	    Launch.Dock = DockStyle.Bottom;
	    Checkboxes.Dock = DockStyle.Fill;
            this.Controls.Add(SetCWD);
	    this.Controls.Add(Checkboxes);
	    this.Controls.Add(Launch);
        }

        private void SetCWDClick(object source, EventArgs e) {
	    var Browser = new FolderBrowserDialog();
	    if (Browser.ShowDialog() == DialogResult.OK)
	    {
		CWD = Browser.SelectedPath;
		CWDText.Text = "Current path: " + CWD;
	    }
	}

        private void LaunchClick(object source, EventArgs e) {
	    if (CWD == "")
	    {
		MessageBox.Show("Please select a working directory before launching DIAMetric.");
	    }
	    else {
		Process ExternalProcess = new Process();
		string CommandLineOptions = " ";
		if (DIANNBox.Checked)
		{
		    CommandLineOptions = CommandLineOptions + "--DIANN ";
		}
		if (CCBox.Checked)
		{
		    CommandLineOptions = CommandLineOptions + "--CC ";
		}
		if (DNBox.Checked)
		{
		    CommandLineOptions = CommandLineOptions + "--DN ";
		}
		string DIAMetricPath = Environment.CurrentDirectory + "\\DIAMetric.exe";
		ExternalProcess.StartInfo.FileName = DIAMetricPath;
		ExternalProcess.StartInfo.WorkingDirectory = CWD;
		ExternalProcess.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
		ExternalProcess.StartInfo.Arguments = CommandLineOptions;
		ExternalProcess.Start();
		ExternalProcess.WaitForExit();
	    }
        }

	[STAThread]
        public static void Main(String[] args) {
            Application.Run(new WinFormExample());
        }
    }
}
