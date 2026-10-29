using System;
using System.IO;
using System.Linq;
using System.Drawing;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

[assembly: System.Reflection.AssemblyTitle("RianFile")]
[assembly: System.Reflection.AssemblyProduct("RianFile")]
[assembly: System.Reflection.AssemblyVersion("1.0.0.0")]
[assembly: System.Reflection.AssemblyFileVersion("1.0.0.0")]

namespace RianFile {
 public class Settings {
  public string Output = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
  public bool Beside, OpenAfter;
  public bool Rename = true;
  public string LibreOffice = "", Renderer = "";
  public int Timeout = 180;
  public static string Location { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RianFile", "settings.xml"); } }
  public static Settings Load() { try { using(var s=File.OpenRead(Location)) return (Settings)new XmlSerializer(typeof(Settings)).Deserialize(s); } catch { return new Settings(); } }
  public void Save() { Directory.CreateDirectory(Path.GetDirectoryName(Location)); using(var s=File.Create(Location)) new XmlSerializer(typeof(Settings)).Serialize(s,this); }
 }
 static class Program {
  [STAThread] static int Main(string[] args) {
   if(args.Length>0 && args[0]=="--worker") {
    try { Engine.Convert(args[1],args[2],args[3],args[4]); return 0; }
    catch(Exception e) { File.WriteAllText(args[2]+".error",e.GetBaseException().Message); return 1; }
   }
   Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false); Application.Run(new MainForm()); return 0;
  }
 }
 public class MainForm : Form {
  Settings settings=Settings.Load(); ListView files=new ListView(); ComboBox format=new ComboBox();
  TextBox output=new TextBox(), lo=new TextBox(), renderer=new TextBox();
  CheckBox beside=new CheckBox(), rename=new CheckBox(), open=new CheckBox(); NumericUpDown timeout=new NumericUpDown();
  ProgressBar progress=new ProgressBar(); Label status=new Label(), note=new Label(); TabControl tabs=new TabControl();
  Button convert; bool busy; string lastFolder;
  public MainForm() {
   Icon=System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);
   Text="RianFile"; ClientSize=new Size(950,650); MinimumSize=new Size(850,620); BackColor=Color.White; Font=new Font("Segoe UI",9);
   var header=new Panel { Dock=DockStyle.Top,Height=77,BackColor=Color.White };
   header.Controls.Add(new Label { Text="RianFile", Font=new Font(Font,FontStyle.Bold),AutoSize=true,Location=new Point(14,15) });
   header.Controls[0].Font=new Font("Segoe UI",18,FontStyle.Bold);
   header.Controls.Add(new Label { Text="Multi-Format File Converter",AutoSize=true,Location=new Point(16,49),ForeColor=Color.DimGray });
   var bottom=new Panel { Dock=DockStyle.Bottom,Height=48,Padding=new Padding(10) };
   status.Dock=DockStyle.Fill; status.Text="Ready"; status.TextAlign=ContentAlignment.MiddleLeft; bottom.Controls.Add(status);
   tabs.Dock=DockStyle.Fill; var page=new TabPage("Converter") { BackColor=Color.White,Padding=new Padding(12) }; var prefs=new TabPage("Settings") { BackColor=Color.White,Padding=new Padding(16) }; tabs.TabPages.AddRange(new[]{page,prefs});
   Controls.Add(tabs); Controls.Add(bottom); Controls.Add(header);
   var layout=new TableLayoutPanel { Dock=DockStyle.Fill,ColumnCount=1,RowCount=6 }; layout.RowStyles.Add(new RowStyle(SizeType.Absolute,38)); layout.RowStyles.Add(new RowStyle(SizeType.Percent,100)); layout.RowStyles.Add(new RowStyle(SizeType.Absolute,92)); layout.RowStyles.Add(new RowStyle(SizeType.Absolute,48)); layout.RowStyles.Add(new RowStyle(SizeType.Absolute,26)); layout.RowStyles.Add(new RowStyle(SizeType.Absolute,43)); page.Controls.Add(layout);
   var tools=new FlowLayoutPanel { Dock=DockStyle.Fill }; tools.Controls.Add(Button("Add Files", delegate { using(var d=new OpenFileDialog { Multiselect=true,Filter="Supported documents|*.docx;*.pptx;*.pdf" }) if(d.ShowDialog()==DialogResult.OK) Add(d.FileNames); }));
   tools.Controls.Add(Button("Remove",delegate { foreach(ListViewItem i in files.SelectedItems) files.Items.Remove(i); RefreshSupport(); })); tools.Controls.Add(Button("Clear",delegate { files.Items.Clear(); RefreshSupport(); })); tools.Controls.Add(new Label { Text="Drag and drop DOCX, PPTX or PDF files below.",AutoSize=true,Margin=new Padding(16,9,0,0) }); layout.Controls.Add(tools,0,0);
   files.Dock=DockStyle.Fill; files.View=View.Details; files.FullRowSelect=true; files.GridLines=true; files.HideSelection=false; files.ShowItemToolTips=true; files.Columns.Add("File",235); files.Columns.Add("Type",65); files.Columns.Add("Size",85); files.Columns.Add("Status",95); files.Columns.Add("Details / output",370); files.AllowDrop=true;
   files.DoubleClick+=delegate {if(files.SelectedItems.Count>0)MessageBox.Show(files.SelectedItems[0].SubItems[4].Text,"File details");};
   files.DragEnter+=delegate(object s,DragEventArgs e) { if(!busy && e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect=DragDropEffects.Copy; }; files.DragDrop+=delegate(object s,DragEventArgs e) { Add((string[])e.Data.GetData(DataFormats.FileDrop)); }; layout.Controls.Add(files,0,1);
   var options=new TableLayoutPanel { Dock=DockStyle.Fill,ColumnCount=3,RowCount=2,BackColor=Color.FromArgb(246,246,246),Padding=new Padding(8) }; options.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,105)); options.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100)); options.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,100));
   options.Controls.Add(new Label {Text="Convert To",AutoSize=true,Margin=new Padding(3,6,3,3)},0,0); format.DropDownStyle=ComboBoxStyle.DropDownList; format.Items.AddRange(new object[]{"Word (.docx)","PowerPoint (.pptx)","PDF (.pdf)"}); format.SelectedIndex=2; format.Width=220; format.SelectedIndexChanged+=delegate { RefreshSupport(); }; options.Controls.Add(format,1,0);
   options.Controls.Add(new Label {Text="Output folder",AutoSize=true,Margin=new Padding(3,6,3,3)},0,1); output.Dock=DockStyle.Fill; output.Text=settings.Output; options.Controls.Add(output,1,1); options.Controls.Add(Button("Browse...",delegate { BrowseFolder(); }),2,1); layout.Controls.Add(options,0,2);
   note.Dock=DockStyle.Fill; note.ForeColor=Color.DimGray; note.TextAlign=ContentAlignment.MiddleLeft; layout.Controls.Add(note,0,3); progress.Dock=DockStyle.Fill; layout.Controls.Add(progress,0,4);
   var actions=new FlowLayoutPanel { Dock=DockStyle.Fill }; convert=Button("Convert",async delegate { await RunQueue(); }); actions.Controls.Add(convert); actions.Controls.Add(Button("Open Folder",delegate { OpenFolder(lastFolder ?? output.Text); })); layout.Controls.Add(actions,0,5);
   var p=new TableLayoutPanel { Dock=DockStyle.Top,Height=410,ColumnCount=3,RowCount=9 }; p.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,200)); p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100)); p.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,100)); prefs.Controls.Add(p);
   var defaultOutput=new TextBox {Text=output.Text,Dock=DockStyle.Fill};defaultOutput.TextChanged+=delegate{if(output.Text!=defaultOutput.Text)output.Text=defaultOutput.Text;};output.TextChanged+=delegate{if(defaultOutput.Text!=output.Text)defaultOutput.Text=output.Text;};
   p.Controls.Add(new Label {Text="Default output folder",AutoSize=true},0,0);p.Controls.Add(defaultOutput,1,0);p.Controls.Add(Button("Browse...",delegate{BrowseFolder();}),2,0);
   beside.Text="Save beside original files"; beside.Checked=settings.Beside; beside.AutoSize=true; p.Controls.Add(beside,1,1); beside.CheckedChanged+=delegate { output.Enabled=!beside.Checked; };output.Enabled=!beside.Checked;
   rename.Text="Automatically rename duplicate files"; rename.Checked=settings.Rename; rename.AutoSize=true; p.Controls.Add(rename,1,2);
   open.Text="Open output folder after conversion"; open.Checked=settings.OpenAfter; open.AutoSize=true; p.Controls.Add(open,1,3);
   p.Controls.Add(new Label {Text="LibreOffice executable",AutoSize=true},0,4); lo.Text=settings.LibreOffice; lo.Dock=DockStyle.Fill; p.Controls.Add(lo,1,4); p.Controls.Add(Button("Browse...",delegate { BrowseExe(lo); }),2,4);
   p.Controls.Add(new Label {Text="PDF renderer (pdftoppm.exe)",AutoSize=true},0,5); renderer.Text=settings.Renderer; renderer.Dock=DockStyle.Fill; p.Controls.Add(renderer,1,5); p.Controls.Add(Button("Browse...",delegate { BrowseExe(renderer); }),2,5);
   p.Controls.Add(new Label {Text="Timeout per file (seconds)",AutoSize=true},0,6); timeout.Minimum=10; timeout.Maximum=3600; timeout.Value=Math.Max(10,Math.Min(3600,settings.Timeout)); p.Controls.Add(timeout,1,6);
   p.Controls.Add(new Label {Text="Settings are saved locally. All conversions run offline.\nOffice conversions require installed Word / PowerPoint.\nPDF rendering requires Poppler's pdftoppm.exe.",AutoSize=true},1,7);
   p.Controls.Add(Button("Save Settings",delegate { try { SaveSettings(); MessageBox.Show("Settings saved."); } catch(Exception e) { MessageBox.Show(e.Message); } }),1,8);
   FormClosing+=delegate(object s,FormClosingEventArgs e) { if(busy) { e.Cancel=true; MessageBox.Show("Wait for the conversion queue to finish before closing."); } else try { SaveSettings(); } catch {} };
   RefreshSupport();
  }
  Button Button(string text,EventHandler click) { var b=new Button {Text=text,AutoSize=true,MinimumSize=new Size(88,29),UseVisualStyleBackColor=true}; b.Click+=click; return b; }
  void BrowseFolder() { using(var d=new FolderBrowserDialog {SelectedPath=output.Text}) if(d.ShowDialog()==DialogResult.OK) output.Text=d.SelectedPath; }
  void BrowseExe(TextBox box) { using(var d=new OpenFileDialog {Filter="Executable|*.exe"}) if(d.ShowDialog()==DialogResult.OK) box.Text=d.FileName; }
  void SaveSettings() { settings.Output=output.Text; settings.Beside=beside.Checked; settings.Rename=rename.Checked; settings.OpenAfter=open.Checked; settings.LibreOffice=lo.Text; settings.Renderer=renderer.Text; settings.Timeout=(int)timeout.Value; settings.Save(); }
  string Extension {get {return new[]{".docx",".pptx",".pdf"}[format.SelectedIndex];}}
  void Add(string[] paths) { if(busy)return; var rejected=new System.Collections.Generic.List<string>(); foreach(var path in paths) { if(!File.Exists(path) || !new[]{".docx",".pptx",".pdf"}.Contains(Path.GetExtension(path).ToLowerInvariant())) { rejected.Add(path); continue; } if(files.Items.Cast<ListViewItem>().Any(i=>String.Equals((string)i.Tag,path,StringComparison.OrdinalIgnoreCase)))continue; var f=new FileInfo(path); var row=new ListViewItem(new[]{f.Name,f.Extension.ToUpperInvariant(),(f.Length/1024d).ToString("N1")+" KB","Ready",path}); row.Tag=path; files.Items.Add(row); } RefreshSupport(); if(rejected.Count>0)MessageBox.Show("Unsupported or missing files:\n"+String.Join("\n",rejected)); }
  void RefreshSupport() { if(busy)return; foreach(ListViewItem i in files.Items) { bool same=Path.GetExtension((string)i.Tag).Equals(Extension,StringComparison.OrdinalIgnoreCase); i.SubItems[3].Text=same?"Unsupported":"Ready"; i.SubItems[4].Text=same?"Same-format conversion is not offered in version 1.0.":(string)i.Tag; } note.Text=Extension==".pptx"?"Word → PowerPoint uses text slides; layout is simplified. PDF → PowerPoint uses one image per page.":Extension==".docx"?"PowerPoint → Word extracts slide text and pictures. PDF → Word uses Word's PDF import; complex layouts and scans may lose fidelity.":"PDF export uses LibreOffice when configured, otherwise locally installed Microsoft Office."; convert.Enabled=files.Items.Cast<ListViewItem>().Any(i=>i.SubItems[3].Text!="Unsupported"); }
  static void OpenFolder(string path) { try { if(!Directory.Exists(path))throw new IOException("Output folder does not exist."); Process.Start(new ProcessStartInfo("explorer.exe",Engine.Quote(path)) {UseShellExecute=true}); } catch(Exception e) { MessageBox.Show(e.Message); } }
  async Task RunQueue() {
   try { SaveSettings(); if(!settings.Beside && !Path.IsPathRooted(settings.Output))throw new IOException("Choose an absolute output folder."); } catch(Exception e) { MessageBox.Show(e.Message); return; }
   var jobs=files.Items.Cast<ListViewItem>().Where(i=>!Path.GetExtension((string)i.Tag).Equals(Extension,StringComparison.OrdinalIgnoreCase)).ToArray(); if(jobs.Length==0)return;
   busy=true; tabs.TabPages[1].Enabled=false; convert.Enabled=false; format.Enabled=false; output.Enabled=false; files.AllowDrop=false;
   // Disable list mutation controls while retaining the progress display.
   tabs.TabPages[0].Controls[0].Controls[0].Enabled=false;
   int completed=0,failed=0; progress.Maximum=jobs.Length; progress.Value=0; string ext=Extension;
   foreach(var row in jobs) {
    string input=(string)row.Tag; row.SubItems[3].Text="Converting"; status.Text="Converting: "+Path.GetFileName(input)+" | Total "+jobs.Length+" | Completed "+completed+" | Failed "+failed;
    try { string folder=settings.Beside?Path.GetDirectoryName(input):settings.Output; string result=await Task.Run(()=>Engine.Run(input,ext,folder,settings)); lastFolder=folder; row.SubItems[3].Text="Completed"; row.SubItems[4].Text=result; completed++; }
    catch(Exception e) {row.SubItems[3].Text="Failed"; row.SubItems[4].Text=e.GetBaseException().Message; failed++;} row.ToolTipText=row.SubItems[4].Text; progress.Value++;
   }
   status.Text="Completed: "+completed+" | Failed: "+failed+" | Total: "+jobs.Length; busy=false; tabs.TabPages[1].Enabled=true; format.Enabled=true; output.Enabled=!beside.Checked; files.AllowDrop=true; tabs.TabPages[0].Controls[0].Controls[0].Enabled=true; convert.Enabled=true;
   if(settings.OpenAfter && completed>0) OpenFolder(lastFolder);
  }
 }
}
