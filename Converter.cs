using System;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO.Compression;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Threading;
using System.Text;
using System.Windows.Automation;

namespace RianFile {
 public static class Engine {
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
  delegate bool WindowVisitor(IntPtr window,IntPtr data);
  [DllImport("user32.dll")] static extern bool EnumWindows(WindowVisitor visitor,IntPtr data);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr window,WindowVisitor visitor,IntPtr data);
  [DllImport("user32.dll")] static extern int GetDlgCtrlID(IntPtr window);
  [DllImport("user32.dll")] static extern IntPtr GetParent(IntPtr window);
  static string pidFile;
  public static string Quote(string s) { return "\""+s.Replace("\"", "")+"\""; }
  public static string FindRenderer(string configured) {
   if(!String.IsNullOrWhiteSpace(configured)) { if(!File.Exists(configured))throw new IOException("PDF renderer not found. Set pdftoppm.exe in Settings."); return configured; }
   var local=Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"tools","pdftoppm.exe"); if(File.Exists(local))return local;
   foreach(var dir in (Environment.GetEnvironmentVariable("PATH")??"").Split(';')) { try { var p=Path.Combine(dir.Trim('"'),"pdftoppm.exe"); if(File.Exists(p))return p; } catch {} }
   throw new IOException("PDF → PowerPoint requires Poppler pdftoppm.exe. Install Poppler and set its executable in Settings.");
  }
  public static string Run(string input,string ext,string folder,Settings settings) {
   if(!File.Exists(input))throw new IOException("Source file is missing.");
   if(!Path.IsPathRooted(folder))throw new IOException("Choose an absolute output folder.");
   Directory.CreateDirectory(folder);
   string work=Path.Combine(Path.GetTempPath(),"RianFile-"+Guid.NewGuid().ToString("N")); Directory.CreateDirectory(work);
   try {
    string copy=Path.Combine(work,"input"+Path.GetExtension(input).ToLowerInvariant()); File.Copy(input,copy); Validate(copy);
    string staged=Path.Combine(work,"output"+ext); string render=settings.Renderer;
    if(Path.GetExtension(copy)==".pdf" && ext==".pptx")render=FindRenderer(render);
    var start=new ProcessStartInfo(System.Reflection.Assembly.GetExecutingAssembly().Location,"--worker "+Quote(copy)+" "+Quote(staged)+" "+Quote(settings.LibreOffice)+" "+Quote(render)) {UseShellExecute=false,CreateNoWindow=true,WorkingDirectory=work};
    using(var proc=Process.Start(start)) {
     try { if(!proc.WaitForExit(settings.Timeout*1000))throw new TimeoutException("Conversion timed out after "+settings.Timeout+" seconds."+(File.Exists(staged+".stage")?" Stage: "+File.ReadAllText(staged+".stage"):"")); if(proc.ExitCode!=0)throw new IOException(File.Exists(staged+".error")?File.ReadAllText(staged+".error"):"Conversion engine failed."); }
     finally { if(!proc.HasExited)try { KillTree(proc); }catch{} Cleanup(staged+".pids"); }
    }
    Validate(staged);
    // File.Move never replaces a destination: this also handles concurrent filename collisions.
    string stem=Path.GetFileNameWithoutExtension(input); int index=0;
    while(true) {
     string target=Path.Combine(folder,stem+(index==0?"":" ("+index+")")+ext);
     if(String.Equals(Path.GetFullPath(target),Path.GetFullPath(input),StringComparison.OrdinalIgnoreCase))throw new IOException("Source files cannot be replaced.");
     try { File.Move(staged,target); return target; }
     catch(IOException) { if(!File.Exists(target))throw; if(!settings.Rename)throw new IOException("Output already exists. Enable automatic renaming or choose another folder."); index++; }
    }
   } finally { try {Directory.Delete(work,true);}catch{} }
  }
  [DllImport("user32.dll")] static extern IntPtr GetDlgItem(IntPtr window,int id);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr window,StringBuilder text,int count);
  [DllImport("user32.dll")] static extern bool PostMessage(IntPtr window,uint message,IntPtr wParam,IntPtr lParam);
  static readonly object pidLock=new object();
  static readonly HashSet<int> enginePids=new HashSet<int>();
  static void Record(Process process) {lock(pidLock){File.AppendAllText(pidFile,process.Id+"|"+process.StartTime.ToUniversalTime().Ticks+Environment.NewLine);enginePids.Add(process.Id);}}
  // PDFREFLOW is launched by COM, outside Word's child-process tree. Track it explicitly.
  // Its informational PDF-conversion acknowledgment ignores Word.DisplayAlerts.
  sealed class PdfImportMonitor : IDisposable {
   volatile bool stop; Thread thread; int[] previous; HashSet<int> tracked=new HashSet<int>();
   public PdfImportMonitor(bool enabled) {
    if(!enabled)return;
    previous=Process.GetProcessesByName("PDFREFLOW").Select(p=>p.Id).ToArray();
    if(previous.Length>0)throw new IOException("Close existing Word PDF import dialogs before converting PDF to Word.");
    thread=new Thread(Watch) {IsBackground=true};thread.Start();
   }
   void Watch() {while(!stop) {
    foreach(var p in Process.GetProcessesByName("PDFREFLOW"))using(p)try {if(!previous.Contains(p.Id) && tracked.Add(p.Id))Record(p);}catch{}
    EnumWindows(delegate(IntPtr window,IntPtr data) {
     uint pid;GetWindowThreadProcessId(window,out pid);bool owned;lock(pidLock)owned=enginePids.Contains((int)pid);if(!owned)return true;
     try {
      var root=AutomationElement.FromHandle(window);
      var label=root.FindFirst(TreeScope.Descendants,new AndCondition(new PropertyCondition(AutomationElement.AutomationIdProperty,"4001"),new PropertyCondition(AutomationElement.ControlTypeProperty,ControlType.Text)));
      if(label!=null && label.Current.Name.StartsWith("Word will now convert your PDF",StringComparison.Ordinal)) {
       var parent=TreeWalker.ControlViewWalker.GetParent(label);
       var button=parent.FindFirst(TreeScope.Descendants,new AndCondition(new PropertyCondition(AutomationElement.AutomationIdProperty,"1"),new PropertyCondition(AutomationElement.ControlTypeProperty,ControlType.Button)));
       object pattern;if(button!=null && button.TryGetCurrentPattern(InvokePattern.Pattern,out pattern))((InvokePattern)pattern).Invoke();
      }
     }catch{}return true;
    },IntPtr.Zero);
    Thread.Sleep(100);
   }}
   public void Dispose(){stop=true;if(thread!=null)thread.Join(2000);}
  }
  public static void Validate(string path) {
   if(!File.Exists(path) || new FileInfo(path).Length==0)throw new IOException("Conversion produced no output.");
   string ext=Path.GetExtension(path).ToLowerInvariant();
   if(ext==".pdf") { using(var s=File.OpenRead(path)) { var b=new byte[5]; if(s.Read(b,0,5)!=5 || System.Text.Encoding.ASCII.GetString(b)!="%PDF-")throw new IOException("Invalid or corrupted PDF."); } return; }
   if(ext!=".docx" && ext!=".pptx")throw new IOException("Unsupported file format.");
   if(new FileInfo(path).Length<22)throw new IOException("This Office document is damaged or is not a valid "+ext.Substring(1).ToUpperInvariant()+" file.");
   try { using(var z=ZipFile.OpenRead(path)) { string entry=ext==".docx"?"word/document.xml":"ppt/presentation.xml"; if(z.GetEntry("[Content_Types].xml")==null || z.GetEntry(entry)==null)throw new IOException("Invalid or corrupted Office document."); using(var s=z.GetEntry(entry).Open())XDocument.Load(s); } }
   catch(InvalidDataException e){throw new IOException("This Office document is damaged or is not a valid "+ext.Substring(1).ToUpperInvariant()+" file.",e);}
   catch(System.Xml.XmlException e){throw new IOException("This Office document contains damaged document data.",e);}
  }
  static void Cleanup(string file) {
   if(!File.Exists(file))return;
   foreach(string line in File.ReadAllLines(file))try { var parts=line.Split('|'); using(var p=Process.GetProcessById(Int32.Parse(parts[0]))) if(p.StartTime.ToUniversalTime().Ticks==Int64.Parse(parts[1]) && new[]{"WINWORD","POWERPNT","PDFREFLOW","soffice","soffice.bin","pdftoppm"}.Contains(p.ProcessName,StringComparer.OrdinalIgnoreCase)) KillTree(p); }catch{}
  }
  static void KillTree(Process process) {
   using(var killer=Process.Start(new ProcessStartInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),"taskkill.exe"),"/PID "+process.Id+" /T /F") {UseShellExecute=false,CreateNoWindow=true,RedirectStandardOutput=true,RedirectStandardError=true})) {killer.WaitForExit(5000);}
   if(!process.HasExited)process.Kill(); process.WaitForExit(3000);
  }
  static dynamic Office(string progid) {
   // PowerPoint can reuse an interactive instance. Never attach to a user's open presentation.
   if(progid=="PowerPoint.Application" && Process.GetProcessesByName("POWERPNT").Length>0)throw new IOException("Close PowerPoint before converting, so RianFile can use a separate engine safely.");
   var t=Type.GetTypeFromProgID(progid); if(t==null)throw new IOException("Required local engine is missing: "+progid+". Install Microsoft Office or configure LibreOffice for PDF export.");
   string processName=progid=="Word.Application"?"WINWORD":"POWERPNT";
   var previous=Process.GetProcessesByName(processName).Select(p=>p.Id).ToArray();
   dynamic app=Activator.CreateInstance(t);
   var created=Process.GetProcessesByName(processName).Where(p=>!previous.Contains(p.Id)).ToArray();
   if(created.Length!=1) { Release((object)app); throw new IOException("Cannot establish a dedicated Office process. Close Office and retry."); }
   using(var p=created[0])Record(p);
   return app;
  }
  static void Release(object o) { if(o!=null && Marshal.IsComObject(o))try {Marshal.FinalReleaseComObject(o);}catch{} }
  public static void Convert(string input,string output,string libreOffice,string renderer) {
   pidFile=output+".pids";
   string src=Path.GetExtension(input).ToLowerInvariant(), dst=Path.GetExtension(output).ToLowerInvariant();
   if(src==dst)throw new IOException("Same-format processing is not supported.");
   if(dst==".pdf" && !String.IsNullOrWhiteSpace(libreOffice)) {
    string exe=Directory.Exists(libreOffice)?Path.Combine(libreOffice,"program","soffice.exe"):libreOffice;
    if(!File.Exists(exe))throw new IOException("LibreOffice executable was not found. Update Settings.");
    var profile=new Uri(Path.Combine(Path.GetDirectoryName(output),"lo-profile")).AbsoluteUri;
    RunTool(exe,"-env:UserInstallation="+Quote(profile)+" --headless --convert-to pdf --outdir "+Quote(Path.GetDirectoryName(output))+" "+Quote(input));
    string generated=Path.ChangeExtension(input,".pdf"); if(!File.Exists(generated))throw new IOException("LibreOffice did not produce a PDF. The document may be corrupted."); File.Move(generated,output); return;
   }
   if(src==".pdf" && dst==".pptx") { PdfSlides(input,output,renderer); return; }
   if(src==".pptx") { FromPresentation(input,output,dst); return; }
   dynamic word=null,doc=null;
   using(var monitor=new PdfImportMonitor(src==".pdf")) try {
    File.WriteAllText(output+".stage","Starting Word");
    word=Office("Word.Application"); word.Visible=false; word.DisplayAlerts=0; word.AutomationSecurity=3;
    File.WriteAllText(output+".stage","Opening source in Word");
    doc=word.Documents.Open(FileName:input,ConfirmConversions:false,ReadOnly:true,AddToRecentFiles:false,Visible:false,OpenAndRepair:false,NoEncodingDialog:true);
    File.WriteAllText(output+".stage","Writing converted document");
    if(dst==".pdf") doc.ExportAsFixedFormat(output,17);
    else if(dst==".docx") doc.SaveAs2(FileName:output,FileFormat:16,AddToRecentFiles:false);
    else {
     string text=(string)doc.Content.Text; if(String.IsNullOrWhiteSpace(text))throw new IOException("No text was found in this Word document. Image-only documents cannot become text slides.");
     TextSlides(text,output);
    }
   } finally { File.WriteAllText(output+".stage","Closing Word"); if(doc!=null)try{doc.Close(0);}catch{} if(word!=null)try{word.Quit(0);}catch{} Release((object)doc);Release((object)word); }
  }
  static void RunTool(string exe,string args) {
   using(var p=Process.Start(new ProcessStartInfo(exe,args){UseShellExecute=false,CreateNoWindow=true})) {
    File.AppendAllText(pidFile,p.Id+"|"+p.StartTime.ToUniversalTime().Ticks+Environment.NewLine);
    p.WaitForExit(); if(p.ExitCode!=0)throw new IOException(Path.GetFileName(exe)+" failed (exit "+p.ExitCode+"). The input may be damaged or password protected.");
   }
  }
  static void TextSlides(string text,string output) {
   dynamic ppt=null,pres=null;
   try {
    ppt=Office("PowerPoint.Application"); ppt.AutomationSecurity=3; pres=ppt.Presentations.Add(0); pres.PageSetup.SlideWidth=960; pres.PageSetup.SlideHeight=540;
    var blocks=new List<string>(); string current="";
    foreach(string line in text.Replace("\a","").Split(new[]{'\r','\n'},StringSplitOptions.RemoveEmptyEntries)) {
     for(int offset=0;offset<line.Length;offset+=700) { string piece=line.Substring(offset,Math.Min(700,line.Length-offset)); if(current.Length+piece.Length>750) {blocks.Add(current);current="";} current+=piece+"\r\n"; }
    }
    if(current.Length>0)blocks.Add(current);
    for(int n=0;n<blocks.Count;n++) { dynamic slide=pres.Slides.Add(n+1,12); dynamic title=slide.Shapes.AddTextbox(1,40,25,880,50); title.TextFrame.TextRange.Text="Document — "+(n+1); title.TextFrame.TextRange.Font.Size=28; dynamic body=slide.Shapes.AddTextbox(1,40,95,880,415); body.TextFrame.TextRange.Text=blocks[n]; body.TextFrame.TextRange.Font.Size=20; body.TextFrame.WordWrap=-1; body.TextFrame2.AutoSize=2; }
    pres.SaveAs(output,24);
   } finally { if(pres!=null)try{pres.Close();}catch{} if(ppt!=null)try{ppt.Quit();}catch{} Release((object)pres);Release((object)ppt); }
  }
  static void FromPresentation(string input,string output,string dst) {
   dynamic ppt=null,pres=null,word=null,doc=null;
   try {
    ppt=Office("PowerPoint.Application"); ppt.AutomationSecurity=3; pres=ppt.Presentations.Open(input,-1,0,0);
    if(dst==".pdf") { pres.SaveAs(output,32); return; }
    word=Office("Word.Application"); word.Visible=false; word.DisplayAlerts=0; word.AutomationSecurity=3; doc=word.Documents.Add();
    for(int n=1;n<=pres.Slides.Count;n++) {
     dynamic slide=pres.Slides[n]; Append(doc,"Slide "+n+"\r",true);
     for(int j=1;j<=slide.Shapes.Count;j++)ExtractShape(slide.Shapes[j],doc,Path.GetDirectoryName(output),n+"-"+j);
     Append(doc,"\r",false);
    }
    doc.SaveAs2(FileName:output,FileFormat:16,AddToRecentFiles:false);
   } finally { if(doc!=null)try{doc.Close(0);}catch{} if(word!=null)try{word.Quit(0);}catch{} if(pres!=null)try{pres.Close();}catch{} if(ppt!=null)try{ppt.Quit();}catch{} Release((object)doc);Release((object)word);Release((object)pres);Release((object)ppt); }
  }
  static void Append(dynamic doc,string text,bool heading) { dynamic range=doc.Range(doc.Content.End-1,doc.Content.End-1); range.Text=text; range.Font.Bold=heading?-1:0; range.Font.Size=heading?18:11; }
  static void ExtractShape(dynamic shape,dynamic doc,string folder,string id) {
   if(shape.Type==6) {for(int n=1;n<=shape.GroupItems.Count;n++)ExtractShape(shape.GroupItems[n],doc,folder,id+"-"+n);return;}
   if(shape.HasTextFrame==-1 && shape.TextFrame.HasText==-1)Append(doc,(string)shape.TextFrame.TextRange.Text+"\r",false);
   if(shape.HasTable==-1) {for(int r=1;r<=shape.Table.Rows.Count;r++) {string line="";for(int c=1;c<=shape.Table.Columns.Count;c++)line+=(string)shape.Table.Cell(r,c).Shape.TextFrame.TextRange.Text+"\t";Append(doc,line+"\r",false);}}
   if(shape.Type==13 || shape.Type==11) {
    string path=Path.Combine(folder,"picture-"+id+".png"); shape.Export(path,2);
    dynamic range=doc.Range(doc.Content.End-1,doc.Content.End-1); dynamic pic=doc.InlineShapes.AddPicture(path,false,true,range); pic.LockAspectRatio=-1; if(pic.Width>450)pic.Width=450; Append(doc,"\r",false);
   }
  }
  static void PdfSlides(string input,string output,string renderer) {
   string prefix=Path.Combine(Path.GetDirectoryName(output),"page"); RunTool(FindRenderer(renderer),"-png -r 120 "+Quote(input)+" "+Quote(prefix));
   var pages=Directory.GetFiles(Path.GetDirectoryName(output),"page-*.png").OrderBy(p=>Int32.Parse(Path.GetFileNameWithoutExtension(p).Substring(5))).ToArray(); if(pages.Length==0)throw new IOException("The PDF renderer produced no pages.");
   dynamic ppt=null,pres=null;
   try { ppt=Office("PowerPoint.Application"); ppt.AutomationSecurity=3; pres=ppt.Presentations.Add(0);
    using(var image=System.Drawing.Image.FromFile(pages[0])) {pres.PageSetup.SlideWidth=720;pres.PageSetup.SlideHeight=720f*image.Height/image.Width;}
    float width=(float)pres.PageSetup.SlideWidth,height=(float)pres.PageSetup.SlideHeight;
    for(int n=0;n<pages.Length;n++)using(var img=System.Drawing.Image.FromFile(pages[n])) { dynamic slide=pres.Slides.Add(n+1,12); float scale=Math.Min(width/img.Width,height/img.Height);float w=img.Width*scale,h=img.Height*scale;slide.Shapes.AddPicture(pages[n],0,-1,(width-w)/2,(height-h)/2,w,h); }
    pres.SaveAs(output,24);
   } finally { if(pres!=null)try{pres.Close();}catch{} if(ppt!=null)try{ppt.Quit();}catch{} Release((object)pres);Release((object)ppt); }
  }
 }
}
