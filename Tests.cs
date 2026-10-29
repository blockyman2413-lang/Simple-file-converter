using System;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.IO.Compression;
using RianFile;
class Tests {
 static System.Collections.Generic.List<Process> owned=new System.Collections.Generic.List<Process>();
 static dynamic Office(string name) {string process=name.StartsWith("Word")?"WINWORD":"POWERPNT";var before=Process.GetProcessesByName(process).Select(p=>p.Id).ToArray();dynamic app=Activator.CreateInstance(Type.GetTypeFromProgID(name));owned.AddRange(Process.GetProcessesByName(process).Where(p=>!before.Contains(p.Id)));return app;}
 static void Cleanup() {foreach(var p in owned)try{if(!p.WaitForExit(2000)){p.Kill();p.WaitForExit(3000);}}catch{}owned.Clear();}
 static void Assert(bool ok,string message) {if(!ok)throw new Exception(message);}
 static string Text(string path) {using(var z=ZipFile.OpenRead(path))return String.Join("\n",z.Entries.Where(e=>e.FullName.EndsWith(".xml")).Select(e=>{using(var r=new StreamReader(e.Open()))return r.ReadToEnd();}));}
 static void Fails(Action a,string name) {try{a();}catch{Console.WriteLine("PASS "+name);return;}throw new Exception("Expected failure: "+name);}
 [STAThread] static int Main(string[] args) {
  string root=Path.GetFullPath("test-results");Directory.CreateDirectory(root);
  byte[] savedSettings=File.Exists(Settings.Location)?File.ReadAllBytes(Settings.Location):null;
  try {
   if(args.Length>0 && args[0]=="--ui-only"){UiTests.Run(root);return 0;}
   dynamic word=Office("Word.Application"); word.Visible=false;word.DisplayAlerts=0;
   dynamic doc=null;
   try {
doc=word.Documents.Add();doc.Content.Text="RianFile verification alpha\rSecond paragraph with expected content.\r";doc.SaveAs2(Path.Combine(root,"sample.docx"),16);doc.ExportAsFixedFormat(Path.Combine(root,"sample.pdf"),17);}finally{if(doc!=null)doc.Close(0);word.Quit(0);}
   GC.Collect();GC.WaitForPendingFinalizers();
   dynamic ppt=Office("PowerPoint.Application");dynamic pres=null;
   try {
pres=ppt.Presentations.Add(0);dynamic slide=pres.Slides.Add(1,12);dynamic shape=slide.Shapes.AddTextbox(1,30,30,600,150);shape.TextFrame.TextRange.Text="RianFile verification alpha";
    using(var bmp=new System.Drawing.Bitmap(60,40)){using(var g=System.Drawing.Graphics.FromImage(bmp))g.Clear(System.Drawing.Color.CornflowerBlue);bmp.Save(Path.Combine(root,"fixture.png"));}
    slide.Shapes.AddPicture(Path.Combine(root,"fixture.png"),0,-1,40,200,120,80);pres.SaveAs(Path.Combine(root,"sample.pptx"),24);
   }finally{if(pres!=null)pres.Close();ppt.Quit();}
   GC.Collect();GC.WaitForPendingFinalizers();Cleanup();
   var hashes=Directory.GetFiles(root,"sample.*").ToDictionary(p=>p,p=>System.Convert.ToBase64String(System.Security.Cryptography.SHA256.Create().ComputeHash(File.ReadAllBytes(p))));
   var settings=new Settings {Timeout=60};
   foreach(string src in new[]{".docx",".pptx",".pdf"})foreach(string dst in new[]{".docx",".pptx",".pdf"})if(src!=dst) {
    string folder=Path.Combine(root,src.Substring(1)+"-to-"+dst.Substring(1));string path=Engine.Run(Path.Combine(root,"sample"+src),dst,folder,settings);Engine.Validate(path);
    if(dst==".docx") {Assert(Text(path).Contains("verification alpha"),"Missing Word text");if(src==".pptx")using(var zip=ZipFile.OpenRead(path))Assert(zip.Entries.Any(e=>e.FullName.StartsWith("word/media/")),"Missing slide image");dynamic app=Office("Word.Application");app.DisplayAlerts=0;dynamic d=null;try{d=app.Documents.Open(path,ReadOnly:true,Visible:false);Assert(d.Content.Text.Contains("verification alpha"),"Word could not read content");}finally{if(d!=null)d.Close(0);app.Quit(0);}}
    if(dst==".pptx") {if(src==".docx")Assert(Text(path).Contains("verification alpha"),"Missing slide text");else using(var zip=ZipFile.OpenRead(path))Assert(zip.Entries.Count(e=>e.FullName.StartsWith("ppt/media/"))==1,"PDF page not embedded");dynamic app=Office("PowerPoint.Application");dynamic d=null;try{d=app.Presentations.Open(path,-1,0,0);Assert(d.Slides.Count>0,"No slides");}finally{if(d!=null)d.Close();app.Quit();}System.Threading.Thread.Sleep(1000);}
    if(dst==".pdf") {string renderer=Engine.FindRenderer("");string output=Path.Combine(folder,"render");using(var p=Process.Start(new ProcessStartInfo(renderer,"-f 1 -singlefile -png -scale-to 800 "+Engine.Quote(path)+" "+Engine.Quote(output)){UseShellExecute=false,CreateNoWindow=true})){p.WaitForExit();Assert(p.ExitCode==0 && File.Exists(output+".png"),"PDF could not be rendered");}string textTool=Path.Combine(Path.GetDirectoryName(renderer),"pdftotext.exe");if(File.Exists(textTool)){using(var p=Process.Start(new ProcessStartInfo(textTool,Engine.Quote(path)+" "+Engine.Quote(output+".txt")){UseShellExecute=false,CreateNoWindow=true})){p.WaitForExit();Assert(p.ExitCode==0 && File.ReadAllText(output+".txt").Contains("verification alpha"),"PDF content missing");}}}
    Cleanup(); Console.WriteLine("PASS "+src+" → "+dst+": "+path);
   }
   foreach(var item in hashes)Assert(item.Value==System.Convert.ToBase64String(System.Security.Cryptography.SHA256.Create().ComputeHash(File.ReadAllBytes(item.Key))),"Source changed");Console.WriteLine("PASS source hashes unchanged");
   string duplicates=Path.Combine(root,"duplicates");string first=Engine.Run(Path.Combine(root,"sample.docx"),".pdf",duplicates,settings);string second=Engine.Run(Path.Combine(root,"sample.docx"),".pdf",duplicates,settings);Assert(first!=second,"Duplicate overwritten");Console.WriteLine("PASS duplicate renaming");settings.Rename=false;Fails(()=>Engine.Run(Path.Combine(root,"sample.docx"),".pdf",duplicates,settings),"duplicate rejection");settings.Rename=true;
   File.WriteAllText(Path.Combine(root,"corrupt.docx"),"broken");Fails(()=>Engine.Run(Path.Combine(root,"corrupt.docx"),".pdf",root,settings),"corrupted document");File.WriteAllText(Path.Combine(root,"unsupported.txt"),"text");Fails(()=>Engine.Run(Path.Combine(root,"unsupported.txt"),".pdf",root,settings),"unsupported format");Fails(()=>Engine.Run(Path.Combine(root,"missing.docx"),".pdf",root,settings),"missing input");settings.LibreOffice=Path.Combine(root,"missing-engine.exe");Fails(()=>Engine.Run(Path.Combine(root,"sample.docx"),".pdf",root,settings),"missing conversion engine");settings.LibreOffice="";settings.Renderer=Path.Combine(root,"missing-renderer.exe");Fails(()=>Engine.Run(Path.Combine(root,"sample.pdf"),".pptx",root,settings),"missing PDF renderer");
   settings.Renderer="";
   File.WriteAllText(Path.Combine(root,"not-a-folder"),"file");Fails(()=>Engine.Run(Path.Combine(root,"sample.docx"),".pdf",Path.Combine(root,"not-a-folder"),settings),"invalid output directory");
   UiTests.Run(root);
   Console.WriteLine("ALL CHECKS PASSED");return 0;
  }catch(Exception e){Console.WriteLine("FAIL "+e);return 1;}
  finally {Cleanup();if(savedSettings!=null)File.WriteAllBytes(Settings.Location,savedSettings);else if(File.Exists(Settings.Location))File.Delete(Settings.Location);}
 }
}
