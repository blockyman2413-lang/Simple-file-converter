using System;
using System.IO;
using System.Diagnostics;
using RianFile;
class FailureTests {
 [STAThread] static int Main() {
  try {
   var settings=new Settings {Timeout=2,LibreOffice=Path.GetFullPath("bin/SlowEngine.exe")};
   bool timedOut=false;
   try {Engine.Run(Path.GetFullPath("test-results/sample.docx"),".pdf",Path.GetFullPath("test-results/timeout"),settings);}catch(TimeoutException){timedOut=true;}
   if(!timedOut || Directory.GetFiles("test-results/timeout").Length!=0 || Process.GetProcessesByName("SlowEngine").Length>0)throw new Exception("Timeout did not clean up the job and engine.");
   Console.WriteLine("PASS timeout: no output published; worker and engine cleaned up");return 0;
  }catch(Exception e){Console.WriteLine("FAIL "+e);return 1;}
 }
}
