using System;
using System.IO;
using System.Drawing;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using RianFile;
static class UiTests {
 public static void Run(string root) {
  Exception failure=null;int ticks=0;
  Control.CheckForIllegalCrossThreadCalls=true;
  using(var form=new MainForm())using(var timer=new Timer {Interval=50}) {
   timer.Tick+=delegate{ticks++;};
   var flags=BindingFlags.NonPublic|BindingFlags.Instance;
   ((TextBox)typeof(MainForm).GetField("output",flags).GetValue(form)).Text=Path.Combine(root,"batch");
   ((CheckBox)typeof(MainForm).GetField("beside",flags).GetValue(form)).Checked=false;
   ((CheckBox)typeof(MainForm).GetField("open",flags).GetValue(form)).Checked=false;
   ((CheckBox)typeof(MainForm).GetField("rename",flags).GetValue(form)).Checked=true;
   ((TextBox)typeof(MainForm).GetField("lo",flags).GetValue(form)).Text="";
   form.Shown+=async delegate {
    try {
     typeof(MainForm).GetMethod("Add",flags).Invoke(form,new object[]{new[]{Path.Combine(root,"sample.docx"),Path.Combine(root,"corrupt.docx"),Path.Combine(root,"sample.pptx")}});
     timer.Start();await (Task)typeof(MainForm).GetMethod("RunQueue",flags).Invoke(form,null);timer.Stop();
     var rows=(ListView)typeof(MainForm).GetField("files",flags).GetValue(form);
     if(rows.Items[0].SubItems[3].Text!="Completed" || rows.Items[1].SubItems[3].Text!="Failed" || rows.Items[2].SubItems[3].Text!="Completed")throw new Exception("Queue did not continue after failure.");
     var progress=(ProgressBar)typeof(MainForm).GetField("progress",flags).GetValue(form);
     if(progress.Value!=3 || ticks<10)throw new Exception("Progress or UI responsiveness failed.");
     using(var bmp=new Bitmap(form.Width,form.Height)){form.DrawToBitmap(bmp,new Rectangle(0,0,form.Width,form.Height));bmp.Save(Path.Combine(root,"interface.png"));}
     var tabs=(TabControl)typeof(MainForm).GetField("tabs",flags).GetValue(form);tabs.SelectedIndex=1;form.Refresh();
     using(var bmp=new Bitmap(form.Width,form.Height)){form.DrawToBitmap(bmp,new Rectangle(0,0,form.Width,form.Height));bmp.Save(Path.Combine(root,"settings.png"));}
     Console.WriteLine("PASS UI batch continues after failure; progress 3/3; UI timer ticks: "+ticks);
    }catch(Exception e){failure=e;}finally{form.Close();}
   };
   Application.Run(form);
  }
  if(failure!=null)throw failure;
 }
}
