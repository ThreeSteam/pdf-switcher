using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Win32;

public class Observable : INotifyPropertyChanged {
 public event PropertyChangedEventHandler PropertyChanged;
 public void Changed(params string[] keys) {foreach(var k in keys) {var h=PropertyChanged;if(h!=null)h(this,new PropertyChangedEventArgs(k));}}
}
public class FileRow : Observable {
 public string Source {get;set;}
 public string Name {get{return Path.GetFileName(Source);}}
 public string Folder {get{return Path.GetDirectoryName(Source);}}
 public string Output {get{return Path.ChangeExtension(Source,".pdf");}}
 public string OutputName {get{return Path.GetFileName(Output);}}
 public string Status {get;private set;}
 public string Detail {get;private set;}
 public string Kind="pending";
 public string Badge {get{return Kind=="success"?"#EAF6EF":Kind=="error"?"#FFF0EF":Kind=="skipped"?"#FFF6E6":"#EEF3FB";}}
 public string Ink {get{return Kind=="success"?"#248452":Kind=="error"?"#CA534A":Kind=="skipped"?"#A77828":"#53749D";}}
 public void Set(string label,string kind,string detail=null){Status=label;Kind=kind;Detail=detail??label;Changed("Status","Detail","Badge","Ink");}
}
public class Converter : Observable {
 public ObservableCollection<FileRow> Files {get;private set;}
 public Window Window;
 DataGrid grid; ComboBox review,conflict; Border surface;
 readonly DispatcherTimer timer=new DispatcherTimer();
 readonly JavaScriptSerializer json=new JavaScriptSerializer();
 readonly Queue<FileRow> queue=new Queue<FileRow>();
 readonly string session=Path.Combine(Path.GetTempPath(),"PDF-Switcher",Guid.NewGuid().ToString("N"));
 Process worker; FileRow active; string job,mode="auto",message="准备好文档后，就可以开始了。";
 DateTime started; bool busy,cancelling,overwrite; int total,done,failed,skipped;
 string testDir;
 public bool Busy {get{return busy;}}
 public bool CanEdit {get{return !busy;}}
 public bool CanStart {get{return !busy&&Files.Any(f=>f.Kind!="success");}}
 public bool CanRetry {get{return !busy&&Files.Any(f=>f.Kind=="error");}}
 public double Progress {get{return total==0?0:100.0*done/total;}}
 public string Summary {get{return Files.Count+" 份文档  ·  "+Files.Count(f=>f.Kind=="success")+" 份已完成";}}
 public string Message {get{return message;}set{message=value;Changed("Message");}}
 void Notify(){Changed("Busy","CanEdit","CanStart","CanRetry","Progress","Summary");}
 T Find<T>(string n) where T:FrameworkElement {return (T)Window.FindName(n);}
 public Converter(){Files=new ObservableCollection<FileRow>();timer.Interval=TimeSpan.FromMilliseconds(400);timer.Tick+=(s,e)=>Poll();}
 public void Wire(Window w) {
  Window=w;w.DataContext=this;grid=Find<DataGrid>("FilesGrid");surface=Find<Border>("DropSurface");review=Find<ComboBox>("Review");conflict=Find<ComboBox>("Conflict");
  w.PreviewDragOver+=(s,e)=>{e.Effects=CanEdit&&e.Data.GetDataPresent(DataFormats.FileDrop)?DragDropEffects.Copy:DragDropEffects.None;e.Handled=true;surface.BorderBrush=new SolidColorBrush(Color.FromRgb(37,99,235));};
  w.DragLeave+=(s,e)=>ResetBorder();
  w.PreviewDrop+=(s,e)=>{ResetBorder();if(CanEdit&&e.Data.GetDataPresent(DataFormats.FileDrop))Add((string[])e.Data.GetData(DataFormats.FileDrop));e.Handled=true;};
  Find<Button>("Add").Click+=(s,e)=>{var d=new OpenFileDialog{Multiselect=true,Filter="Word 文档|*.docx;*.doc;*.docm",Title="添加要转换的文件"};if(d.ShowDialog(w)==true)Add(d.FileNames);};
  Find<Button>("Remove").Click+=(s,e)=>Remove();
  grid.PreviewKeyDown+=(s,e)=>{if(e.Key==Key.Delete&&CanEdit){Remove();e.Handled=true;}};
  Find<Button>("Clear").Click+=(s,e)=>{if(CanEdit){Files.Clear();done=total=0;Message="已清空，可以添加下一组文件。";Notify();}};
  Find<Button>("Start").Click+=(s,e)=>Begin(false);
  Find<Button>("Retry").Click+=(s,e)=>Begin(true);
  Find<Button>("Stop").Click+=(s,e)=>{cancelling=true;foreach(var f in queue)f.Set("已取消","pending");queue.Clear();Message="将在当前文件转换完成后停止。";};
  Find<Button>("OpenFolder").Click+=(s,e)=>OpenFolder();
  grid.MouseDoubleClick+=(s,e)=>{var f=grid.SelectedItem as FileRow;if(f!=null){if(f.Kind=="error")MessageBox.Show(w,f.Detail,"转换失败详情");else OpenFolder();}};
  w.Closing+=(s,e)=>{if(busy){if(MessageBox.Show(w,"关闭后将停止当前转换及剩余任务。已经生成的 PDF 会保留。","关闭 PDF Switcher",MessageBoxButton.OKCancel)==MessageBoxResult.Cancel){e.Cancel=true;return;}StopOwned();}};
 }
 void ResetBorder(){surface.BorderBrush=new SolidColorBrush(Color.FromRgb(220,229,242));}
 void Remove(){if(!CanEdit)return;foreach(var f in grid.SelectedItems.Cast<FileRow>().ToArray())Files.Remove(f);Notify();}
 static bool Supported(string p){return new[]{".doc",".docx",".docm"}.Contains(Path.GetExtension(p).ToLowerInvariant())&&!Path.GetFileName(p).StartsWith("~$");}
 public void Add(IEnumerable<string> paths) {
  if(busy)return;int ignored=0,duplicates=0,added=0;
  foreach(string input in paths) {
   try {
    var items=Directory.Exists(input)?Directory.GetFiles(input):new[]{input};
    foreach(var path in items){
     if(!File.Exists(path)||!Supported(path)){ignored++;continue;}
     string full=Path.GetFullPath(path);
     if(Files.Any(existing=>string.Equals(existing.Source,full,StringComparison.OrdinalIgnoreCase))){duplicates++;continue;}
     var f=new FileRow{Source=full};f.Set(File.Exists(f.Output)?"已有同名 PDF":"待转换","pending");Files.Add(f);added++;
    }
   }catch(Exception ex){Message="添加失败："+ex.Message;}
  }
  Message="已添加 "+added+" 份文档"+(duplicates>0?"，去重 "+duplicates+" 份":"")+(ignored>0?"，忽略 "+ignored+" 个不支持的文件":"")+"。点击开始转换。";Notify();
 }
 public void Begin(bool retry) {
  if(busy)return;
  var selected=Files.Where(f=>retry?f.Kind=="error":f.Kind!="success").ToArray();
  if(selected.Length==0){Message="请先添加 Word 文档。";return;}
  // Two sources with one PDF name cannot both preserve their name. Require a choice.
  var collision=selected.GroupBy(f=>f.Output,StringComparer.OrdinalIgnoreCase).FirstOrDefault(g=>g.Count()>1);
  if(collision!=null){Message="输出名称重复，请移除其中一份："+Path.GetFileName(collision.Key);return;}
  overwrite=conflict.SelectedIndex==1;mode=new[]{"auto","clean","markup"}[Math.Max(0,review.SelectedIndex)];
  Directory.CreateDirectory(session);
  foreach(var name in new[]{"ConvertWorker.ps1","WordNative.ps1"})using(var stream=Assembly.GetExecutingAssembly().GetManifestResourceStream(name))using(var file=File.Create(Path.Combine(session,name)))stream.CopyTo(file);
  queue.Clear();foreach(var f in selected){f.Set("等待转换","pending");queue.Enqueue(f);}
  done=failed=skipped=0;total=selected.Length;busy=true;cancelling=false;Notify();Next();if(busy)timer.Start();
 }
 void Next(){
  if(queue.Count==0){Finish();return;}
  active=queue.Dequeue();job=Path.Combine(session,Guid.NewGuid().ToString("N")+".json");
  try{
   File.WriteAllText(job,json.Serialize(new{Source=active.Source,Overwrite=overwrite,Review=mode}),new UTF8Encoding(true));
   var exe=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),@"WindowsPowerShell\v1.0\powershell.exe");
   worker=Process.Start(new ProcessStartInfo(exe,"-NoProfile -STA -ExecutionPolicy Bypass -File \""+Path.Combine(session,"ConvertWorker.ps1")+"\" -JobFile \""+job+"\""){UseShellExecute=false,CreateNoWindow=true,WorkingDirectory=session});
   started=DateTime.Now;active.Set("准备转换","active");
  }catch(Exception ex){Complete(ex.Message,false,"");}
 }
 Dictionary<string,object> Read(string path){return json.Deserialize<Dictionary<string,object>>(File.ReadAllText(path,Encoding.UTF8));}
 void Poll(){
  if(!busy||worker==null)return;
  try {
   if(!worker.HasExited){
    string phase="准备转换";DateTime at=started;
    try{var p=Read(job+".phase");phase=Convert.ToString(p["Phase"]);at=DateTimeOffset.Parse(Convert.ToString(p["At"])).LocalDateTime;}catch{}
    active.Set(phase,"active");Message=(cancelling?"取消剩余中 · ":"")+"正在转换 "+(done+1)+" / "+total+" · "+active.Name+" · "+(int)(DateTime.Now-started).TotalSeconds+" 秒";
    if((DateTime.Now-at).TotalSeconds>180){StopOwned();Complete("转换超时（"+phase+"），可关闭文档中的提示窗口后重试。",false,"");}return;
   }
   var r=Read(job+".result");bool skip=Convert.ToBoolean(r["Skipped"]);Complete(Convert.ToBoolean(r["Success"])||skip?null:Convert.ToString(r["Error"]),skip,Convert.ToString(r["Mode"]));
  }catch(Exception ex){StopOwned();Complete("转换进程异常："+ex.Message,false,"");}
 }
 void Complete(string error,bool skip,string resultMode){
  if(worker!=null){worker.Dispose();worker=null;}
  active.Set(error!=null?"转换失败":skip?"已跳过":"已完成 · "+resultMode,error!=null?"error":skip?"skipped":"success",error??(skip?"同名 PDF 已存在，未覆盖。":active.Output));
  if(error!=null)failed++;if(skip)skipped++;done++;Notify();Next();
 }
 void StopOwned(){
  try{if(worker!=null&&!worker.HasExited){worker.Kill();worker.WaitForExit(2000);}}catch{}
  try{var owner=Read(job+".owner");using(var p=Process.GetProcessById(Convert.ToInt32(owner["Id"])))if(p.ProcessName.Equals("WINWORD",StringComparison.OrdinalIgnoreCase)&&p.StartTime.Ticks==Convert.ToInt64(owner["Ticks"]))p.Kill();}catch{}
 }
 void Finish(){
  timer.Stop();busy=false;Message=(cancelling?"已停止":"转换完成")+"：成功 "+(done-failed-skipped)+"，跳过 "+skipped+"，失败 "+failed+"。";Notify();
  var first=Files.FirstOrDefault(f=>f.Kind=="success");if(first!=null)grid.SelectedItem=first;
  if(testDir!=null){File.WriteAllText(Path.Combine(testDir,"result.json"),json.Serialize(new{Total=total,Done=done,Failed=failed,Skipped=skipped,Statuses=Files.Select(f=>f.Status).ToArray(),Details=Files.Select(f=>f.Detail).ToArray()}));Program.ExitCode=failed>0?1:0;Window.Close();}
 }
 void OpenFolder(){var f=grid.SelectedItem as FileRow;if(f==null)f=Files.FirstOrDefault();if(f==null){Message="请先选择一份文件。";return;}try{Process.Start("explorer.exe","/select,\""+(File.Exists(f.Output)?f.Output:f.Source)+"\"");}catch(Exception ex){Message=ex.Message;}}
 public void Test(string dir,string option){testDir=dir;Add(Directory.GetFiles(dir));if(option=="overwrite")conflict.SelectedIndex=1;if(option=="clean")review.SelectedIndex=1;if(option=="markup")review.SelectedIndex=2;Begin(false);if(!busy){File.WriteAllText(Path.Combine(dir,"test-error.txt"),Message);Program.ExitCode=1;Window.Close();}}
}
public static class Program {
 public static int ExitCode;
 [DllImport("dwmapi.dll")]static extern int DwmSetWindowAttribute(IntPtr hwnd,int attr,ref int value,int size);
 [STAThread]public static int Main(string[] args){
  try {
   var app=new Application();var asm=Assembly.GetExecutingAssembly();
   using(var s=asm.GetManifestResourceStream("Theme.xaml"))app.Resources=(ResourceDictionary)XamlReader.Load(s);
   Window window;using(var s=asm.GetManifestResourceStream("MainWindow.xaml"))window=(Window)XamlReader.Load(s);
   var model=new Converter();model.Wire(window);
   window.SourceInitialized+=(s,e)=>{try{int rounded=2;DwmSetWindowAttribute(new WindowInteropHelper(window).Handle,33,ref rounded,4);}catch{}};
   window.ContentRendered+=(s,e)=>{
    if(args.Length>1&&args[0]=="--test"){model.Test(args[1],args.Length>2?args[2]:"auto");return;}
    if(args.Length>1&&args[0]=="--preview"){
     var t=new DispatcherTimer{Interval=TimeSpan.FromMilliseconds(800)};t.Tick+=(a,b)=>{t.Stop();Screenshot(window,args[1]);window.Close();};t.Start();return;
    }
    if(args.Length>0)model.Add(args.Where(p=>File.Exists(p)||Directory.Exists(p)));
   };
   app.Run(window);return ExitCode;
  }catch(Exception ex){File.WriteAllText(Path.Combine(Path.GetTempPath(),"PDF-Switcher-error.txt"),ex.ToString());if(args.Length==0)MessageBox.Show(ex.Message,"PDF Switcher 启动失败");return 1;}
 }
 static void Screenshot(Window window,string file){window.UpdateLayout();var b=new RenderTargetBitmap((int)window.ActualWidth,(int)window.ActualHeight,96,96,PixelFormats.Pbgra32);b.Render(window);var p=new PngBitmapEncoder();p.Frames.Add(BitmapFrame.Create(b));using(var s=File.Create(file))p.Save(s);}
}
