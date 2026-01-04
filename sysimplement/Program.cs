using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;

namespace SysImplement
{
    class Program
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool Wow64DisableWow64FsRedirection(ref IntPtr ptr);
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool Wow64RevertWow64FsRedirection(IntPtr ptr);
        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int vKey);

        static string selectedImage = "";
        static string selectedIndex = "1";
        static string selectedDriverPack = "";
        static List<string> logBuffer = new List<string>();
        static bool showConsole = false;
        static string currentAction = "Aguardando...";
        static int progressValue = 0;
        static DiskInfo targetDisk;
        static string bootMode = "UEFI";
        static readonly object _lock = new object();

        [STAThread]
        static void Main(string[] args)
        {
            try { Console.SetWindowSize(100, 32); } catch { }
            Console.Title = "Wtec SysImplement v2.0";
            Console.CursorVisible = false;
            ShowSplash();

            while (true)
            {
                int menuOpcao = ShowMainMenu();
                if (menuOpcao == 4) break;
                try
                {
                    switch (menuOpcao)
                    {
                        case 0: if (FullDeploymentFlow()) RunDeployment(true, true, true, !string.IsNullOrEmpty(selectedDriverPack)); break;
                        case 1: if (DiskOnlyFlow()) RunDeployment(true, false, false, false); break;
                        case 2: if (BootOnlyFlow()) RunDeployment(false, false, true, false); break;
                        case 3: if (DriverOnlyFlow()) RunDeployment(false, false, false, true); break;
                    }
                }
                catch (Exception ex)
                {
                    lock (_lock) { logBuffer.Add("ERRO: " + ex.Message); }
                    showConsole = true;
                    RenderUI();
                    Console.ReadKey();
                }
            }
        }

        #region FLUXO DE EXECUCAO
        static void RunDeployment(bool doDisk, bool doDism, bool doBoot, bool doDrivers)
        {
            progressValue = 0; logBuffer.Clear(); showConsole = false;
            Thread uiThread = new Thread(() => {
                while (progressValue < 100)
                {
                    if ((GetAsyncKeyState(0x12) & 0x8000) != 0 && (GetAsyncKeyState(0x43) & 0x8000) != 0)
                    {
                        lock (_lock) { showConsole = !showConsole; }
                        Thread.Sleep(500);
                    }
                    RenderUI(); Thread.Sleep(200);
                }
            })
            { IsBackground = true };
            uiThread.Start();

            try
            {
                string osDrive = "", sysDrive = "";
                if (doDisk)
                {
                    currentAction = "Particionando...";
                    StringBuilder ds = new StringBuilder();
                    ds.AppendLine($"select disk {targetDisk.Index}");
                    ds.AppendLine("clean");
                    if (bootMode == "UEFI")
                    {
                        ds.AppendLine("convert gpt");
                        ds.AppendLine("create partition efi size=100");
                        ds.AppendLine("format quick fs=fat32 label=\"SYSTEM\"");
                        ds.AppendLine("assign");
                    }
                    else
                    {
                        ds.AppendLine("convert mbr");
                        ds.AppendLine("create partition primary size=100");
                        ds.AppendLine("format quick fs=ntfs label=\"SYSTEM\"");
                        ds.AppendLine("active");
                        ds.AppendLine("assign");
                    }
                    ds.AppendLine("create partition primary");
                    ds.AppendLine("format quick fs=ntfs label=\"WINDOWS\"");
                    ds.AppendLine("assign");
                    RunDiskpartScript(ds.ToString());
                }

                Thread.Sleep(2000);
                foreach (DriveInfo d in DriveInfo.GetDrives().Where(x => x.IsReady))
                {
                    if (d.VolumeLabel == "WINDOWS") osDrive = d.Name;
                    if (d.VolumeLabel == "SYSTEM") sysDrive = d.Name;
                }

                if (doDism && !string.IsNullOrEmpty(osDrive))
                {
                    currentAction = "Aplicando Windows...";
                    ExecuteWithProgress("dism.exe", $"/Apply-Image /ImageFile:\"{selectedImage}\" /Index:{selectedIndex} /ApplyDir:{osDrive}");
                }

                if (doDrivers && !string.IsNullOrEmpty(selectedDriverPack)) RestoreDriversOffline(selectedDriverPack, osDrive);

                if (doBoot && !string.IsNullOrEmpty(osDrive))
                {
                    currentAction = "Gravando Bootloader...";
                    string winPath = Path.Combine(osDrive.Substring(0, 2) + "\\", "Windows");
                    string sysLet = string.IsNullOrEmpty(sysDrive) ? osDrive.Substring(0, 2) : sysDrive.Substring(0, 2);
                    string bcdArgs = $"\"{winPath}\" /s {sysLet} /f {(bootMode == "UEFI" ? "UEFI" : "BIOS")}";
                    ExecuteNativeProcess("bcdboot.exe", bcdArgs);
                }

                ExecuteNativeProcess("dism.exe", "/Cleanup-Wim");
                progressValue = 100;
                currentAction = "FINALIZADO!";
                lock (_lock) { showConsole = true; }
            }
            catch (Exception ex) { lock (_lock) { logBuffer.Add("FALHA: " + ex.Message); } progressValue = 100; showConsole = true; }
            RenderUI();
            Console.ReadKey(true);
        }
        #endregion

        #region UTILITARIOS DE PROCESSO
        static void ExecuteNativeProcess(string exe, string args)
        {
            IntPtr ptr = IntPtr.Zero;
            bool is64 = Environment.Is64BitOperatingSystem;
            if (is64) Wow64DisableWow64FsRedirection(ref ptr);
            try
            {
                var psi = new ProcessStartInfo(exe, args) { CreateNoWindow = true, UseShellExecute = false };
                using (var p = Process.Start(psi)) { p?.WaitForExit(); }
            }
            finally { if (is64) Wow64RevertWow64FsRedirection(ptr); }
        }

        static void ExecuteWithProgress(string exe, string args)
        {
            IntPtr ptr = IntPtr.Zero;
            bool is64 = Environment.Is64BitOperatingSystem;
            if (is64) Wow64DisableWow64FsRedirection(ref ptr);
            try
            {
                var psi = new ProcessStartInfo(exe, args) { CreateNoWindow = true, UseShellExecute = false, RedirectStandardOutput = true };
                using (var proc = Process.Start(psi))
                {
                    while (!proc.StandardOutput.EndOfStream)
                    {
                        string l = proc.StandardOutput.ReadLine();
                        if (string.IsNullOrEmpty(l)) continue;
                        lock (_lock)
                        {
                            logBuffer.Add(l);
                            if (l.Contains("%"))
                            {
                                var digit = new string(l.Where(char.IsDigit).ToArray());
                                if (int.TryParse(digit, out int v)) progressValue = 20 + (int)(v * 0.75);
                            }
                        }
                    }
                    proc.WaitForExit();
                }
            }
            finally { if (is64) Wow64RevertWow64FsRedirection(ptr); }
        }

        static void RestoreDriversOffline(string drvBackupPath, string targetOsPath)
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "WtecDrivers");
            try
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
                Directory.CreateDirectory(tempDir);
                currentAction = "Extraindo drivers...";
                ZipFile.ExtractToDirectory(drvBackupPath, tempDir);
                currentAction = "Injetando drivers...";
                ExecuteWithProgress("dism.exe", $"/Image:{targetOsPath} /Add-Driver /Driver:\"{tempDir}\" /Recurse /ForceUnsigned");
            }
            catch { }
            finally { try { Directory.Delete(tempDir, true); } catch { } }
        }

        static void RunDiskpartScript(string script)
        {
            string p = Path.Combine(Path.GetTempPath(), "dp.txt"); File.WriteAllText(p, script);
            ExecuteNativeProcess("diskpart.exe", $"/s \"{p}\"");
        }
        #endregion

        #region INTERFACE E NAVEGACAO
        static void DrawBox(int x, int y, int w, int h, string t)
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.SetCursorPosition(x, y); Console.Write("╔" + new string('═', w - 2) + "╗");
            for (int i = 1; i < h - 1; i++) { Console.SetCursorPosition(x, y + i); Console.Write("║" + new string(' ', w - 2) + "║"); }
            Console.SetCursorPosition(x, y + h - 1); Console.Write("╚" + new string('═', w - 2) + "╝");
            if (!string.IsNullOrEmpty(t)) { Console.SetCursorPosition(x + (w / 2) - (t.Length / 2), y); Console.Write(t); }
        }

        static string Truncate(string s, int m) => s.Length <= m ? s : s.Substring(0, m - 3) + "...";

        static void RenderUI()
        {
            lock (_lock)
            {
                if (showConsole)
                {
                    Console.BackgroundColor = ConsoleColor.Black; Console.Clear();
                    DrawBox(2, 2, 96, 28, " LOGS (ALT+C VOLTAR) ");
                    int st = Math.Max(0, logBuffer.Count - 24);
                    for (int i = 0; i < Math.Min(logBuffer.Count, 24); i++)
                    {
                        Console.SetCursorPosition(4, 4 + i);
                        Console.Write(Truncate(logBuffer[st + i], 90).PadRight(91));
                    }
                    return;
                }
                Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
                DrawBox(15, 12, 70, 8, " PROCESSANDO ");
                Console.SetCursorPosition(18, 14); Console.Write("AÇÃO: " + Truncate(currentAction, 50).PadRight(51));
                Console.SetCursorPosition(18, 16);
                int bw = (int)(progressValue * 0.45);
                Console.ForegroundColor = ConsoleColor.Green; Console.Write(new string('█', bw));
                Console.ForegroundColor = ConsoleColor.Gray; Console.Write(new string('░', 45 - bw) + $" {progressValue}%");
            }
        }

        static int ShowMainMenu()
        {
            int sel = 0; string[] items = { " [1] INSTALAR COMPLETO ", " [2] APENAS DISCO ", " [3] REPARAR BOOT ", " [4] INJETAR DRIVERS ", " [5] SAIR " };
            while (true)
            {
                Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
                DrawBox(25, 10, 50, 11, " MENU PRINCIPAL ");
                for (int i = 0; i < items.Length; i++)
                {
                    Console.SetCursorPosition(27, 13 + i);
                    if (i == sel) { Console.BackgroundColor = ConsoleColor.Cyan; Console.ForegroundColor = ConsoleColor.Black; }
                    else { Console.BackgroundColor = ConsoleColor.DarkBlue; Console.ForegroundColor = ConsoleColor.White; }
                    Console.Write(items[i].PadRight(46));
                }
                var k = Console.ReadKey(true).Key;
                if (k == ConsoleKey.UpArrow && sel > 0) sel--;
                if (k == ConsoleKey.DownArrow && sel < items.Length - 1) sel++;
                if (k == ConsoleKey.Enter) return sel;
            }
        }

        static string FileBrowser(string startPath, string title, string filter)
        {
            string path = startPath; int sel = 0; int scroll = 0;
            const int maxView = 20;

            while (true)
            {
                Console.BackgroundColor = ConsoleColor.DarkGray; Console.Clear();
                DrawBox(2, 1, 95, 28, $" {title} ");
                List<string> ent = new List<string>();
                if (string.IsNullOrEmpty(path)) ent = DriveInfo.GetDrives().Where(d => d.IsReady).Select(d => d.Name).ToList();
                else
                {
                    ent.Add(".. [ VOLTAR ]");
                    try
                    {
                        ent.AddRange(Directory.GetDirectories(path).Select(d => "[" + Path.GetFileName(d) + "]"));
                        ent.AddRange(Directory.GetFiles(path).Where(f => filter.Contains(Path.GetExtension(f).ToLower())).Select(Path.GetFileName));
                    }
                    catch { }
                }

                if (sel < scroll) scroll = sel;
                if (sel >= scroll + maxView) scroll = sel - maxView + 1;

                for (int i = 0; i < Math.Min(ent.Count - scroll, maxView); i++)
                {
                    int idx = i + scroll;
                    Console.SetCursorPosition(5, 3 + i);
                    if (idx == sel) { Console.BackgroundColor = ConsoleColor.Cyan; Console.ForegroundColor = ConsoleColor.Black; }
                    else { Console.BackgroundColor = ConsoleColor.DarkGray; Console.ForegroundColor = ConsoleColor.White; }
                    Console.Write($"> {Truncate(ent[idx], 88).PadRight(90)}");
                }

                var k = Console.ReadKey(true).Key;
                if (k == ConsoleKey.Escape) return "";
                if (k == ConsoleKey.UpArrow && sel > 0) sel--;
                if (k == ConsoleKey.DownArrow && sel < ent.Count - 1) sel++;
                if (k == ConsoleKey.Enter)
                {
                    string s = ent[sel];
                    if (string.IsNullOrEmpty(path)) { path = s; sel = 0; scroll = 0; }
                    else if (s == ".. [ VOLTAR ]") { path = Path.GetDirectoryName(path.TrimEnd('\\')) ?? ""; sel = 0; scroll = 0; }
                    else if (s.StartsWith("[")) { path = Path.Combine(path, s.Trim('[', ']')); sel = 0; scroll = 0; }
                    else return Path.Combine(path, s);
                }
            }
        }

        static string BootModeSelector()
        {
            while (true)
            {
                Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
                DrawBox(30, 10, 40, 8, " MODO DE BOOT ");
                Console.SetCursorPosition(33, 13); Console.Write("[1] UEFI (Padrao GPT)");
                Console.SetCursorPosition(33, 14); Console.Write("[2] LEGACY (BIOS/MBR)");
                Console.SetCursorPosition(33, 16); Console.Write("Escolha: ");
                var k = Console.ReadKey(true).KeyChar;
                if (k == '1') return "UEFI";
                if (k == '2') return "BIOS";
            }
        }

        static string GetImageIndex(string path)
        {
            Console.Clear(); DrawBox(5, 2, 90, 26, " VERSOES DISPONIVEIS ");
            IntPtr ptr = IntPtr.Zero;
            bool is64 = Environment.Is64BitOperatingSystem;
            if (is64) Wow64DisableWow64FsRedirection(ref ptr);

            string outD = "";
            try
            {
                var psi = new ProcessStartInfo("dism.exe", $"/Get-WimInfo /WimFile:\"{path}\"") { UseShellExecute = false, RedirectStandardOutput = true, CreateNoWindow = true };
                using (var p = Process.Start(psi)) { outD = p.StandardOutput.ReadToEnd(); p.WaitForExit(); }
            }
            finally { if (is64) Wow64RevertWow64FsRedirection(ptr); }

            string[] lines = outD.Split('\n'); int y = 4;
            foreach (var line in lines)
            {
                if ((line.Contains("Index") || line.Contains("Name") || line.Contains("Índice") || line.Contains("Nome")) && !line.Contains("Informações"))
                {
                    Console.SetCursorPosition(8, y++);
                    Console.Write(line.Trim());
                    if (y > 24) break;
                }
            }
            Console.SetCursorPosition(5, 28); Console.Write("Digite o numero do Index: ");
            Console.CursorVisible = true; string r = Console.ReadLine(); Console.CursorVisible = false;
            return string.IsNullOrEmpty(r) ? "1" : r;
        }

        // Restante dos métodos (DiskSelector, ConfirmAction, etc) mantidos da versão funcional
        static int DiskSelector()
        {
            int sel = 0;
            while (true)
            {
                var dsks = new List<DiskInfo>();
                ManagementObjectSearcher s = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive");
                foreach (ManagementObject d in s.Get()) dsks.Add(new DiskInfo { Index = Convert.ToInt32(d["Index"]), Model = d["Model"].ToString(), SizeGB = Convert.ToDouble(d["Size"]) / 1073741824.0 });
                Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear(); DrawBox(10, 5, 80, 20, " SELECAO DE DISCO ");
                for (int i = 0; i < dsks.Count; i++)
                {
                    Console.SetCursorPosition(13, 8 + i);
                    if (i == sel) { Console.BackgroundColor = ConsoleColor.Cyan; Console.ForegroundColor = ConsoleColor.Black; }
                    else { Console.BackgroundColor = ConsoleColor.DarkBlue; Console.ForegroundColor = ConsoleColor.White; }
                    Console.Write($"> DISCO {dsks[i].Index}: {dsks[i].Model} ({dsks[i].SizeGB:F1} GB)".PadRight(70));
                }
                var k = Console.ReadKey(true).Key;
                if (k == ConsoleKey.Enter) { targetDisk = dsks[sel]; return dsks[sel].Index; }
                if (k == ConsoleKey.UpArrow && sel > 0) sel--;
                if (k == ConsoleKey.DownArrow && sel < dsks.Count - 1) sel++;
                if (k == ConsoleKey.Escape) return -1;
            }
        }
        static bool FullDeploymentFlow() { selectedImage = FileBrowser("", "Selecionar Imagem Windows", ".wim.esd"); if (string.IsNullOrEmpty(selectedImage)) return false; selectedIndex = GetImageIndex(selectedImage); if (ConfirmAction("Deseja injetar drivers?")) selectedDriverPack = FileBrowser("", "Selecionar Pack de Drivers", ".drvbackup"); if (DiskSelector() == -1) return false; bootMode = BootModeSelector(); return ShowSummary("INSTALACAO COMPLETA"); }
        static bool DiskOnlyFlow() { if (DiskSelector() == -1) return false; bootMode = BootModeSelector(); return ShowSummary("APENAS DISCO"); }
        static bool BootOnlyFlow() { if (DiskSelector() == -1) return false; bootMode = BootModeSelector(); return ShowSummary("REPARAR BOOT"); }
        static bool DriverOnlyFlow() { selectedDriverPack = FileBrowser("", "Selecionar Drivers", ".drvbackup"); if (string.IsNullOrEmpty(selectedDriverPack)) return false; if (DiskSelector() == -1) return false; return ShowSummary("INJETAR DRIVERS"); }
        static bool ConfirmAction(string msg) { Console.Clear(); DrawBox(20, 12, 60, 6, " CONFIRMAR "); Console.SetCursorPosition(22, 14); Console.Write(Truncate(msg, 56)); Console.SetCursorPosition(38, 16); Console.Write("[S] SIM  [N] NAO"); return Console.ReadKey(true).Key == ConsoleKey.S; }
        static bool ShowSummary(string title) { Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear(); DrawBox(20, 10, 60, 11, $" {title} "); Console.SetCursorPosition(23, 13); Console.Write("DISCO: " + targetDisk.Model); Console.SetCursorPosition(23, 14); Console.Write("BOOT:  " + bootMode); Console.SetCursorPosition(0, 31); Console.BackgroundColor = ConsoleColor.Gray; Console.ForegroundColor = ConsoleColor.Black; Console.Write(" [ENTER] INICIAR | [ESC] SAIR ".PadRight(100)); return Console.ReadKey(true).Key == ConsoleKey.Enter; }
        static void ShowSplash() { Console.Clear(); Console.ForegroundColor = ConsoleColor.Cyan; Console.SetCursorPosition(35, 12); Console.Write("Wtec SysImplement v2.9.9"); Thread.Sleep(800); }
        struct DiskInfo { public int Index; public string Model; public double SizeGB; }
        #endregion
    }
}