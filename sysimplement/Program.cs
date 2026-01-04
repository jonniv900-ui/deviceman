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
        static string selectedImage = "";
        static string selectedIndex = "1";
        static string selectedDriverPack = "";
        static List<string> logBuffer = new List<string>();
        static bool showConsole = false;
        static string currentAction = "Aguardando...";
        static int progressValue = 0;
        static DiskInfo targetDisk = new DiskInfo { Index = -1 };
        static string bootMode = "UEFI";
        static readonly object _lock = new object();
        static bool isRunning = false; // NOVA: Trava para o menu

        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int vKey);

        [STAThread]
        static void Main(string[] args)
        {
            try { Console.SetWindowSize(100, 32); } catch { }
            Console.Title = "Wtec SysImplement v3.4 - Strict Flow Edition";
            Console.CursorVisible = false;

            ShowSplash();

            while (true)
            {
                if (isRunning) { Thread.Sleep(500); continue; } // Impede menu se estiver rodando

                int menuOpcao = ShowMainMenu();
                if (menuOpcao == 4) break;

                bool proceed = false;
                bool dsk = false, dsm = false, boot = false, drv = false;

                switch (menuOpcao)
                {
                    case 0: proceed = FullDeploymentFlow(); dsk = dsm = boot = true; drv = !string.IsNullOrEmpty(selectedDriverPack); break;
                    case 1: proceed = DiskOnlyFlow(); dsk = true; break;
                    case 2: proceed = BootOnlyFlow(); boot = true; break;
                    case 3: proceed = DriverOnlyFlow(); drv = true; break;
                }

                if (proceed)
                {
                    isRunning = true;
                    RunDeployment(dsk, dsm, boot, drv);

                    // ESSENCIAL: Espera o usuário ler o resultado antes de voltar ao menu
                    Console.SetCursorPosition(18, 18);
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.Write("Pressione qualquer tecla para retornar ao menu...");
                    Console.ReadKey(true);
                    isRunning = false;
                }
            }
        }

        #region MOTOR SÍNCRONO REFORÇADO
        static void RunDeployment(bool doDisk, bool doDism, bool doBoot, bool doDrivers)
        {
            progressValue = 0; logBuffer.Clear(); showConsole = false;

            // Thread de UI permanece para manter o progresso visual
            Thread uiThread = new Thread(() => {
                while (progressValue < 100 || isRunning)
                {
                    if ((GetAsyncKeyState(0x12) & 0x8000) != 0 && (GetAsyncKeyState(0x43) & 0x8000) != 0)
                    {
                        lock (_lock) { showConsole = !showConsole; }
                        Thread.Sleep(500);
                    }
                    RenderUI(); Thread.Sleep(200);
                    if (progressValue >= 100) break;
                }
            })
            { IsBackground = true };
            uiThread.Start();

            try
            {
                string osDrive = "", sysDrive = "";

                if (doDisk)
                {
                    currentAction = "Limpando disco e redefinindo ID...";
                    // Uniqueid ajuda o Windows a não se perder com partições antigas em cache
                    string script = (bootMode == "UEFI")
                        ? $"select disk {targetDisk.Index}\noffline disk\nonline disk\nclean\nconvert gpt\ncreate partition efi size=100\nformat quick fs=fat32 label=\"SYSTEM\"\nassign\ncreate partition msr size=16\ncreate partition primary\nformat quick fs=ntfs label=\"WINDOWS\"\nassign\nrescan"
                        : $"select disk {targetDisk.Index}\noffline disk\nonline disk\nclean\nconvert mbr\ncreate partition primary size=100\nformat quick fs=ntfs label=\"SYSTEM\"\nassign\nactive\ncreate partition primary\nformat quick fs=ntfs label=\"WINDOWS\"\nassign\nrescan";

                    if (!RunDiskpartBlocking(script)) throw new Exception("Diskpart não conseguiu finalizar as operações.");
                    Thread.Sleep(3000);
                }

                currentAction = "Verificando letras de unidades...";
                // Loop de 5 tentativas para montagem
                for (int i = 0; i < 5; i++)
                {
                    foreach (DriveInfo d in DriveInfo.GetDrives().Where(x => x.IsReady))
                    {
                        if (d.VolumeLabel == "WINDOWS") osDrive = d.Name;
                        if (d.VolumeLabel == "SYSTEM") sysDrive = d.Name;
                    }
                    if (!string.IsNullOrEmpty(osDrive)) break;
                    Thread.Sleep(1000);
                }

                if (doDism)
                {
                    if (string.IsNullOrEmpty(osDrive)) throw new Exception("Drive 'WINDOWS' não montou a tempo.");
                    currentAction = "Extraindo imagem para o disco...";
                    ExecuteWithProgress("dism.exe", $"/Apply-Image /ImageFile:\"{selectedImage}\" /Index:{selectedIndex} /ApplyDir:{osDrive}", 0, 80);
                }

                if (doDrivers && !string.IsNullOrEmpty(selectedDriverPack))
                {
                    RestoreDriversOffline(selectedDriverPack, osDrive);
                }

                if (doBoot)
                {
                    if (string.IsNullOrEmpty(sysDrive)) throw new Exception("Drive 'SYSTEM' não disponível para boot.");
                    currentAction = "Criando setor de inicialização...";
                    string sysLetter = sysDrive.Substring(0, 2);
                    ExecuteHidden("bcdboot.exe", $"{osDrive}Windows /s {sysLetter} /f {(bootMode == "UEFI" ? "UEFI" : "BIOS")}");
                }

                progressValue = 100;
                currentAction = "PROCESSO FINALIZADO!";
            }
            catch (Exception ex)
            {
                lock (_lock) { logBuffer.Add("FALHA CRÍTICA: " + ex.Message); }
                progressValue = 100;
                showConsole = true;
            }
        }

        static bool RunDiskpartBlocking(string script)
        {
            string path = Path.Combine(Path.GetTempPath(), "dp.txt");
            File.WriteAllText(path, script);

            ProcessStartInfo psi = new ProcessStartInfo("diskpart.exe", $"/s \"{path}\"")
            {
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true
            };

            using (Process p = Process.Start(psi))
            {
                while (!p.StandardOutput.EndOfStream)
                {
                    string line = p.StandardOutput.ReadLine();
                    lock (_lock) { logBuffer.Add("[DISKPART] " + line); }
                }
                p.WaitForExit();
                return p.ExitCode == 0;
            }
        }
        #endregion

        #region FILE BROWSER COM SCROLL (CORRIGIDO)
        static string FileBrowser(string path, string title, string filter)
        {
            int sel = 0; int scroll = 0; const int limit = 20;
            string curDir = path;

            while (true)
            {
                Console.BackgroundColor = ConsoleColor.DarkGray; Console.Clear();
                DrawBox(2, 1, 95, 28, $" {title} ");

                List<string> items = new List<string>();
                if (string.IsNullOrEmpty(curDir)) items = DriveInfo.GetDrives().Where(d => d.IsReady).Select(d => d.Name).ToList();
                else
                {
                    items.Add(".. [ VOLTAR ]");
                    try
                    {
                        items.AddRange(Directory.GetDirectories(curDir).Select(d => "[" + Path.GetFileName(d) + "]"));
                        items.AddRange(Directory.GetFiles(curDir).Where(f => filter.Split('|').Any(ex => f.EndsWith(ex, StringComparison.OrdinalIgnoreCase))));
                    }
                    catch { items.Add("!! ACESSO NEGADO !!"); }
                }

                if (sel >= scroll + limit) scroll = sel - limit + 1;
                if (sel < scroll) scroll = sel;

                for (int i = 0; i < limit; i++)
                {
                    int idx = scroll + i;
                    if (idx >= items.Count) break;
                    Console.SetCursorPosition(5, 4 + i);
                    if (idx == sel) { Console.BackgroundColor = ConsoleColor.Cyan; Console.ForegroundColor = ConsoleColor.Black; }
                    else { Console.BackgroundColor = ConsoleColor.DarkGray; Console.ForegroundColor = ConsoleColor.White; }
                    Console.Write($"> {Truncate(items[idx], 85).PadRight(87)}");
                }

                var k = Console.ReadKey(true).Key;
                if (k == ConsoleKey.UpArrow && sel > 0) sel--;
                if (k == ConsoleKey.DownArrow && sel < items.Count - 1) sel++;
                if (k == ConsoleKey.Escape) return "";
                if (k == ConsoleKey.Enter)
                {
                    string s = items[sel];
                    if (string.IsNullOrEmpty(curDir)) curDir = s;
                    else if (s == ".. [ VOLTAR ]") curDir = Path.GetDirectoryName(curDir.TrimEnd('\\')) ?? "";
                    else if (s.StartsWith("[")) { curDir = Path.Combine(curDir, s.Trim('[', ']')); sel = 0; scroll = 0; }
                    else if (!s.StartsWith("!!")) return s;
                }
            }
        }
        #endregion

        #region COMPONENTES DE INTERFACE
        static void RenderUI()
        {
            lock (_lock)
            {
                if (showConsole)
                {
                    Console.BackgroundColor = ConsoleColor.Black; Console.Clear();
                    DrawBox(2, 2, 96, 28, " LOGS TÉCNICOS (ALT+C VOLTA) ");
                    int st = Math.Max(0, logBuffer.Count - 24);
                    for (int i = 0; i < Math.Min(logBuffer.Count, 24); i++)
                    {
                        Console.SetCursorPosition(4, 4 + i);
                        Console.Write(Truncate(logBuffer[st + i], 90).PadRight(91));
                    }
                    return;
                }
                Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
                DrawBox(15, 12, 70, 8, " OPERAÇÃO EM CURSO ");
                Console.SetCursorPosition(18, 14); Console.Write("STATUS: " + Truncate(currentAction, 50).PadRight(51));
                int bw = (int)(progressValue * 0.45);
                Console.SetCursorPosition(18, 16);
                Console.ForegroundColor = ConsoleColor.Green; Console.Write(new string('█', bw));
                Console.ForegroundColor = ConsoleColor.Gray; Console.Write(new string('░', 45 - bw) + $" {progressValue}%");
            }
        }

        static int ShowMainMenu()
        {
            int sel = 0; string[] items = { " [1] IMPLANTAR WINDOWS ", " [2] FORMATAR DISCO ", " [3] REPARAR BOOT ", " [4] INJETAR DRIVERS ", " [5] SAIR " };
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

        static string GetImageIndex(string path)
        {
            Console.Clear(); DrawBox(5, 5, 90, 20, " LEITURA DA IMAGEM ");
            var psi = new ProcessStartInfo("dism.exe", $"/Get-WimInfo /WimFile:\"{path}\"") { UseShellExecute = false, RedirectStandardOutput = true, CreateNoWindow = true };
            var p = Process.Start(psi); string outD = p.StandardOutput.ReadToEnd(); p.WaitForExit();
            string[] lines = outD.Split('\n'); int y = 7;
            foreach (var l in lines) if (l.Contains("Index") || l.Contains("Índice") || l.Contains("Name") || l.Contains("Nome")) if (y < 24) { Console.SetCursorPosition(8, y++); Console.Write(l.Trim()); }
            Console.SetCursorPosition(8, 26); Console.Write("Informe o Índice: "); Console.CursorVisible = true; string r = Console.ReadLine(); Console.CursorVisible = false;
            return string.IsNullOrEmpty(r) ? "1" : r;
        }

        static void ShowSplash()
        {
            Console.Clear(); Console.ForegroundColor = ConsoleColor.Cyan;
            string[] logo = {
                "  ██╗    ██╗████████╗███████╗ ██████╗  ",
                "  ██║    ██║╚══██╔══╝██╔════╝██╔════╝  ",
                "  ██║ █╗ ██║   ██║   █████╗  ██║       ",
                "  ██║███╗██║   ██║   ██╔══╝  ██║       ",
                "  ╚███╔███╔╝   ██║   ███████╗╚██████╗  ",
                "   ╚══╝╚══╝    ╚═╝   ╚══════╝ ╚═════╝  ",
                "       TECNOLOGIA E SISTEMAS @ 2026    "
            };
            int y = 5; foreach (string l in logo) { Console.SetCursorPosition((100 - l.Length) / 2, y++); Console.WriteLine(l); }
            Thread.Sleep(1000);
        }

        static int DiskSelector()
        {
            int sel = 0; var dsks = new List<DiskInfo>();
            ManagementObjectSearcher s = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive");
            foreach (ManagementObject d in s.Get()) dsks.Add(new DiskInfo { Index = Convert.ToInt32(d["Index"]), Model = d["Model"].ToString(), SizeGB = Convert.ToDouble(d["Size"]) / 1073741824.0 });
            while (true)
            {
                Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
                DrawBox(10, 5, 80, 20, " SELECIONE O DISCO ");
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

        static bool FullDeploymentFlow()
        {
            selectedImage = FileBrowser("", "Imagem do Windows", ".wim|.esd");
            if (string.IsNullOrEmpty(selectedImage)) return false;
            selectedIndex = GetImageIndex(selectedImage);
            if (ConfirmAction("Injetar Pack de Drivers?")) selectedDriverPack = FileBrowser("", "Drivers", ".drvbackup");
            if (DiskSelector() == -1) return false;
            bootMode = BootModeSelector();
            return ShowSummary("GERAL");
        }

        static bool DiskOnlyFlow() { if (DiskSelector() == -1) return false; bootMode = BootModeSelector(); return ShowSummary("DISCO"); }
        static bool BootOnlyFlow() { if (DiskSelector() == -1) return false; bootMode = BootModeSelector(); return ShowSummary("BOOT"); }
        static bool DriverOnlyFlow() { selectedDriverPack = FileBrowser("", "Drivers", ".drvbackup"); if (string.IsNullOrEmpty(selectedDriverPack) || DiskSelector() == -1) return false; return ShowSummary("DRIVERS"); }
        static void DrawBox(int x, int y, int w, int h, string t) { Console.ForegroundColor = ConsoleColor.White; Console.SetCursorPosition(x, y); Console.Write("╔" + new string('═', w - 2) + "╗"); for (int i = 1; i < h - 1; i++) { Console.SetCursorPosition(x, y + i); Console.Write("║" + new string(' ', w - 2) + "║"); } Console.SetCursorPosition(x, y + h - 1); Console.Write("╚" + new string('═', w - 2) + "╝"); if (!string.IsNullOrEmpty(t)) { Console.SetCursorPosition(x + (w / 2) - (t.Length / 2), y); Console.Write(t); } }
        static bool ConfirmAction(string m) { Console.Clear(); DrawBox(20, 12, 60, 6, " CONFIRMAR "); Console.SetCursorPosition(23, 14); Console.Write(m); Console.SetCursorPosition(23, 16); Console.Write("[S] SIM | [N] NÃO"); return Console.ReadKey(true).Key == ConsoleKey.S; }
        static string BootModeSelector() { int sel = 0; string[] m = { "UEFI (GPT)", "LEGACY (MBR)" }; while (true) { Console.Clear(); DrawBox(35, 12, 30, 6, " BOOT "); for (int i = 0; i < 2; i++) { Console.SetCursorPosition(38, 14 + i); if (sel == i) { Console.BackgroundColor = ConsoleColor.Cyan; Console.ForegroundColor = ConsoleColor.Black; } else { Console.BackgroundColor = ConsoleColor.DarkBlue; Console.ForegroundColor = ConsoleColor.White; } Console.Write(m[i]); } var k = Console.ReadKey(true).Key; if (k == ConsoleKey.UpArrow || k == ConsoleKey.DownArrow) sel = 1 - sel; if (k == ConsoleKey.Enter) return (sel == 0) ? "UEFI" : "LEGACY"; } }
        static bool ShowSummary(string t) { Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear(); DrawBox(20, 10, 60, 11, $" {t} "); Console.SetCursorPosition(23, 13); Console.Write("DISCO: " + targetDisk.Model); Console.SetCursorPosition(23, 14); Console.Write("MODO:  " + bootMode); Console.SetCursorPosition(23, 15); Console.Write("WIM:   " + Path.GetFileName(selectedImage)); Console.ForegroundColor = ConsoleColor.Red; Console.SetCursorPosition(23, 17); Console.Write("O DISCO SERÁ FORMATADO!"); Console.ForegroundColor = ConsoleColor.White; Console.SetCursorPosition(23, 19); Console.Write("[ENTER] INICIAR | [ESC] SAIR"); return Console.ReadKey(true).Key == ConsoleKey.Enter; }
        static void ExecuteHidden(string e, string a) { Process.Start(new ProcessStartInfo(e, a) { CreateNoWindow = true, UseShellExecute = false }).WaitForExit(); }
        static void RestoreDriversOffline(string p, string os) { string tmp = Path.Combine(Path.GetTempPath(), "WtecDrivers"); if (Directory.Exists(tmp)) Directory.Delete(tmp, true); Directory.CreateDirectory(tmp); using (ZipArchive arc = ZipFile.OpenRead(p)) { foreach (var en in arc.Entries) { string f = Path.Combine(tmp, en.FullName); if (!Directory.Exists(Path.GetDirectoryName(f))) Directory.CreateDirectory(Path.GetDirectoryName(f)); if (!string.IsNullOrEmpty(en.Name)) en.ExtractToFile(f, true); } } ExecuteWithProgress("dism.exe", $"/Image:{os} /Add-Driver /Driver:\"{tmp}\" /Recurse", 85, 10); }
        static void ExecuteWithProgress(string exe, string args, int baseP, int w) { var psi = new ProcessStartInfo(exe, args) { CreateNoWindow = true, UseShellExecute = false, RedirectStandardOutput = true }; using (var p = Process.Start(psi)) { while (!p.StandardOutput.EndOfStream) { string l = p.StandardOutput.ReadLine(); if (l.Contains("%")) { var d = new string(l.Where(char.IsDigit).ToArray()); if (int.TryParse(d, out int v)) progressValue = baseP + (int)(v * (w / 100.0)); } lock (_lock) logBuffer.Add(l); } p.WaitForExit(); } }
        static string Truncate(string s, int m) => s.Length <= m ? s : s.Substring(0, m - 3) + "...";
        struct DiskInfo { public int Index; public string Model; public double SizeGB; }
        #endregion
    }
}