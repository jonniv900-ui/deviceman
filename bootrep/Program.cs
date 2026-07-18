using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;

namespace WtecBootRepair
{
    class Program
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool Wow64DisableWow64FsRedirection(ref IntPtr ptr);
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool Wow64RevertWow64FsRedirection(IntPtr ptr);

        // API oficial (Windows 8+) para saber se o firmware é BIOS (Legacy) ou UEFI -
        // muito mais confiável do que tentar inferir isso por heurística.
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool GetFirmwareType(out FirmwareType firmwareType);

        public enum FirmwareType { Unknown = 0, Bios = 1, Uefi = 2, Max = 3 }

        struct WindowsInstall { public string Drive; public string Version; public string Edition; }
        struct DiskInfo { public int Index; public string Model; public double SizeGB; }

        [STAThread]
        static void Main()
        {
            try { Console.SetWindowSize(100, 32); } catch { }
            Console.Title = "Wtec BootRepair v1.0";
            Console.CursorVisible = false;
            ShowSplash();

            GetFirmwareType(out FirmwareType firmware);
            if (firmware == FirmwareType.Unknown || firmware == FirmwareType.Max)
                firmware = FirmwareType.Uefi; // fallback: a maioria das máquinas atuais é UEFI

            while (true)
            {
                var installs = ScanWindowsInstallations();
                int action = ShowMainMenu(firmware, installs);
                if (action == -1) break;

                switch (action)
                {
                    case 0: ReconstructBootloaderFlow(firmware, installs); break;
                    case 1: MarkPartitionActiveFlow(); break;
                    case 2: BackupBcdFlow(); break;
                }
            }
        }

        #region TELA PRINCIPAL
        static int ShowMainMenu(FirmwareType firmware, List<WindowsInstall> installs)
        {
            int sel = 0;
            string[] items =
            {
                " RECONSTRUIR BOOTLOADER (bcdboot) ",
                " MARCAR PARTICAO COMO ATIVA (Legacy/BIOS) ",
                " BACKUP DO BCD ATUAL ",
                " SAIR "
            };

            while (true)
            {
                Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
                DrawBox(10, 2, 80, 24, " WTEC BOOTREPAIR ");

                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.SetCursorPosition(13, 4);
                Console.Write($"Firmware detectado: {(firmware == FirmwareType.Uefi ? "UEFI" : "LEGACY (BIOS)")}");

                Console.ForegroundColor = ConsoleColor.White;
                Console.SetCursorPosition(13, 6);
                Console.Write($"Instalações do Windows encontradas: {installs.Count}");
                for (int i = 0; i < installs.Count && i < 4; i++)
                {
                    Console.SetCursorPosition(15, 7 + i);
                    Console.Write($"- {installs[i].Drive} (build {installs[i].Version})");
                }

                int menuStart = 13;
                for (int i = 0; i < items.Length; i++)
                {
                    Console.SetCursorPosition(13, menuStart + i);
                    if (i == sel) { Console.BackgroundColor = ConsoleColor.Cyan; Console.ForegroundColor = ConsoleColor.Black; }
                    else { Console.BackgroundColor = ConsoleColor.DarkBlue; Console.ForegroundColor = ConsoleColor.White; }
                    Console.Write(items[i].PadRight(60));
                }

                DrawFooter(" [SETAS] Navegar | [ENTER] Selecionar | [ESC] Sair ");
                var k = Console.ReadKey(true).Key;
                if (k == ConsoleKey.Escape) return -1;
                if (k == ConsoleKey.UpArrow && sel > 0) sel--;
                if (k == ConsoleKey.DownArrow && sel < items.Length - 1) sel++;
                if (k == ConsoleKey.Enter) return sel == 3 ? -1 : sel;
            }
        }
        #endregion

        #region RECONSTRUÇÃO DO BOOTLOADER
        static void ReconstructBootloaderFlow(FirmwareType firmware, List<WindowsInstall> installs)
        {
            if (installs.Count == 0)
            {
                ShowMessage("Nenhuma instalação do Windows foi encontrada nos discos conectados.", ConsoleColor.DarkRed);
                return;
            }

            WindowsInstall target = installs.Count == 1 ? installs[0] : SelectInstall(installs);
            if (string.IsNullOrEmpty(target.Drive)) return;

            string sysLetter;
            if (firmware == FirmwareType.Uefi)
            {
                string found = FindEfiSystemPartition();
                if (found == null)
                {
                    if (!ConfirmAction("Partição EFI (FAT32, pasta \\EFI\\Microsoft\\Boot) não foi encontrada. " +
                                       "Usar o próprio drive do Windows como destino mesmo assim?")) return;
                    sysLetter = target.Drive;
                }
                else sysLetter = found;
            }
            else
            {
                // Legacy/BIOS: normalmente não há letra própria para a partição de sistema;
                // gravamos direto no drive do Windows, que é o comportamento padrão do bcdboot.
                sysLetter = target.Drive;
            }

            if (!ConfirmAction($"Reconstruir o bootloader para {target.Drive} usando modo " +
                               $"{(firmware == FirmwareType.Uefi ? "UEFI" : "LEGACY")} (sistema: {sysLetter})?")) return;

            BackupBcd(sysLetter, silent: true);

            string winPath = Path.Combine(target.Drive, "Windows");
            string args = $"\"{winPath}\" /s {sysLetter.TrimEnd('\\')} /f {(firmware == FirmwareType.Uefi ? "UEFI" : "BIOS")}";
            var output = ExecuteCaptured("bcdboot.exe", args);

            ShowLogResult(" RESULTADO - bcdboot ", output);
        }

        static WindowsInstall SelectInstall(List<WindowsInstall> installs)
        {
            int sel = 0;
            while (true)
            {
                Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
                DrawBox(15, 8, 70, 14, " SELECIONE A INSTALACAO DO WINDOWS ");
                for (int i = 0; i < installs.Count; i++)
                {
                    Console.SetCursorPosition(18, 11 + i);
                    if (i == sel) { Console.BackgroundColor = ConsoleColor.Cyan; Console.ForegroundColor = ConsoleColor.Black; }
                    else { Console.BackgroundColor = ConsoleColor.DarkBlue; Console.ForegroundColor = ConsoleColor.White; }
                    Console.Write($"> {installs[i].Drive} - build {installs[i].Version}".PadRight(60));
                }
                DrawFooter(" [SETAS] Navegar | [ENTER] Selecionar | [ESC] Cancelar ");
                var k = Console.ReadKey(true).Key;
                if (k == ConsoleKey.Escape) return default;
                if (k == ConsoleKey.UpArrow && sel > 0) sel--;
                if (k == ConsoleKey.DownArrow && sel < installs.Count - 1) sel++;
                if (k == ConsoleKey.Enter) return installs[sel];
            }
        }

        static List<WindowsInstall> ScanWindowsInstallations()
        {
            var results = new List<WindowsInstall>();
            foreach (var d in DriveInfo.GetDrives())
            {
                if (!d.IsReady) continue;
                try
                {
                    string kernelPath = Path.Combine(d.RootDirectory.FullName, "Windows", "System32", "ntoskrnl.exe");
                    if (File.Exists(kernelPath))
                    {
                        string ver = "N/A";
                        try { ver = FileVersionInfo.GetVersionInfo(kernelPath).FileVersion ?? "N/A"; } catch { }
                        results.Add(new WindowsInstall { Drive = d.Name, Version = ver, Edition = "" });
                    }
                }
                catch { }
            }
            return results;
        }

        /// <summary>Procura por uma partição FAT32 com a estrutura \EFI\Microsoft\Boot,
        /// que é a marca registrada de uma partição de sistema EFI válida.</summary>
        static string FindEfiSystemPartition()
        {
            foreach (var d in DriveInfo.GetDrives())
            {
                if (!d.IsReady) continue;
                try
                {
                    if (Directory.Exists(Path.Combine(d.RootDirectory.FullName, "EFI", "Microsoft", "Boot")))
                        return d.Name;
                    if (string.Equals(d.VolumeLabel, "SYSTEM", StringComparison.OrdinalIgnoreCase))
                        return d.Name;
                }
                catch { }
            }
            return null;
        }
        #endregion

        #region MARCAR PARTIÇÃO ATIVA (LEGACY)
        static void MarkPartitionActiveFlow()
        {
            var disks = ListDisks();
            if (disks.Count == 0) { ShowMessage("Nenhum disco encontrado.", ConsoleColor.DarkRed); return; }

            int diskIndex = SelectDisk(disks);
            if (diskIndex == -1) return;

            var partitions = RunDiskpartCapture($"select disk {diskIndex}\r\nlist partition\r\n");
            Console.BackgroundColor = ConsoleColor.Black; Console.Clear();
            DrawBox(2, 2, 96, 22, $" DISCO {diskIndex} - PARTICOES ");
            int y = 4;
            foreach (var line in partitions)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                Console.SetCursorPosition(4, y++);
                Console.ForegroundColor = ConsoleColor.White;
                Console.Write(Truncate(line.Trim(), 90));
                if (y > 22) break;
            }
            Console.SetCursorPosition(4, 24); Console.Write("Número da partição a marcar como ativa: ");
            Console.CursorVisible = true;
            string part = Console.ReadLine() ?? "";
            Console.CursorVisible = false;
            if (string.IsNullOrWhiteSpace(part)) return;

            if (!ConfirmAction($"Marcar a partição {part} do disco {diskIndex} como ATIVA?")) return;
            var result = RunDiskpartCapture($"select disk {diskIndex}\r\nselect partition {part}\r\nactive\r\n");
            ShowLogResult(" RESULTADO - marcar ativa ", result);
        }

        static int SelectDisk(List<DiskInfo> disks)
        {
            int sel = 0;
            while (true)
            {
                Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
                DrawBox(10, 6, 80, 16, " SELECIONE O DISCO ");
                for (int i = 0; i < disks.Count; i++)
                {
                    Console.SetCursorPosition(13, 9 + i);
                    if (i == sel) { Console.BackgroundColor = ConsoleColor.Cyan; Console.ForegroundColor = ConsoleColor.Black; }
                    else { Console.BackgroundColor = ConsoleColor.DarkBlue; Console.ForegroundColor = ConsoleColor.White; }
                    Console.Write($"> DISCO {disks[i].Index}: {Truncate(disks[i].Model, 50)} ({disks[i].SizeGB:F1} GB)".PadRight(70));
                }
                DrawFooter(" [SETAS] Navegar | [ENTER] Selecionar | [ESC] Cancelar ");
                var k = Console.ReadKey(true).Key;
                if (k == ConsoleKey.Escape) return -1;
                if (k == ConsoleKey.UpArrow && sel > 0) sel--;
                if (k == ConsoleKey.DownArrow && sel < disks.Count - 1) sel++;
                if (k == ConsoleKey.Enter) return disks[sel].Index;
            }
        }

        static List<DiskInfo> ListDisks()
        {
            var disks = new List<DiskInfo>();
            try
            {
                using var s = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive");
                foreach (ManagementObject d in s.Get())
                    disks.Add(new DiskInfo
                    {
                        Index = Convert.ToInt32(d["Index"]),
                        Model = d["Model"]?.ToString() ?? "Desconhecido",
                        SizeGB = Convert.ToDouble(d["Size"] ?? 0) / 1073741824.0
                    });
            }
            catch { }
            return disks.OrderBy(x => x.Index).ToList();
        }
        #endregion

        #region BACKUP DO BCD
        static void BackupBcdFlow()
        {
            string sysLetter = FindEfiSystemPartition() ?? "C:\\";
            var result = BackupBcd(sysLetter, silent: false);
            ShowLogResult(" RESULTADO - backup do BCD ", result);
        }

        static List<string> BackupBcd(string sysLetter, bool silent)
        {
            try
            {
                string backupDir = Path.Combine(sysLetter, "BCD_Backup");
                Directory.CreateDirectory(backupDir);
                string backupFile = Path.Combine(backupDir, $"BCD_{DateTime.Now:yyyyMMdd_HHmmss}.bak");
                var output = ExecuteCaptured("bcdedit.exe", $"/export \"{backupFile}\"");
                if (!silent) output.Insert(0, $"Backup salvo em: {backupFile}");
                return output;
            }
            catch (Exception ex)
            {
                var r = new List<string> { "Falha ao gerar backup do BCD: " + ex.Message };
                return r;
            }
        }
        #endregion

        #region EXECUÇÃO DE PROCESSOS
        static List<string> ExecuteCaptured(string exe, string args)
        {
            var result = new List<string>();
            IntPtr ptr = IntPtr.Zero;
            bool is64 = Environment.Is64BitOperatingSystem;
            if (is64) Wow64DisableWow64FsRedirection(ref ptr);
            try
            {
                var psi = new ProcessStartInfo(exe, args)
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                using var p = Process.Start(psi);
                if (p != null)
                {
                    while (!p.StandardOutput.EndOfStream) result.Add(p.StandardOutput.ReadLine());
                    string err = p.StandardError.ReadToEnd();
                    if (!string.IsNullOrWhiteSpace(err)) result.Add("ERRO: " + err.Trim());
                    p.WaitForExit();
                }
            }
            catch (Exception ex) { result.Add("ERRO: " + ex.Message); }
            finally { if (is64) Wow64RevertWow64FsRedirection(ptr); }
            return result;
        }

        static List<string> RunDiskpartCapture(string script)
        {
            var result = new List<string>();
            string scriptPath = Path.Combine(Path.GetTempPath(), "wtec_br.txt");
            File.WriteAllText(scriptPath, script);
            var output = ExecuteCaptured("diskpart.exe", $"/s \"{scriptPath}\"");
            try { File.Delete(scriptPath); } catch { }
            return output;
        }
        #endregion

        #region UI HELPERS
        static void DrawBox(int x, int y, int w, int h, string t)
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.SetCursorPosition(x, y); Console.Write("╔" + new string('═', w - 2) + "╗");
            for (int i = 1; i < h - 1; i++) { Console.SetCursorPosition(x, y + i); Console.Write("║" + new string(' ', w - 2) + "║"); }
            Console.SetCursorPosition(x, y + h - 1); Console.Write("╚" + new string('═', w - 2) + "╝");
            if (!string.IsNullOrEmpty(t)) { Console.SetCursorPosition(x + (w / 2) - (t.Length / 2), y); Console.Write(t); }
        }

        static void DrawFooter(string text)
        {
            Console.SetCursorPosition(0, 29);
            Console.BackgroundColor = ConsoleColor.Gray; Console.ForegroundColor = ConsoleColor.Black;
            Console.Write(text.PadRight(100));
        }

        static string Truncate(string s, int m) => string.IsNullOrEmpty(s) ? "" : (s.Length <= m ? s : s.Substring(0, m - 3) + "...");

        static bool ConfirmAction(string msg)
        {
            Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
            DrawBox(10, 11, 80, 7, " CONFIRMAR ");
            Console.SetCursorPosition(13, 13); Console.Write(Truncate(msg, 76));
            Console.SetCursorPosition(13, 15); Console.Write("[S] SIM   [N] NAO");
            return Console.ReadKey(true).Key == ConsoleKey.S;
        }

        static void ShowMessage(string msg, ConsoleColor color)
        {
            Console.BackgroundColor = color; Console.Clear();
            DrawBox(10, 12, 80, 6, " AVISO ");
            Console.ForegroundColor = ConsoleColor.White;
            Console.SetCursorPosition(13, 14); Console.Write(Truncate(msg, 76));
            DrawFooter(" Pressione qualquer tecla para voltar ");
            Console.ReadKey(true);
        }

        static void ShowLogResult(string title, List<string> lines)
        {
            Console.BackgroundColor = ConsoleColor.Black; Console.Clear();
            DrawBox(2, 2, 96, 26, title);
            int y = 4;
            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                Console.SetCursorPosition(4, y++);
                Console.ForegroundColor = line.ToLower().Contains("erro") || line.ToLower().Contains("error") ? ConsoleColor.Red : ConsoleColor.White;
                Console.Write(Truncate(line.Trim(), 90));
                if (y > 26) break;
            }
            DrawFooter(" Pressione qualquer tecla para continuar ");
            Console.ReadKey(true);
        }

        static void ShowSplash()
        {
            Console.BackgroundColor = ConsoleColor.Black; Console.Clear();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.SetCursorPosition(35, 12); Console.Write("Wtec BootRepair v1.0");
            System.Threading.Thread.Sleep(700);
        }
        #endregion
    }
}
