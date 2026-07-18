using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace WtecLocalPasswordReset
{
    // NOTA: esta ferramenta assume o mesmo contexto do resto do kit (Wtec) - uso por
    // técnico com acesso físico e autorização para reparar a máquina. O método usado
    // (troca do Sticky Keys/Utilman por cmd.exe) é a técnica clássica e reversível de
    // reset de senha offline, e não escreve diretamente na estrutura binária do SAM -
    // isso evita o risco de corromper o banco de contas do Windows.
    class Program
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool Wow64DisableWow64FsRedirection(ref IntPtr ptr);
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool Wow64RevertWow64FsRedirection(IntPtr ptr);

        struct WindowsInstall { public string Drive; public string Version; }

        static readonly (string file, string desc)[] Targets =
        {
            ("sethc.exe", "Sticky Keys - pressione SHIFT 5x na tela de logon"),
            ("Utilman.exe", "Facilidade de Acesso - clique no ícone no canto inferior esquerdo da tela de logon")
        };

        [STAThread]
        static void Main()
        {
            try { Console.SetWindowSize(100, 32); } catch { }
            Console.Title = "Wtec LocalPasswordReset v1.0";
            Console.CursorVisible = false;
            ShowSplash();

            while (true)
            {
                int action = ShowMainMenu();
                if (action == -1) break;
                switch (action)
                {
                    case 0: ActivateBypassFlow(); break;
                    case 1: RestoreBypassFlow(); break;
                    case 2: ListAccountsFlow(); break;
                }
            }
        }

        #region MENU PRINCIPAL
        static int ShowMainMenu()
        {
            int sel = 0;
            string[] items =
            {
                " ATIVAR BYPASS DE EMERGENCIA (RESET DE SENHA) ",
                " REVERTER BYPASS (RESTAURAR ARQUIVO ORIGINAL) ",
                " VER CONTAS DE USUARIO LOCAIS (SOMENTE LEITURA) ",
                " SAIR "
            };
            while (true)
            {
                Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
                DrawBox(10, 3, 80, 22, " WTEC LOCAL PASSWORD RESET ");
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.SetCursorPosition(13, 6);
                Console.Write("Use apenas em máquinas que você tem autorização para reparar.");

                for (int i = 0; i < items.Length; i++)
                {
                    Console.SetCursorPosition(13, 10 + i);
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

        #region ATIVAR / REVERTER BYPASS
        static void ActivateBypassFlow()
        {
            var installs = ScanWindowsInstallations();
            if (installs.Count == 0) { ShowMessage("Nenhuma instalação do Windows encontrada nos discos.", ConsoleColor.DarkRed); return; }
            var target = installs.Count == 1 ? installs[0] : SelectInstall(installs, "SELECIONE A INSTALACAO DO WINDOWS");
            if (string.IsNullOrEmpty(target.Drive)) return;

            int methodIdx = SelectMethod(target.Drive);
            if (methodIdx == -1) return;
            var (file, desc) = Targets[methodIdx];

            string sys32 = Path.Combine(target.Drive, "Windows", "System32");
            string targetPath = Path.Combine(sys32, file);
            string backupPath = targetPath + ".wtecbak";

            if (!File.Exists(targetPath)) { ShowMessage($"Arquivo {file} não encontrado em {sys32}.", ConsoleColor.DarkRed); return; }

            if (!ConfirmAction($"Ativar bypass via {file}? Um cmd.exe (SYSTEM) vai abrir na tela de logon: {desc}")) return;

            var log = new List<string>();
            try
            {
                if (!File.Exists(backupPath))
                {
                    TakeOwnershipAndGrant(targetPath, log);
                    File.Copy(targetPath, backupPath, overwrite: false);
                    log.Add($"Backup criado: {backupPath}");
                }
                else
                {
                    log.Add("Backup já existia (bypass provavelmente já estava ativo) - mantido.");
                }

                string cmdExe = Path.Combine(sys32, "cmd.exe");
                File.Copy(cmdExe, targetPath, overwrite: true);
                log.Add($"{file} substituído por cmd.exe com sucesso.");
                log.Add("");
                log.Add("PRÓXIMOS PASSOS:");
                log.Add("1. Reinicie e dê boot no disco (remova o pendrive/PE).");
                log.Add($"2. Na tela de logon: {desc}");
                log.Add("3. No prompt (SYSTEM) que abrir, rode:  net user NOMEUSUARIO NOVASENHA");
                log.Add("4. Faça login normalmente com a nova senha.");
                log.Add("5. IMPORTANTE: rode este utilitário de novo e escolha 'REVERTER BYPASS'");
                log.Add("   pra restaurar o arquivo original - deixar o cmd.exe ativo na tela");
                log.Add("   de logon é uma porta aberta de segurança.");
            }
            catch (Exception ex) { log.Add("ERRO: " + ex.Message); }

            ShowLogResult(" RESULTADO - ATIVAR BYPASS ", log);
        }

        static void RestoreBypassFlow()
        {
            var installs = ScanWindowsInstallations();
            if (installs.Count == 0) { ShowMessage("Nenhuma instalação do Windows encontrada nos discos.", ConsoleColor.DarkRed); return; }
            var target = installs.Count == 1 ? installs[0] : SelectInstall(installs, "SELECIONE A INSTALACAO DO WINDOWS");
            if (string.IsNullOrEmpty(target.Drive)) return;

            string sys32 = Path.Combine(target.Drive, "Windows", "System32");
            var pending = new List<(string file, string backup)>();
            foreach (var (file, _) in Targets)
            {
                string backup = Path.Combine(sys32, file + ".wtecbak");
                if (File.Exists(backup)) pending.Add((file, backup));
            }

            if (pending.Count == 0) { ShowMessage("Nenhum bypass ativo encontrado (nenhum backup .wtecbak presente).", ConsoleColor.DarkBlue); return; }

            if (!ConfirmAction($"Restaurar {pending.Count} arquivo(s) original(is): {string.Join(", ", pending.ConvertAll(p => p.file))}?")) return;

            var log = new List<string>();
            foreach (var (file, backup) in pending)
            {
                try
                {
                    string targetPath = Path.Combine(sys32, file);
                    TakeOwnershipAndGrant(targetPath, log);
                    File.Copy(backup, targetPath, overwrite: true);
                    File.Delete(backup);
                    log.Add($"{file} restaurado com sucesso.");
                }
                catch (Exception ex) { log.Add($"ERRO ao restaurar {file}: " + ex.Message); }
            }
            ShowLogResult(" RESULTADO - REVERTER BYPASS ", log);
        }

        static int SelectMethod(string winDrive)
        {
            int sel = 0;
            while (true)
            {
                Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
                DrawBox(10, 8, 80, 12, " ESCOLHA O METODO ");
                for (int i = 0; i < Targets.Length; i++)
                {
                    string backup = Path.Combine(winDrive, "Windows", "System32", Targets[i].file + ".wtecbak");
                    string status = File.Exists(backup) ? " [JA ATIVO]" : "";
                    Console.SetCursorPosition(13, 11 + i * 2);
                    if (i == sel) { Console.BackgroundColor = ConsoleColor.Cyan; Console.ForegroundColor = ConsoleColor.Black; }
                    else { Console.BackgroundColor = ConsoleColor.DarkBlue; Console.ForegroundColor = ConsoleColor.White; }
                    Console.Write($"> {Targets[i].file}{status}".PadRight(70));
                    Console.SetCursorPosition(15, 12 + i * 2);
                    Console.BackgroundColor = ConsoleColor.DarkBlue; Console.ForegroundColor = ConsoleColor.Gray;
                    Console.Write(Truncate(Targets[i].desc, 65).PadRight(68));
                }
                DrawFooter(" [SETAS] Navegar | [ENTER] Selecionar | [ESC] Cancelar ");
                var k = Console.ReadKey(true).Key;
                if (k == ConsoleKey.Escape) return -1;
                if (k == ConsoleKey.UpArrow && sel > 0) sel--;
                if (k == ConsoleKey.DownArrow && sel < Targets.Length - 1) sel++;
                if (k == ConsoleKey.Enter) return sel;
            }
        }
        #endregion

        #region VER CONTAS (SOMENTE LEITURA)
        static void ListAccountsFlow()
        {
            var installs = ScanWindowsInstallations();
            if (installs.Count == 0) { ShowMessage("Nenhuma instalação do Windows encontrada nos discos.", ConsoleColor.DarkRed); return; }
            var target = installs.Count == 1 ? installs[0] : SelectInstall(installs, "SELECIONE A INSTALACAO DO WINDOWS");
            if (string.IsNullOrEmpty(target.Drive)) return;

            var log = new List<string>();
            string hiveName = "WtecOfflineSAM";
            string samPath = Path.Combine(target.Drive, "Windows", "System32", "config", "SAM");

            if (!File.Exists(samPath)) { ShowMessage("Arquivo SAM não encontrado nessa instalação.", ConsoleColor.DarkRed); return; }

            try
            {
                ExecuteCaptured("reg.exe", $"load HKLM\\{hiveName} \"{samPath}\"");
                var output = ExecuteCaptured("reg.exe", $"query HKLM\\{hiveName}\\SAM\\Domains\\Account\\Users\\Names");
                foreach (var line in output)
                {
                    string trimmed = line.Trim();
                    if (trimmed.StartsWith("HKEY_LOCAL_MACHINE"))
                    {
                        int idx = trimmed.LastIndexOf('\\');
                        if (idx >= 0 && idx < trimmed.Length - 1)
                            log.Add("- " + trimmed.Substring(idx + 1));
                    }
                }
                if (log.Count == 0) log.Add("Nenhuma conta encontrada (ou formato de retorno inesperado).");
            }
            catch (Exception ex) { log.Add("ERRO: " + ex.Message); }
            finally { ExecuteCaptured("reg.exe", $"unload HKLM\\{hiveName}"); }

            ShowLogResult($" CONTAS LOCAIS - {target.Drive} ", log);
        }
        #endregion

        #region PERMISSOES / ARQUIVOS
        /// <summary>Toma posse do arquivo (normalmente pertence ao TrustedInstaller) e concede
        /// controle total ao grupo Administradores via SID (S-1-5-32-544), evitando depender
        /// do nome do grupo, que muda conforme o idioma do Windows (ex: "Administradores" em pt-BR).</summary>
        static void TakeOwnershipAndGrant(string path, List<string> log)
        {
            var o1 = ExecuteCaptured("takeown.exe", $"/f \"{path}\"");
            log.AddRange(o1);
            var o2 = ExecuteCaptured("icacls.exe", $"\"{path}\" /grant *S-1-5-32-544:F");
            log.AddRange(o2);
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
                        results.Add(new WindowsInstall { Drive = d.Name, Version = ver });
                    }
                }
                catch { }
            }
            return results;
        }

        static WindowsInstall SelectInstall(List<WindowsInstall> installs, string title)
        {
            int sel = 0;
            while (true)
            {
                Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
                DrawBox(15, 8, 70, 14, $" {title} ");
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
                Console.SetCursorPosition(4, y++);
                Console.ForegroundColor = line.ToLower().Contains("erro") ? ConsoleColor.Red : ConsoleColor.White;
                Console.Write(Truncate(line, 90));
                if (y > 26) break;
            }
            DrawFooter(" Pressione qualquer tecla para continuar ");
            Console.ReadKey(true);
        }

        static void ShowSplash()
        {
            Console.BackgroundColor = ConsoleColor.Black; Console.Clear();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.SetCursorPosition(28, 12); Console.Write("Wtec LocalPasswordReset v1.0");
            System.Threading.Thread.Sleep(700);
        }
        #endregion
    }
}
