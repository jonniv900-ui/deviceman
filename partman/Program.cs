using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;

namespace WtecPartitionManager
{
    class Program
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool Wow64DisableWow64FsRedirection(ref IntPtr ptr);
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool Wow64RevertWow64FsRedirection(IntPtr ptr);

        struct DiskInfo { public int Index; public string Model; public double SizeGB; public string Type; }

        [STAThread]
        static void Main()
        {
            try { Console.SetWindowSize(100, 32); } catch { }
            Console.Title = "Wtec PartitionManager v1.0";
            Console.CursorVisible = false;
            ShowSplash();

            while (true)
            {
                int diskIndex = DiskSelector();
                if (diskIndex == -1) break;
                DiskActionsMenu(diskIndex);
            }
        }

        #region TELA PRINCIPAL / SELEÇÃO DE DISCO
        static int DiskSelector()
        {
            int sel = 0;
            while (true)
            {
                var disks = ListDisks();
                Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
                DrawBox(5, 3, 90, 22, " GERENCIADOR DE PARTICOES - SELECIONE UM DISCO ");

                if (disks.Count == 0)
                {
                    Console.SetCursorPosition(8, 8); Console.ForegroundColor = ConsoleColor.Red;
                    Console.Write("Nenhum disco físico encontrado.");
                }
                for (int i = 0; i < disks.Count; i++)
                {
                    Console.SetCursorPosition(8, 6 + i);
                    if (i == sel) { Console.BackgroundColor = ConsoleColor.Cyan; Console.ForegroundColor = ConsoleColor.Black; }
                    else { Console.BackgroundColor = ConsoleColor.DarkBlue; Console.ForegroundColor = ConsoleColor.White; }
                    Console.Write($"> DISCO {disks[i].Index}: {Truncate(disks[i].Model, 55)} ({disks[i].SizeGB:F1} GB) [{disks[i].Type}]".PadRight(80));
                }

                DrawFooter(" [SETAS] Navegar | [ENTER] Selecionar disco | [R] Atualizar | [ESC] Sair ");
                var k = Console.ReadKey(true).Key;
                if (k == ConsoleKey.Escape) return -1;
                if (k == ConsoleKey.R) continue;
                if (k == ConsoleKey.UpArrow && sel > 0) sel--;
                if (k == ConsoleKey.DownArrow && sel < disks.Count - 1) sel++;
                if (k == ConsoleKey.Enter && disks.Count > 0) return disks[sel].Index;
            }
        }

        static List<DiskInfo> ListDisks()
        {
            var disks = new List<DiskInfo>();
            try
            {
                using var s = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive");
                foreach (ManagementObject d in s.Get())
                {
                    disks.Add(new DiskInfo
                    {
                        Index = Convert.ToInt32(d["Index"]),
                        Model = d["Model"]?.ToString() ?? "Desconhecido",
                        SizeGB = Convert.ToDouble(d["Size"] ?? 0) / 1073741824.0,
                        Type = d["MediaType"]?.ToString()?.Contains("Fixed") == true ? "Fixo" : "Removível"
                    });
                }
            }
            catch { }
            return disks.OrderBy(x => x.Index).ToList();
        }
        #endregion

        #region MENU DE AÇÕES DO DISCO
        static void DiskActionsMenu(int diskIndex)
        {
            int sel = 0;
            string[] items =
            {
                " VER PARTICOES ",
                " CRIAR PARTICAO ",
                " FORMATAR PARTICAO ",
                " ATRIBUIR / TROCAR LETRA ",
                " ESTENDER PARTICAO (USAR ESPACO LIVRE) ",
                " EXCLUIR PARTICAO ",
                " LIMPAR DISCO INTEIRO (CLEAN) ",
                " CONVERTER ESTILO (MBR <-> GPT) ",
                " VOLTAR "
            };

            while (true)
            {
                Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
                DrawBox(20, 4, 60, 20, $" DISCO {diskIndex} - ACOES ");
                for (int i = 0; i < items.Length; i++)
                {
                    Console.SetCursorPosition(23, 7 + i);
                    if (i == sel) { Console.BackgroundColor = ConsoleColor.Cyan; Console.ForegroundColor = ConsoleColor.Black; }
                    else { Console.BackgroundColor = ConsoleColor.DarkBlue; Console.ForegroundColor = ConsoleColor.White; }
                    Console.Write(items[i].PadRight(54));
                }
                DrawFooter(" [SETAS] Navegar | [ENTER] Executar | [ESC] Voltar ");
                var k = Console.ReadKey(true).Key;
                if (k == ConsoleKey.Escape) return;
                if (k == ConsoleKey.UpArrow && sel > 0) sel--;
                if (k == ConsoleKey.DownArrow && sel < items.Length - 1) sel++;
                if (k != ConsoleKey.Enter) continue;

                switch (sel)
                {
                    case 0: ShowPartitions(diskIndex); break;
                    case 1: CreatePartitionFlow(diskIndex); break;
                    case 2: FormatPartitionFlow(diskIndex); break;
                    case 3: AssignLetterFlow(diskIndex); break;
                    case 4: ExtendPartitionFlow(diskIndex); break;
                    case 5: DeletePartitionFlow(diskIndex); break;
                    case 6: CleanDiskFlow(diskIndex); break;
                    case 7: ConvertStyleFlow(diskIndex); break;
                    case 8: return;
                }
            }
        }

        static void ShowPartitions(int diskIndex)
        {
            var output = RunDiskpartCapture($"select disk {diskIndex}\r\nlist partition\r\n");
            Console.BackgroundColor = ConsoleColor.Black; Console.Clear();
            DrawBox(2, 2, 96, 26, $" DISCO {diskIndex} - PARTICOES ");
            int y = 4;
            foreach (var line in output)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                Console.SetCursorPosition(4, y++);
                Console.ForegroundColor = ConsoleColor.White;
                Console.Write(Truncate(line.Trim(), 90));
                if (y > 26) break;
            }
            DrawFooter(" Pressione qualquer tecla para voltar ");
            Console.ReadKey(true);
        }
        #endregion

        #region FLUXOS DE AÇÃO
        static void CreatePartitionFlow(int diskIndex)
        {
            string size = PromptInput("Tamanho em MB (deixe vazio para usar todo espaço livre restante):", 60, 12);
            string fs = PromptChoice("Sistema de arquivos:", new[] { "NTFS", "FAT32", "EXFAT" }, 60, 14);
            string label = PromptInput("Rótulo do volume (opcional):", 60, 16);

            var script = new System.Text.StringBuilder();
            script.AppendLine($"select disk {diskIndex}");
            script.AppendLine(string.IsNullOrWhiteSpace(size) ? "create partition primary" : $"create partition primary size={size}");
            script.AppendLine($"format quick fs={fs.ToLower()}" + (string.IsNullOrWhiteSpace(label) ? "" : $" label=\"{label}\""));
            script.AppendLine("assign");

            if (!ConfirmAction($"Criar partição {(string.IsNullOrWhiteSpace(size) ? "(todo espaço livre)" : size + "MB")} em {fs}?")) return;
            RunAndShowResult(script.ToString(), "CRIANDO PARTICAO...");
        }

        static void FormatPartitionFlow(int diskIndex)
        {
            string part = PromptInput("Número da partição a formatar (veja em 'VER PARTICOES'):", 60, 12);
            if (string.IsNullOrWhiteSpace(part)) return;
            string fs = PromptChoice("Novo sistema de arquivos:", new[] { "NTFS", "FAT32", "EXFAT" }, 60, 14);
            string label = PromptInput("Novo rótulo (opcional):", 60, 16);

            if (!ConfirmDestructive($"FORMATAR a partição {part} do disco {diskIndex}? TODOS OS DADOS SERAO PERDIDOS.")) return;

            var script = $"select disk {diskIndex}\r\nselect partition {part}\r\nformat quick fs={fs.ToLower()}" +
                         (string.IsNullOrWhiteSpace(label) ? "" : $" label=\"{label}\"") + " override\r\n";
            RunAndShowResult(script, "FORMATANDO PARTICAO...");
        }

        static void AssignLetterFlow(int diskIndex)
        {
            string part = PromptInput("Número da partição:", 60, 12);
            if (string.IsNullOrWhiteSpace(part)) return;
            string letter = PromptInput("Nova letra (ex: E) - deixe vazio para remover a letra atual:", 60, 14);

            string script;
            if (string.IsNullOrWhiteSpace(letter))
                script = $"select disk {diskIndex}\r\nselect partition {part}\r\nremove\r\n";
            else
                script = $"select disk {diskIndex}\r\nselect partition {part}\r\nassign letter={letter.Trim().TrimEnd(':')}\r\n";

            if (!ConfirmAction(string.IsNullOrWhiteSpace(letter) ? $"Remover letra da partição {part}?" : $"Atribuir letra {letter}: à partição {part}?")) return;
            RunAndShowResult(script, "ALTERANDO LETRA...");
        }

        static void ExtendPartitionFlow(int diskIndex)
        {
            string part = PromptInput("Número da partição a estender (usa o espaço livre contíguo):", 60, 12);
            if (string.IsNullOrWhiteSpace(part)) return;
            if (!ConfirmAction($"Estender a partição {part} usando todo o espaço livre disponível?")) return;
            var script = $"select disk {diskIndex}\r\nselect partition {part}\r\nextend\r\n";
            RunAndShowResult(script, "ESTENDENDO PARTICAO...");
        }

        static void DeletePartitionFlow(int diskIndex)
        {
            string part = PromptInput("Número da partição a EXCLUIR:", 60, 12);
            if (string.IsNullOrWhiteSpace(part)) return;
            if (!ConfirmDestructive($"EXCLUIR permanentemente a partição {part} do disco {diskIndex}? TODOS OS DADOS SERAO PERDIDOS.")) return;
            var script = $"select disk {diskIndex}\r\nselect partition {part}\r\ndelete partition override\r\n";
            RunAndShowResult(script, "EXCLUINDO PARTICAO...");
        }

        static void CleanDiskFlow(int diskIndex)
        {
            if (!ConfirmDestructive($"APAGAR TODAS AS PARTICOES do disco {diskIndex}? Esta ação é IRREVERSÍVEL.")) return;
            var script = $"select disk {diskIndex}\r\nclean\r\n";
            RunAndShowResult(script, "LIMPANDO DISCO...");
        }

        static void ConvertStyleFlow(int diskIndex)
        {
            string style = PromptChoice("Converter disco (LIMPO, sem partições) para:", new[] { "GPT", "MBR" }, 60, 12);
            if (!ConfirmDestructive($"O disco {diskIndex} precisa estar limpo (sem partições) para converter para {style}. " +
                                    "Isso vai LIMPAR o disco e então converter. Continuar?")) return;
            var script = $"select disk {diskIndex}\r\nclean\r\nconvert {style.ToLower()}\r\n";
            RunAndShowResult(script, $"CONVERTENDO PARA {style}...");
        }
        #endregion

        #region EXECUÇÃO / DISKPART / IO
        static void RunAndShowResult(string diskpartScript, string title)
        {
            Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
            DrawBox(20, 12, 60, 6, $" {title} ");
            Console.SetCursorPosition(23, 15); Console.Write("Executando, aguarde...");
            var output = RunDiskpartCapture(diskpartScript);

            Console.Clear();
            DrawBox(2, 2, 96, 26, " RESULTADO ");
            int y = 4;
            foreach (var line in output)
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

        static List<string> RunDiskpartCapture(string script)
        {
            var result = new List<string>();
            string scriptPath = Path.Combine(Path.GetTempPath(), "wtec_pm.txt");
            File.WriteAllText(scriptPath, script);

            IntPtr ptr = IntPtr.Zero;
            bool is64 = Environment.Is64BitOperatingSystem;
            if (is64) Wow64DisableWow64FsRedirection(ref ptr);
            try
            {
                var psi = new ProcessStartInfo("diskpart.exe", $"/s \"{scriptPath}\"")
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true
                };
                using var p = Process.Start(psi);
                if (p != null)
                {
                    while (!p.StandardOutput.EndOfStream)
                        result.Add(p.StandardOutput.ReadLine());
                    p.WaitForExit();
                }
            }
            catch (Exception ex) { result.Add("ERRO: " + ex.Message); }
            finally { if (is64) Wow64RevertWow64FsRedirection(ptr); try { File.Delete(scriptPath); } catch { } }

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
            DrawBox(15, 12, 70, 6, " CONFIRMAR ");
            Console.SetCursorPosition(18, 14); Console.Write(Truncate(msg, 66));
            Console.SetCursorPosition(18, 16); Console.Write("[S] SIM   [N] NAO");
            return Console.ReadKey(true).Key == ConsoleKey.S;
        }

        /// <summary>Confirmação reforçada para operações destrutivas (perda de dados):
        /// exige digitar a palavra CONFIRMAR, não basta apertar uma tecla.</summary>
        static bool ConfirmDestructive(string msg)
        {
            Console.BackgroundColor = ConsoleColor.DarkRed; Console.Clear();
            DrawBox(10, 10, 80, 9, " ATENCAO - OPERACAO DESTRUTIVA ");
            Console.ForegroundColor = ConsoleColor.White;
            Console.SetCursorPosition(13, 13); Console.Write(Truncate(msg, 76));
            Console.SetCursorPosition(13, 15); Console.Write("Digite CONFIRMAR (maiúsculas) e pressione Enter para prosseguir:");
            Console.SetCursorPosition(13, 16); Console.CursorVisible = true;
            string typed = Console.ReadLine() ?? "";
            Console.CursorVisible = false;
            return typed.Trim() == "CONFIRMAR";
        }

        static string PromptInput(string label, int x, int y)
        {
            Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
            DrawBox(x - 10, y - 4, 80, 12, " ENTRADA ");
            Console.SetCursorPosition(x - 7, y); Console.Write(label);
            Console.SetCursorPosition(x - 7, y + 2); Console.CursorVisible = true;
            string r = Console.ReadLine() ?? "";
            Console.CursorVisible = false;
            return r;
        }

        static string PromptChoice(string label, string[] options, int x, int y)
        {
            int sel = 0;
            while (true)
            {
                Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
                DrawBox(x - 10, y - 4, 80, 14, " ESCOLHA ");
                Console.SetCursorPosition(x - 7, y); Console.Write(label);
                for (int i = 0; i < options.Length; i++)
                {
                    Console.SetCursorPosition(x - 7, y + 2 + i);
                    if (i == sel) { Console.BackgroundColor = ConsoleColor.Cyan; Console.ForegroundColor = ConsoleColor.Black; }
                    else { Console.BackgroundColor = ConsoleColor.DarkBlue; Console.ForegroundColor = ConsoleColor.White; }
                    Console.Write(("> " + options[i]).PadRight(30));
                }
                var k = Console.ReadKey(true).Key;
                if (k == ConsoleKey.UpArrow && sel > 0) sel--;
                if (k == ConsoleKey.DownArrow && sel < options.Length - 1) sel++;
                if (k == ConsoleKey.Enter) return options[sel];
            }
        }

        static void ShowSplash()
        {
            Console.BackgroundColor = ConsoleColor.Black; Console.Clear();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.SetCursorPosition(32, 12); Console.Write("Wtec PartitionManager v1.0");
            System.Threading.Thread.Sleep(700);
        }
        #endregion
    }
}
