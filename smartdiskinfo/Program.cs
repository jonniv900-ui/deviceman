using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;

namespace WtecSmartHealthCheck
{
    class Program
    {
        struct DiskEntry
        {
            public int Index;
            public string Model;
            public string PnpDeviceId;
            public double SizeGB;
            public string InterfaceType;
        }

        class SmartAttribute
        {
            public byte Id;
            public byte Current;
            public byte Worst;
            public ulong Raw;
        }

        class SmartResult
        {
            public bool Supported;
            public bool PredictFailure;
            public string Verdict = "N/D";
            public ConsoleColor VerdictColor = ConsoleColor.Gray;
            public List<SmartAttribute> Attributes = new();
        }

        // Nomes amigáveis dos atributos S.M.A.R.T. mais relevantes (padrão da indústria;
        // alguns fabricantes usam IDs/significados ligeiramente diferentes).
        static readonly Dictionary<byte, string> KnownAttributes = new()
        {
            { 1, "Taxa de Erro de Leitura" },
            { 3, "Tempo de Spin-Up" },
            { 4, "Contagem de Start/Stop" },
            { 5, "Setores Realocados" },
            { 7, "Taxa de Erro de Busca" },
            { 9, "Horas Ligado (Power-On Hours)" },
            { 10, "Contagem de Retry de Spin" },
            { 12, "Contagem de Ciclos Liga/Desliga" },
            { 187, "Erros Não-Corrigíveis Reportados" },
            { 188, "Contagem de Timeout de Comando" },
            { 190, "Temperatura (Airflow)" },
            { 194, "Temperatura (°C, aproximado)" },
            { 196, "Eventos de Realocação" },
            { 197, "Setores Pendentes Atuais" },
            { 198, "Setores Incorrigíveis (Offline)" },
            { 199, "Erros CRC UDMA" },
            { 241, "Total de LBAs Escritos" },
            { 242, "Total de LBAs Lidos" },
        };

        // IDs considerados críticos para o veredito de saúde (indicam degradação física real).
        static readonly byte[] CriticalIds = { 5, 196, 197, 198 };

        [STAThread]
        static void Main()
        {
            try { Console.SetWindowSize(100, 32); } catch { }
            Console.Title = "Wtec SmartHealthCheck v1.0";
            Console.CursorVisible = false;
            ShowSplash();

            while (true)
            {
                var disks = ListDisks();
                int sel = ShowDiskList(disks);
                if (sel == -1) break;
                ShowDiskDetail(disks[sel]);
            }
        }

        #region LISTAGEM DE DISCOS
        static List<DiskEntry> ListDisks()
        {
            var disks = new List<DiskEntry>();
            try
            {
                using var s = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive");
                foreach (ManagementObject d in s.Get())
                {
                    disks.Add(new DiskEntry
                    {
                        Index = Convert.ToInt32(d["Index"]),
                        Model = d["Model"]?.ToString() ?? "Desconhecido",
                        PnpDeviceId = d["PNPDeviceID"]?.ToString() ?? "",
                        SizeGB = Convert.ToDouble(d["Size"] ?? 0) / 1073741824.0,
                        InterfaceType = d["InterfaceType"]?.ToString() ?? "N/A"
                    });
                }
            }
            catch { }
            return disks.OrderBy(x => x.Index).ToList();
        }

        static int ShowDiskList(List<DiskEntry> disks)
        {
            int sel = 0;
            var cache = new Dictionary<int, SmartResult>();
            while (true)
            {
                Console.BackgroundColor = ConsoleColor.DarkBlue; Console.Clear();
                DrawBox(3, 2, 94, 24, " VERIFICADOR DE SAUDE DO DISCO (S.M.A.R.T.) ");

                if (disks.Count == 0)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.SetCursorPosition(6, 6); Console.Write("Nenhum disco físico encontrado.");
                }

                for (int i = 0; i < disks.Count; i++)
                {
                    if (!cache.ContainsKey(i)) cache[i] = ReadSmart(disks[i].PnpDeviceId);
                    var r = cache[i];

                    Console.SetCursorPosition(6, 5 + i * 2);
                    if (i == sel) { Console.BackgroundColor = ConsoleColor.Cyan; Console.ForegroundColor = ConsoleColor.Black; }
                    else { Console.BackgroundColor = ConsoleColor.DarkBlue; Console.ForegroundColor = ConsoleColor.White; }
                    Console.Write($"> DISCO {disks[i].Index}: {Truncate(disks[i].Model, 45)} ({disks[i].SizeGB:F0} GB, {disks[i].InterfaceType})".PadRight(84));

                    Console.SetCursorPosition(8, 6 + i * 2);
                    Console.BackgroundColor = ConsoleColor.DarkBlue;
                    Console.ForegroundColor = r.VerdictColor;
                    Console.Write($"Status: {r.Verdict}".PadRight(30));
                }

                DrawFooter(" [SETAS] Navegar | [ENTER] Ver detalhes | [R] Atualizar | [ESC] Sair ");
                var k = Console.ReadKey(true).Key;
                if (k == ConsoleKey.Escape) return -1;
                if (k == ConsoleKey.R) { cache.Clear(); continue; }
                if (k == ConsoleKey.UpArrow && sel > 0) sel--;
                if (k == ConsoleKey.DownArrow && sel < disks.Count - 1) sel++;
                if (k == ConsoleKey.Enter && disks.Count > 0) return sel;
            }
        }
        #endregion

        #region DETALHE DO DISCO
        static void ShowDiskDetail(DiskEntry disk)
        {
            var r = ReadSmart(disk.PnpDeviceId);

            Console.BackgroundColor = ConsoleColor.Black; Console.Clear();
            DrawBox(2, 1, 96, 27, $" DISCO {disk.Index} - {Truncate(disk.Model, 50)} ");

            Console.ForegroundColor = r.VerdictColor;
            Console.SetCursorPosition(4, 3);
            Console.Write($"STATUS GERAL: {r.Verdict}" + (r.PredictFailure ? "  (driver reporta risco iminente de falha!)" : ""));

            if (!r.Supported)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.SetCursorPosition(4, 5);
                Console.Write("Este disco não expõe dados S.M.A.R.T. detalhados via WMI");
                Console.SetCursorPosition(4, 6);
                Console.Write("(comum em NVMe, discos USB e alguns controladores RAID).");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.SetCursorPosition(4, 5);
                Console.Write("ID   Atributo                              Atual  Pior  Valor Bruto");
                Console.SetCursorPosition(4, 6);
                Console.Write(new string('-', 88));

                int y = 7;
                foreach (var a in r.Attributes.OrderBy(a => a.Id))
                {
                    if (y > 25) break;
                    string name = KnownAttributes.TryGetValue(a.Id, out var n) ? n : $"Atributo #{a.Id}";
                    bool critical = CriticalIds.Contains(a.Id) && a.Raw > 0;
                    Console.ForegroundColor = critical ? ConsoleColor.Red : ConsoleColor.White;
                    Console.SetCursorPosition(4, y++);
                    Console.Write($"{a.Id,-4} {Truncate(name, 36),-36} {a.Current,5}  {a.Worst,4}  {a.Raw,10}");
                }
            }

            DrawFooter(" Pressione qualquer tecla para voltar ");
            Console.ReadKey(true);
        }
        #endregion

        #region LEITURA S.M.A.R.T. (WMI root\WMI)
        static SmartResult ReadSmart(string pnpDeviceId)
        {
            var result = new SmartResult();
            if (string.IsNullOrEmpty(pnpDeviceId)) return result;

            string baseName = pnpDeviceId.ToLower().Replace("\\", "#");

            try
            {
                using var statusSearcher = new ManagementObjectSearcher(@"root\WMI", "SELECT * FROM MSStorageDriver_FailurePredictStatus");
                foreach (ManagementObject o in statusSearcher.Get())
                {
                    string instance = o["InstanceName"]?.ToString()?.ToLower() ?? "";
                    if (!instance.StartsWith(baseName)) continue;
                    result.Supported = true;
                    result.PredictFailure = Convert.ToBoolean(o["PredictFailure"]);
                    break;
                }
            }
            catch { }

            try
            {
                using var dataSearcher = new ManagementObjectSearcher(@"root\WMI", "SELECT * FROM MSStorageDriver_FailurePredictData");
                foreach (ManagementObject o in dataSearcher.Get())
                {
                    string instance = o["InstanceName"]?.ToString()?.ToLower() ?? "";
                    if (!instance.StartsWith(baseName)) continue;

                    var bytes = o["VendorSpecific"] as byte[];
                    if (bytes == null) continue;
                    result.Supported = true;

                    // Estrutura padrão: 2 bytes reservados + até 30 entradas de 12 bytes cada.
                    for (int i = 0; i < 30; i++)
                    {
                        int offset = 2 + i * 12;
                        if (offset + 11 >= bytes.Length) break;
                        byte id = bytes[offset];
                        if (id == 0) continue;

                        ulong raw = 0;
                        for (int b = 0; b < 6; b++) raw |= (ulong)bytes[offset + 5 + b] << (8 * b);

                        result.Attributes.Add(new SmartAttribute
                        {
                            Id = id,
                            Current = bytes[offset + 3],
                            Worst = bytes[offset + 4],
                            Raw = raw
                        });
                    }
                    break;
                }
            }
            catch { }

            result.Verdict = ComputeVerdict(result, out var color);
            result.VerdictColor = color;
            return result;
        }

        static string ComputeVerdict(SmartResult r, out ConsoleColor color)
        {
            if (!r.Supported)
            {
                color = ConsoleColor.Gray;
                return "N/D (sem dados S.M.A.R.T.)";
            }

            ulong realocados = r.Attributes.FirstOrDefault(a => a.Id == 5)?.Raw ?? 0;
            ulong pendentes = r.Attributes.FirstOrDefault(a => a.Id == 197)?.Raw ?? 0;
            ulong incorrigiveis = r.Attributes.FirstOrDefault(a => a.Id == 198)?.Raw ?? 0;
            ulong eventosRealocacao = r.Attributes.FirstOrDefault(a => a.Id == 196)?.Raw ?? 0;

            if (r.PredictFailure || realocados > 20 || pendentes > 0 || incorrigiveis > 0)
            {
                color = ConsoleColor.Red;
                return "CRITICO - considere substituir o disco";
            }
            if (realocados > 0 || eventosRealocacao > 0)
            {
                color = ConsoleColor.Yellow;
                return "ATENCAO - sinais iniciais de desgaste";
            }
            color = ConsoleColor.Green;
            return "OK";
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

        static void ShowSplash()
        {
            Console.BackgroundColor = ConsoleColor.Black; Console.Clear();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.SetCursorPosition(30, 12); Console.Write("Wtec SmartHealthCheck v1.0");
            System.Threading.Thread.Sleep(700);
        }
        #endregion
    }
}
