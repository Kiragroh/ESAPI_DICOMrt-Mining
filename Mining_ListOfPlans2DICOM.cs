using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;
using OfficeOpenXml;
using VMS.TPS.Common.Model.API;
using EsapiApplication = VMS.TPS.Common.Model.API.Application;

[assembly: AssemblyVersion("1.0.0.3")]
[assembly: AssemblyFileVersion("1.0.0.3")]
[assembly: AssemblyInformationalVersion("ExcelPatientDicomExportMG 1.0")]

namespace ExcelPatientDicomExportMG
{
    // ================================================================================
    //  Excel Patient DICOM Export
    //
    //  Input:
    //    Excel-Datei, erstes Worksheet.
    //    Spalte 1 / A enthält Patient-IDs.
    //    Zeile 1 = Header, Patient-IDs ab Zeile 2.
    //
    //  Export-Schalter in settings.ini:
    //    EXPORT_CBCT                  = TRUE/FALSE
    //    EXPORT_ACQUIRED_DOSE_RTIMAGE = TRUE/FALSE
    //    EXPORT_PLAN_SETS             = TRUE/FALSE
    //
    //  Plan-Set-Export:
    //    RTPLAN, RTDOSE, RTSTRUCT, CT, je nach Unterschalter.
    //
    //  Zielstruktur:
    //    EXPORT_BASE\<PatientId>\PLAN_SET\<CourseId>_<PlanId>\CT\*.dcm
    //    EXPORT_BASE\<PatientId>\PLAN_SET\<CourseId>_<PlanId>\RTSTRUCT\*.dcm
    //    EXPORT_BASE\<PatientId>\PLAN_SET\<CourseId>_<PlanId>\RTPLAN\*.dcm
    //    EXPORT_BASE\<PatientId>\PLAN_SET\<CourseId>_<PlanId>\RTDOSE\*.dcm
    //    EXPORT_BASE\<PatientId>\CBCT\<YYYYMMDD_HHMM>_<StructureSetId>\*.dcm
    //    EXPORT_BASE\<PatientId>\RTIMAGE\<YYYYMMDD>\*.dcm
    // ================================================================================
    internal class Program
    {
        private static Settings _config;

        [STAThread]
        private static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (string.IsNullOrWhiteSpace(exeDir))
                exeDir = AppDomain.CurrentDomain.BaseDirectory;

            string settingsPath = Path.Combine(exeDir, "settings.ini");
            _config = Settings.Load(settingsPath);

            using (ExportOptionsForm form = new ExportOptionsForm(_config))
            {
                DialogResult result = form.ShowDialog();
                if (result != DialogResult.OK)
                    return;

                _config = form.ResultSettings;
            }

            string excelFilePath = GetExcelPath(args);
            if (string.IsNullOrWhiteSpace(excelFilePath))
            {
                Console.WriteLine("Keine Excel-Datei ausgewählt.");
                PauseIfConfigured();
                return;
            }

            _config.LastExcelFile = excelFilePath;
            try
            {
                string lastDir = Path.GetDirectoryName(excelFilePath);
                if (!string.IsNullOrWhiteSpace(lastDir))
                    _config.ExcelInitialDirectory = lastDir;
            }
            catch { }

            try { _config.Save(settingsPath); }
            catch (Exception ex) { Console.WriteLine("WARN: settings.ini konnte nicht gespeichert werden: " + ex.Message); }

            Directory.CreateDirectory(_config.ExportBase);
            SetupLogging(_config.ExportBase);

            Console.WriteLine("╔════════════════════════════════════════════════════╗");
            Console.WriteLine("║   Excel Patient DICOM Export                       ║");
            Console.WriteLine("╚════════════════════════════════════════════════════╝");
            Console.WriteLine($"Start: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            Console.WriteLine($"Settings: {settingsPath}");
            Console.WriteLine($"Excel: {excelFilePath}");
            Console.WriteLine($"Excel Patient-ID: Spalte {_config.ExcelIdColumn} ({ColumnNumberToName(_config.ExcelIdColumn)}), ab Zeile {_config.ExcelStartRow}");
            Console.WriteLine($"EXPORT_BASE = {_config.ExportBase}");
            Console.WriteLine($"EXPORT_CBCT = {_config.ExportCbct}");
            Console.WriteLine($"EXPORT_ACQUIRED_DOSE_RTIMAGE = {_config.ExportAcquiredDoseRtImage}");
            Console.WriteLine($"EXPORT_PLAN_SETS = {_config.ExportPlanSets}");
            Console.WriteLine($"ONLY_TREATED_PLANS = {_config.OnlyTreatedPlans}");
            Console.WriteLine($"DEBUG_LIMIT_ENABLED = {_config.DebugLimitEnabled}; DEBUG_MAX_PATIENTS = {_config.DebugMaxPatients}");
            Console.WriteLine();

            List<string> patientIds = ReadPatientIdsFromExcel(excelFilePath, _config.ExcelIdColumn, _config.ExcelStartRow)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            Console.WriteLine($"Patient-IDs gefunden: {patientIds.Count}");
            Console.WriteLine();

            if (patientIds.Count == 0)
            {
                Console.WriteLine($"Keine Patient-IDs in Spalte {ColumnNumberToName(_config.ExcelIdColumn)} ab Zeile {_config.ExcelStartRow} gefunden.");
                PauseIfConfigured();
                return;
            }

            int processed = 0;
            int maxPatientsThisRun = _config.DebugLimitEnabled ? _config.DebugMaxPatients : int.MaxValue;

            try
            {
                using (EsapiApplication app = EsapiApplication.CreateApplication())
                {
                    Console.WriteLine("ESAPI verbunden.");
                    Console.WriteLine();

                    foreach (string patientId in patientIds)
                    {
                        if (processed >= maxPatientsThisRun)
                        {
                            Console.WriteLine($"Debug-Abbruch nach {processed} Patient(en).");
                            break;
                        }

                        Patient patient = null;
                        try
                        {
                            Console.WriteLine("────────────────────────────────────────────────────");
                            Console.WriteLine($"Patient: {patientId}");

                            patient = app.OpenPatientById(patientId);
                            if (patient == null)
                            {
                                Console.WriteLine($"WARN: Patient '{patientId}' nicht in ARIA gefunden.");
                                continue;
                            }

                            string patientRoot = Path.Combine(_config.ExportBase, MakeSafe(patient.Id));
                            Directory.CreateDirectory(patientRoot);

                            int movedTotal = 0;

                            // UIDs die ExportPlanSets bereits exportiert hat → ExportAcquiredDoseRtImages überspringt sie
                            var exportedImageUids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                            if (_config.ExportPlanSets)
                                movedTotal += ExportPlanSets(patient, patientRoot, exportedImageUids);

                            if (_config.ExportCbct)
                                movedTotal += ExportCbcts(patient, patientRoot);

                            if (_config.ExportAcquiredDoseRtImage)
                                movedTotal += ExportAcquiredDoseRtImages(patient, patientRoot, exportedImageUids);

                            Console.WriteLine($"Patient abgeschlossen: {patient.Id}, exportierte Dateien: {movedTotal}");
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine($"ERROR Patient {patientId}: {ex}");
                        }
                        finally
                        {
                            try
                            {
                                if (patient != null)
                                    app.ClosePatient();
                            }
                            catch { }

                            processed++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("FATAL:");
                Console.Error.WriteLine(ex);
            }

            Console.WriteLine();
            Console.WriteLine($"Fertig: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            PauseIfConfigured();
        }

        private static string GetExcelPath(string[] args)
        {
            if (args != null && args.Length > 0 && File.Exists(args[0]))
                return args[0];

            if (!string.IsNullOrWhiteSpace(_config.LastExcelFile) && File.Exists(_config.LastExcelFile))
                return _config.LastExcelFile;

            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Excel-Datei mit Patient-IDs auswählen";
                ofd.InitialDirectory = GetExistingDirectoryOrDefault(_config.ExcelInitialDirectory);
                ofd.Filter = "Excel-Dateien (*.xlsx)|*.xlsx|Alle Dateien (*.*)|*.*";
                ofd.CheckFileExists = true;
                ofd.CheckPathExists = true;
                return ofd.ShowDialog() == DialogResult.OK ? ofd.FileName : null;
            }
        }

        private static List<string> ReadPatientIdsFromExcel(string excelFilePath, int idColumn, int startRow)
        {
            var result = new List<string>();

            if (idColumn < 1) idColumn = 1;
            if (startRow < 1) startRow = 1;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var ws = package.Workbook.Worksheets.FirstOrDefault();
                if (ws == null || ws.Dimension == null)
                    return result;

                for (int row = startRow; row <= ws.Dimension.Rows; row++)
                {
                    string patientId = ws.Cells[row, idColumn].Text.Trim();
                    if (!string.IsNullOrWhiteSpace(patientId))
                        result.Add(patientId);
                }
            }

            return result;
        }

        private static string GetExistingDirectoryOrDefault(string preferred)
        {
            if (!string.IsNullOrWhiteSpace(preferred) && Directory.Exists(preferred))
                return preferred;

            if (!string.IsNullOrWhiteSpace(_config.LastExcelFile))
            {
                try
                {
                    string d = Path.GetDirectoryName(_config.LastExcelFile);
                    if (!string.IsNullOrWhiteSpace(d) && Directory.Exists(d))
                        return d;
                }
                catch { }
            }

            return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        }

        private static string ColumnNumberToName(int columnNumber)
        {
            if (columnNumber < 1)
                columnNumber = 1;

            string name = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                name = Convert.ToChar('A' + modulo) + name;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return name;
        }

        // ─── Plan-Sets: RTPLAN, RTDOSE, RTSTRUCT, CT + RTIMAGE pro Plan ──────────────────────
        private static int ExportPlanSets(Patient patient, string patientRoot,
                                           HashSet<string> exportedImageUids)
        {
            int movedTotal = 0;
            int planCount  = 0;

            // Alle ACQUIRED_DOSE-Bilder einmal patientenweit einlesen;
            // pro Plan werden dann die Bilder per Beam-ID zugeordnet.
            List<RtImageEntry> allRtImages = _config.ExportAcquiredDoseRtImage
                ? CollectAcquiredDoseImages(patient)
                : new List<RtImageEntry>();

            if (_config.ExportAcquiredDoseRtImage)
                Console.WriteLine($"  RTIMAGE-Vorabscan: {allRtImages.Count} ACQUIRED_DOSE-Bild(er) patientenweit.");

            foreach (var course in patient.Courses.OrderBy(c => c.Id))
            {
                // ONLY_TREATED_PLANS=TRUE → plan.IsTreated direkt nutzen (kein Reflection-Hack)
                foreach (var plan in course.PlanSetups
                    .OrderBy(p => p.Id)
                    .Where(p => !_config.OnlyTreatedPlans || p.IsTreated))
                {
                    planCount++;

                    // Ordner direkt im Patienten-Root – wie AutoExportMG (kein PLAN_SET-Elternordner)
                    string planFolderName = MakeSafe($"{course.Id}_{plan.Id}");
                    string planRoot       = Path.Combine(patientRoot, planFolderName);

                    Console.WriteLine();
                    Console.WriteLine($"Plan: Course={course.Id}  Plan={plan.Id}  IsTreated={plan.IsTreated}");

                    // ── Planungs-CT ───────────────────────────────────────────────────
                    if (_config.ExportPlanCt)
                    {
                        try
                        {
                            string ctUid = plan.StructureSet?.Image?.Series?.UID ?? "";
                            string ssId  = plan.StructureSet?.Id ?? "";
                            if (!string.IsNullOrWhiteSpace(ctUid))
                            {
                                if (IsCbctStructureSetId(ssId))
                                    Console.WriteLine($"  CT-HINWEIS: StructureSet '{ssId}' klingt nach CBCT – wird trotzdem exportiert.");
                                Console.WriteLine($"  CT: SS='{ssId}'  SeriesUID=…{Tail(ctUid, 20)}");
                                string ctDir = Path.Combine(planRoot, "CT");
                                movedTotal += RunMoveScu("SERIES", "0020,000E", ctUid, patient.Id, ctDir, null);
                            }
                            else
                            {
                                Console.WriteLine($"  CT: keine Series-UID (plan.StructureSet?.Image?.Series?.UID leer, SS='{ssId}').");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  WARN CT Export: {ex.Message}");
                        }
                    }

                    // ── RTSTRUCT ─────────────────────────────────────────────────────
                    if (_config.ExportPlanRtStruct)
                    {
                        try
                        {
                            if (plan.StructureSet != null && !string.IsNullOrWhiteSpace(plan.StructureSet.UID))
                            {
                                string rsDir  = Path.Combine(planRoot, "RTSTRUCT");
                                string rsName = $"RS_{MakeSafe(plan.Id)}_{MakeSafe(plan.StructureSet.Id)}.dcm";
                                movedTotal += RunMoveScu("IMAGE", "0008,0018", plan.StructureSet.UID, patient.Id, rsDir, rsName);
                            }
                            else
                            {
                                Console.WriteLine("  RTSTRUCT: kein StructureSet am Plan.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  WARN RTSTRUCT Export: {ex.Message}");
                        }
                    }

                    // ── RTPLAN ───────────────────────────────────────────────────────
                    if (_config.ExportPlanRtPlan)
                    {
                        try
                        {
                            if (!string.IsNullOrWhiteSpace(plan.UID))
                            {
                                string rpDir  = Path.Combine(planRoot, "RTPLAN");
                                string rpName = $"RP_{MakeSafe(plan.Id)}.dcm";
                                movedTotal += RunMoveScu("IMAGE", "0008,0018", plan.UID, patient.Id, rpDir, rpName);
                            }
                            else
                            {
                                Console.WriteLine("  RTPLAN: Plan UID leer.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  WARN RTPLAN Export: {ex.Message}");
                        }
                    }

                    // ── RTDOSE ───────────────────────────────────────────────────────
                    if (_config.ExportPlanRtDose)
                    {
                        try
                        {
                            if (plan.IsDoseValid && plan.Dose != null && !string.IsNullOrWhiteSpace(plan.Dose.UID))
                            {
                                string rdDir  = Path.Combine(planRoot, "RTDOSE");
                                string rdName = $"RD_{MakeSafe(plan.Id)}.dcm";
                                movedTotal += RunMoveScu("IMAGE", "0008,0018", plan.Dose.UID, patient.Id, rdDir, rdName);
                            }
                            else
                            {
                                Console.WriteLine("  RTDOSE: keine gültige Dosis am Plan.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  WARN RTDOSE Export: {ex.Message}");
                        }
                    }

                    // ── RTIMAGE: Bilder diesem Plan per Beam-ID zuordnen ───────────────────
                    if (_config.ExportAcquiredDoseRtImage && allRtImages.Count > 0)
                    {
                        try
                        {
                            // Beam-IDs des Plans sammeln (nur MV-Felder, kein kV)
                            var beamIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                            try
                            {
                                foreach (var beam in plan.Beams)
                                {
                                    string bid = beam.Id ?? "";
                                    if (bid.Length > 0 &&
                                        bid.IndexOf("kV", StringComparison.OrdinalIgnoreCase) < 0)
                                        beamIds.Add(bid);
                                }
                            }
                            catch { /* BrachyPlan o.ä. ohne Beams */ }

                            if (beamIds.Count > 0)
                            {
                                string rtDir    = Path.Combine(planRoot, "RTIMAGE");
                                int    rtCount  = 0;

                                foreach (var entry in allRtImages)
                                {
                                    if (!beamIds.Contains(entry.ImgId)) continue;

                                    string dayPart    = entry.CreationDateTime.HasValue
                                        ? entry.CreationDateTime.Value.ToLocalTime().ToString("yyyyMMdd")
                                        : "unknown_date";
                                    // Dateiname exakt wie AutoExportMG: Datum_SerienId_FeldId_Typ[_N].dcm
                                    string typeSuffix = ExtractPortalImageTypeSuffix(entry.ImageType);
                                    string destName   = $"{dayPart}_{MakeSafe(entry.SeriesId)}_{MakeSafe(entry.ImgId)}{typeSuffix}.dcm";
                                    string destPath   = Path.Combine(rtDir, destName);

                                    // 1. Datei schon vorhanden → Erfolg, kein movescu (wie AutoExportMG)
                                    if (File.Exists(destPath))
                                    {
                                        Console.WriteLine($"  RTIMAGE bereits vorhanden: {destName}");
                                        exportedImageUids.Add(entry.Uid);
                                        rtCount++;
                                        continue;
                                    }

                                    // 2. UID-Duplikat → AutoExportMG-Pattern: !Add() gibt false wenn schon vorhanden
                                    if (!exportedImageUids.Add(entry.Uid))
                                    {
                                        Console.WriteLine($"  UID-Duplikat übersprungen: {entry.ImgId}");
                                        continue;
                                    }

                                    Console.WriteLine($"  ✓ RTIMAGE [{entry.ImgId}]  {entry.ImageType}  UID=…{Tail(entry.Uid, 18)}");
                                    int n = RunMoveScu("IMAGE", "0008,0018", entry.Uid, patient.Id, rtDir, destName);
                                    movedTotal += n;
                                    if (n > 0) rtCount++;
                                }

                                Console.WriteLine($"  RTIMAGE: {rtCount} Bild(er) für Beams [{string.Join(", ", beamIds.OrderBy(x => x))}]");
                            }
                            else
                            {
                                Console.WriteLine("  RTIMAGE: keine MV-Beam-IDs → übersprungen.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  WARN RTIMAGE Plan: {ex.Message}");
                        }
                    }

                    DeleteEmptyFolders(planRoot);
                }
            }

            Console.WriteLine();
            Console.WriteLine($"Plan-Export: {planCount} Plan(e) verarbeitet" +
                              (_config.OnlyTreatedPlans ? " (Filter: plan.IsTreated)" : " (alle Pläne)") + ".");
            return movedTotal;
        }

        // ─── Treated-Plan-Erkennung ────────────────────────────────────────────────
        private static bool IsTreatedPlan(PlanSetup plan)
        {
            if (plan == null)
                return false;

            // 1. Direkter Zugriff über ExternalPlanSetup.NumberOfFractionsDelivered (zuverlässigste Methode)
            var ext = plan as ExternalPlanSetup;
            if (ext != null)
            {
                try
                {
                    if (ext.IsTreated)
                        return true;
                }
                catch { }
            }

            // 2. ApprovalStatus: TreatmentApproved = für klinische Behandlung freigegeben
            try
            {
                string status = plan.ApprovalStatus.ToString();
                if (status.IndexOf("Treatment", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            catch { }

            // 3. Reflection-Fallback für ältere ESAPI-Versionen / sonstige Plan-Typen
            string[] enumerableProps =
            {
                "TreatmentSessions",
                "TreatmentRecords",
                "TreatmentRecord",
                "DeliveredFractions",
                "SessionRecords"
            };

            foreach (string prop in enumerableProps)
            {
                object value = TryGetPropertyValue(plan, prop);
                if (HasItems(value))
                    return true;
            }

            string[] numericProps =
            {
                "NumberOfFractionsDelivered",
                "DeliveredFractionCount",
                "CompletedFractionCount",
                "TreatedFractionCount",
                "FractionsDelivered"
            };

            foreach (string prop in numericProps)
            {
                object value = TryGetPropertyValue(plan, prop);
                int count;
                if (TryConvertToInt(value, out count) && count > 0)
                    return true;
            }

            string[] textProps =
            {
                "TreatmentStatus",
                "Status"
            };

            foreach (string prop in textProps)
            {
                object value = TryGetPropertyValue(plan, prop);
                string text = value == null ? "" : value.ToString();
                if (text.IndexOf("treated", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("delivered", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("completed", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }

            return false;
        }

        private static object TryGetPropertyValue(object instance, string propertyName)
        {
            if (instance == null || string.IsNullOrWhiteSpace(propertyName))
                return null;

            try
            {
                PropertyInfo pi = instance.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
                if (pi == null || pi.GetIndexParameters().Length > 0)
                    return null;

                return pi.GetValue(instance, null);
            }
            catch
            {
                return null;
            }
        }

        private static bool HasItems(object value)
        {
            if (value == null || value is string)
                return false;

            ICollection collection = value as ICollection;
            if (collection != null)
                return collection.Count > 0;

            IEnumerable enumerable = value as IEnumerable;
            if (enumerable == null)
                return false;

            try
            {
                IEnumerator enumerator = enumerable.GetEnumerator();
                try
                {
                    return enumerator.MoveNext();
                }
                finally
                {
                    IDisposable disposable = enumerator as IDisposable;
                    if (disposable != null)
                        disposable.Dispose();
                }
            }
            catch
            {
                return false;
            }
        }

        private static bool TryConvertToInt(object value, out int result)
        {
            result = 0;
            if (value == null)
                return false;

            try
            {
                if (value is int)
                {
                    result = (int)value;
                    return true;
                }

                if (value is long)
                {
                    long l = (long)value;
                    if (l > int.MaxValue) return false;
                    result = (int)l;
                    return true;
                }

                return int.TryParse(value.ToString(), out result);
            }
            catch
            {
                return false;
            }
        }

        // ─── CBCT: alle als kV/CBCT erkennbaren StructureSet-Image-Serien ───────────
        private static int ExportCbcts(Patient patient, string patientRoot)
        {
            int movedTotal = 0;
            int matched = 0;
            var exportedSeries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            string cbctRoot = Path.Combine(patientRoot, "CBCT");

            Console.WriteLine();
            Console.WriteLine("CBCT: Suche patientenweit nach kV/CBCT-StructureSets.");

            foreach (var ss in patient.StructureSets)
            {
                try
                {
                    string ssId = ss.Id ?? "";
                    if (!IsCbctStructureSetId(ssId))
                        continue;

                    string ctUid = ss.Image?.Series?.UID ?? "";
                    if (string.IsNullOrWhiteSpace(ctUid))
                    {
                        Console.WriteLine($"  CBCT {ssId}: keine CT Series UID.");
                        continue;
                    }

                    if (exportedSeries.Contains(ctUid))
                    {
                        Console.WriteLine($"  CBCT {ssId}: Serie bereits exportiert.");
                        continue;
                    }

                    DateTime? t = null;
                    try { t = ss.Image?.CreationDateTime; } catch { }

                    string datePart = t.HasValue
                        ? t.Value.ToLocalTime().ToString("yyyyMMdd_HHmm")
                        : "unknown_date";

                    string cbctDir = Path.Combine(cbctRoot, $"{datePart}_{MakeSafe(ssId)}");

                    Console.WriteLine($"  CBCT: {ssId}  SeriesUID=…{Tail(ctUid, 18)}");
                    movedTotal += RunMoveScu("SERIES", "0020,000E", ctUid, patient.Id, cbctDir, null);
                    exportedSeries.Add(ctUid);
                    matched++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  WARN CBCT: {ex.Message}");
                }
            }

            Console.WriteLine($"CBCT: {matched} Serie(n) angestoßen.");
            DeleteEmptyFolders(cbctRoot);
            return movedTotal;
        }

        // ─── RTIMAGE: ACQUIRED_DOSE patientenweit – nur nicht plan-verknüpfte Bilder ────
        private static int ExportAcquiredDoseRtImages(Patient patient, string patientRoot,
                                                       HashSet<string> exportedImageUids)
        {
            int movedTotal = 0;
            int matched = 0;
            // UIDs die ExportPlanSets bereits exportiert hat übernehmen
            var exportedImages = new HashSet<string>(exportedImageUids, StringComparer.OrdinalIgnoreCase);

            string rtImageRoot = Path.Combine(patientRoot, "RTIMAGE");

            int studyCount = 0;
            try { studyCount = patient.Studies.Count(); } catch { }

            Console.WriteLine();
            Console.WriteLine($"RTIMAGE: Suche patientenweit nach ACQUIRED_DOSE in {studyCount} Study/Studies.");

            foreach (var study in patient.Studies)
            {
                foreach (var ser in study.Series)
                {
                    try
                    {
                        string modality = "";
                        try { modality = ser.Modality.ToString(); } catch { }

                        if (!modality.Equals("RTIMAGE", StringComparison.OrdinalIgnoreCase))
                            continue;

                        string seriesId = "";
                        try { seriesId = ser.Id ?? ""; } catch { }

                        // kV-Positionierungsbilder nicht als ACQUIRED_DOSE exportieren.
                        if (seriesId.IndexOf("kV", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Console.WriteLine($"  RTIMAGE-Serie übersprungen wegen kV: {seriesId}");
                            continue;
                        }

                        foreach (var img in ser.Images)
                        {
                            try
                            {
                                string imgUid = img.UID ?? "";
                                string imgId  = img.Id  ?? "";

                                if (string.IsNullOrWhiteSpace(imgUid)) continue;
                                if (imgId.IndexOf("kV", StringComparison.OrdinalIgnoreCase) >= 0) continue;

                                string imgType = "";
                                try { imgType = img.ImageType ?? ""; } catch { }

                                if (imgType.IndexOf("ACQUIRED_DOSE", StringComparison.OrdinalIgnoreCase) < 0)
                                    continue;

                                DateTime? t = null;
                                try { t = img.CreationDateTime; } catch { }

                                string dayPart  = t.HasValue
                                    ? t.Value.ToLocalTime().ToString("yyyyMMdd")
                                    : "unknown_date";

                                // Dateiname exakt wie AutoExportMG: Datum_SerienId_FeldId_Typ[_N].dcm (kein Zeitstempel)
                                string rtDir      = Path.Combine(rtImageRoot, dayPart);
                                string typeSuffix = ExtractPortalImageTypeSuffix(imgType);
                                string destName   = $"{dayPart}_{MakeSafe(seriesId)}_{MakeSafe(imgId)}{typeSuffix}.dcm";
                                string destPath   = Path.Combine(rtDir, destName);

                                // 1. Datei vorhanden → Erfolg, kein movescu (wie AutoExportMG)
                                if (File.Exists(destPath))
                                {
                                    Console.WriteLine($"  RTIMAGE bereits vorhanden: {destName}");
                                    exportedImages.Add(imgUid);
                                    matched++;
                                    continue;
                                }

                                // 2. UID-Duplikat → AutoExportMG-Pattern: !Add() gibt false wenn schon vorhanden
                                if (!exportedImages.Add(imgUid))
                                {
                                    Console.WriteLine($"  UID-Duplikat übersprungen: {imgId}");
                                    continue;
                                }

                                Console.WriteLine($"  ✓ RTIMAGE ACQUIRED_DOSE: {imgId}  UID=…{Tail(imgUid, 18)}");
                                movedTotal += RunMoveScu("IMAGE", "0008,0018", imgUid, patient.Id, rtDir, destName);
                                matched++;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"  WARN RTIMAGE IMG: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  WARN RTIMAGE Serie: {ex.Message}");
                    }
                }
            }

            Console.WriteLine($"RTIMAGE: {matched} ACQUIRED_DOSE Image(s) angestoßen.");
            DeleteEmptyFolders(rtImageRoot);
            return movedTotal;
        }

        // Run movescu against the site-configured DICOM export destination.
        private static int RunMoveScu(
            string qLevel,
            string tag,
            string uid,
            string patientId,
            string destDir,
            string overrideFileName)
        {
            if (string.IsNullOrWhiteSpace(uid))
            {
                Console.WriteLine($"  WARN movescu: leere UID für {qLevel}/{tag}.");
                return 0;
            }

            string dcmtkBin = _config.DcmtkPaths.FirstOrDefault(Directory.Exists) ?? "";
            string movescuExe = Path.Combine(dcmtkBin, "movescu.exe");

            if (!File.Exists(movescuExe))
            {
                Console.WriteLine($"  ERROR: movescu.exe nicht gefunden. Geprüfte Pfade: {string.Join(" | ", _config.DcmtkPaths)}");
                return 0;
            }

            string importDir = _config.ScriptExportUsesPatientSubfolder
                ? Path.Combine(_config.EsapiImportBase, patientId)
                : _config.EsapiImportBase;

            Directory.CreateDirectory(importDir);
            Directory.CreateDirectory(destDir);

            var before = new HashSet<string>(
                Directory.Exists(importDir)
                    ? Directory.GetFiles(importDir, "*", SearchOption.AllDirectories)
                    : Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            string args = $"-v -aet {_config.Aet} -aec {_config.Aec} -aem {_config.Aem} -S" +
                          $" -k \"0008,0052={qLevel}\" -k \"{tag}={uid}\"" +
                          $" {_config.DicomHost} {_config.DicomPort}";

            Console.WriteLine($"  movescu {qLevel} {tag}=…{Tail(uid, 20)}");

            var psi = new ProcessStartInfo
            {
                FileName = movescuExe,
                Arguments = args,
                WorkingDirectory = dcmtkBin,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using (var proc = Process.Start(psi))
            {
                if (proc == null)
                {
                    Console.WriteLine("  ERROR: movescu konnte nicht gestartet werden.");
                    return 0;
                }

                // Asynchrones Lesen von StandardError verhindert Deadlocks falls StandardOutput sehr lang wird
                string stderr = "";
                proc.ErrorDataReceived += (s, e) => { if (e.Data != null) stderr += e.Data + Environment.NewLine; };
                proc.BeginErrorReadLine();
                
                // Synchrones Lesen von StandardOutput blockiert bis der Prozess schließt
                string stdout = proc.StandardOutput.ReadToEnd();

                bool exited = proc.WaitForExit(_config.MoveScuTimeoutMs);
                if (!exited)
                {
                    try { proc.Kill(); } catch { }
                    Console.WriteLine($"  ERROR: movescu Timeout nach {_config.MoveScuTimeoutMs / 1000}s.");
                    return 0;
                }

                if (!string.IsNullOrWhiteSpace(stderr))
                    Console.WriteLine($"  movescu stderr: {stderr.Trim()}");
                if (!string.IsNullOrWhiteSpace(stdout) && _config.VerboseMoveScuOutput)
                    Console.WriteLine($"  movescu stdout: {stdout.Trim()}");
            }

            Console.WriteLine("  Warte auf StoreSCP (Dateien werden geschrieben)...");
            int stableCount = 0;
            int lastCount = -1;
            string[] newFiles = new string[0];
            
            // Warte bis die Anzahl der neuen Dateien für 2 Iterationen (2 Sekunden) stabil bleibt
            // Maximal 60 Sekunden warten (falls das Netzwerk extrem langsam ist)
            for (int i = 0; i < 60; i++)
            {
                Thread.Sleep(1000);
                
                newFiles = Directory.GetFiles(importDir, "*", SearchOption.AllDirectories)
                    .Where(f => !before.Contains(f))
                    .ToArray();
                    
                if (newFiles.Length > 0 && newFiles.Length == lastCount)
                {
                    stableCount++;
                    if (stableCount >= 2) break; // 2 Sekunden lang keine neuen Dateien → fertig
                }
                else
                {
                    stableCount = 0;
                }
                
                lastCount = newFiles.Length;
                
                // Falls gar keine neuen Dateien kommen, brechen wir nach 5 Sekunden ohne Änderung ab
                // (Es könnte sein, dass movescu nichts gefunden hat)
                if (newFiles.Length == 0 && i >= 5) break; 
            }

            Console.WriteLine($"  neue Datei(en): {newFiles.Length}");

            int moved = 0;
            int idx = 0;

            foreach (var file in newFiles)
            {
                string fn;

                if (!string.IsNullOrWhiteSpace(overrideFileName) && newFiles.Length == 1)
                    fn = overrideFileName;
                else if (!string.IsNullOrWhiteSpace(overrideFileName))
                    fn = Path.GetFileNameWithoutExtension(overrideFileName)
                         + $"_{++idx}"
                         + Path.GetExtension(overrideFileName);
                else
                    fn = Path.GetFileName(file);

                string dest = Path.Combine(destDir, fn);
                int sfx = 2;

                while (File.Exists(dest))
                {
                    dest = Path.Combine(
                        destDir,
                        Path.GetFileNameWithoutExtension(fn) + $"_{sfx++}" + Path.GetExtension(fn));
                }

                try
                {
                    File.Move(file, dest);
                    Console.WriteLine($"  → {Path.GetFileName(dest)}");
                    moved++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  WARN Move: {ex.Message}");
                }
            }

            DeleteEmptyFolders(importDir);
            DeleteEmptyFolders(destDir);
            return moved;
        }

        private static bool IsCbctStructureSetId(string ssId)
        {
            if (string.IsNullOrWhiteSpace(ssId))
                return false;

            return ssId.IndexOf("kV", StringComparison.OrdinalIgnoreCase) >= 0
                   || ssId.IndexOf("CBCT", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        // ─── RTIMAGE-Hilfsdatentyp und -Sammler ─────────────────────────────────────────
        private struct RtImageEntry
        {
            public string    Uid;
            public string    ImgId;
            public string    SeriesId;
            public string    ImageType;
            public DateTime? CreationDateTime;
        }

        /// <summary>
        /// Sammelt alle ACQUIRED_DOSE-RTIMAGE-Einträge patientenweit.
        /// Kein movescu – nur ESAPI-Metadaten. Schnell.
        /// </summary>
        private static List<RtImageEntry> CollectAcquiredDoseImages(Patient patient)
        {
            var result = new List<RtImageEntry>();
            foreach (var study in patient.Studies)
            {
                foreach (var ser in study.Series)
                {
                    try
                    {
                        string modality = "";
                        try { modality = ser.Modality.ToString(); } catch { }
                        if (!modality.Equals("RTIMAGE", StringComparison.OrdinalIgnoreCase))
                            continue;

                        string seriesId = "";
                        try { seriesId = ser.Id ?? ""; } catch { }
                        if (seriesId.IndexOf("kV", StringComparison.OrdinalIgnoreCase) >= 0)
                            continue; // kV-Setup-Serien überspringen

                        foreach (var img in ser.Images)
                        {
                            try
                            {
                                string uid   = img.UID ?? "";
                                string imgId = img.Id  ?? "";
                                if (string.IsNullOrWhiteSpace(uid)) continue;
                                if (imgId.IndexOf("kV", StringComparison.OrdinalIgnoreCase) >= 0) continue;

                                string imgType = "";
                                try { imgType = img.ImageType ?? ""; } catch { }
                                if (imgType.IndexOf("ACQUIRED_DOSE", StringComparison.OrdinalIgnoreCase) < 0)
                                    continue;

                                DateTime? t = null;
                                try { t = img.CreationDateTime; } catch { }

                                result.Add(new RtImageEntry
                                {
                                    Uid              = uid,
                                    ImgId            = imgId,
                                    SeriesId         = seriesId,
                                    ImageType        = imgType,
                                    CreationDateTime = t,
                                });
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
            }
            return result;
        }

        private static string ExtractPortalImageTypeSuffix(string imageType)
        {
            if (string.IsNullOrWhiteSpace(imageType))
                return "_ACQUIRED_DOSE";

            int portalIdx = imageType.IndexOf("PORTAL", StringComparison.OrdinalIgnoreCase);
            if (portalIdx >= 0)
            {
                int slashIdx = imageType.IndexOf('\\', portalIdx);
                if (slashIdx >= 0 && slashIdx + 1 < imageType.Length)
                    return "_" + MakeSafe(imageType.Substring(slashIdx + 1));
            }

            if (imageType.IndexOf("ACQUIRED_DOSE", StringComparison.OrdinalIgnoreCase) >= 0)
                return "_ACQUIRED_DOSE";

            return "";
        }

        private static void SetupLogging(string exportBase)
        {
            try
            {
                Directory.CreateDirectory(exportBase);
                string logPath = Path.Combine(exportBase, $"excel_export_log_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                var logFileStream = new FileStream(logPath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
                var logWriter = new StreamWriter(logFileStream, Encoding.UTF8) { AutoFlush = true };
                Console.SetOut(new MultiTextWriter(Console.Out, logWriter));
                Console.SetError(new MultiTextWriter(Console.Error, logWriter));
                Console.WriteLine($"Log: {logPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"WARN: Logging in Datei nicht möglich: {ex.Message}");
            }
        }

        private static string MakeSafe(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return "unknown";

            s = s.Replace("/", "_").Replace("\\", "_");

            foreach (char c in Path.GetInvalidFileNameChars())
                s = s.Replace(c, '_');

            return Regex.Replace(s, @"\s+", "_");
        }

        private static string Tail(string s, int n)
        {
            if (string.IsNullOrEmpty(s))
                return "";
            return s.Length <= n ? s : s.Substring(s.Length - n);
        }

        private static void DeleteEmptyFolders(string path)
        {
            if (!Directory.Exists(path))
                return;

            foreach (var dir in Directory.GetDirectories(path))
            {
                DeleteEmptyFolders(dir);

                try
                {
                    if (!Directory.EnumerateFileSystemEntries(dir).Any())
                        Directory.Delete(dir);
                }
                catch { }
            }
        }

        private static void PauseIfConfigured()
        {
            if (!_config.PauseAtEnd)
                return;

            Console.WriteLine();
            Console.WriteLine("Beliebige Taste zum Beenden...");
            Console.ReadKey();
        }
    }

    internal sealed class ExportOptionsForm : Form
    {
        private readonly Settings _settings;

        private CheckBox _cbCbct;
        private CheckBox _cbRtImage;
        private CheckBox _cbPlanSets;
        private CheckBox _cbCt;
        private CheckBox _cbRs;
        private CheckBox _cbRp;
        private CheckBox _cbRd;
        private CheckBox _cbOnlyTreated;
        private CheckBox _cbDebugLimit;
        private NumericUpDown _numDebugMax;
        private NumericUpDown _numExcelColumn;
        private NumericUpDown _numExcelStartRow;
        private TextBox _txtInputExcel;
        private TextBox _txtOutputFolder;
        private Label _lblColumnPreview;

        public Settings ResultSettings { get; private set; }

        public ExportOptionsForm(Settings settings)
        {
            _settings = settings ?? new Settings();
            ResultSettings = _settings;
            InitializeComponent();
            LoadFromSettings();
            UpdateDependentStates();
        }

        private void InitializeComponent()
        {
            Text = "Excel Patient DICOM Export - Optionen";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(760, 650);
            Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            Label title = new Label();
            title.Text = "DICOM-Export aus ARIA über Excel-Patientenliste";
            title.Font = new Font(Font.FontFamily, 12F, FontStyle.Bold);
            title.AutoSize = true;
            title.Location = new Point(16, 14);
            Controls.Add(title);

            Label intro = new Label();
            intro.Text = "Festlegen, welche Daten pro Patient exportiert werden. Danach Start drücken.";
            intro.AutoSize = true;
            intro.Location = new Point(18, 42);
            Controls.Add(intro);

            GroupBox gbInput = new GroupBox();
            gbInput.Text = "Input / Excel";
            gbInput.Location = new Point(16, 72);
            gbInput.Size = new Size(728, 122);
            Controls.Add(gbInput);

            Label lblInput = new Label();
            lblInput.Text = "Excel-Datei:";
            lblInput.Location = new Point(14, 28);
            lblInput.Size = new Size(90, 22);
            gbInput.Controls.Add(lblInput);

            _txtInputExcel = new TextBox();
            _txtInputExcel.Location = new Point(110, 25);
            _txtInputExcel.Size = new Size(490, 23);
            gbInput.Controls.Add(_txtInputExcel);

            Button btnBrowseInput = new Button();
            btnBrowseInput.Text = "Durchsuchen...";
            btnBrowseInput.Location = new Point(610, 24);
            btnBrowseInput.Size = new Size(100, 25);
            btnBrowseInput.Click += delegate { BrowseInputExcel(); };
            gbInput.Controls.Add(btnBrowseInput);

            Label lblColumn = new Label();
            lblColumn.Text = "Patient-ID-Spalte:";
            lblColumn.Location = new Point(14, 67);
            lblColumn.Size = new Size(110, 22);
            gbInput.Controls.Add(lblColumn);

            _numExcelColumn = new NumericUpDown();
            _numExcelColumn.Minimum = 1;
            _numExcelColumn.Maximum = 200;
            _numExcelColumn.Location = new Point(130, 64);
            _numExcelColumn.Size = new Size(70, 23);
            _numExcelColumn.ValueChanged += delegate { UpdateColumnPreview(); };
            gbInput.Controls.Add(_numExcelColumn);

            _lblColumnPreview = new Label();
            _lblColumnPreview.Location = new Point(206, 67);
            _lblColumnPreview.Size = new Size(90, 22);
            gbInput.Controls.Add(_lblColumnPreview);

            Label lblStartRow = new Label();
            lblStartRow.Text = "Startzeile:";
            lblStartRow.Location = new Point(330, 67);
            lblStartRow.Size = new Size(80, 22);
            gbInput.Controls.Add(lblStartRow);

            _numExcelStartRow = new NumericUpDown();
            _numExcelStartRow.Minimum = 1;
            _numExcelStartRow.Maximum = 100000;
            _numExcelStartRow.Location = new Point(410, 64);
            _numExcelStartRow.Size = new Size(90, 23);
            gbInput.Controls.Add(_numExcelStartRow);

            Label lblExcelHelp = new Label();
            lblExcelHelp.Text = "Beispiel: Spalte 1 = A, Startzeile 2 = erste Datenzeile nach Header.";
            lblExcelHelp.Location = new Point(14, 94);
            lblExcelHelp.Size = new Size(660, 20);
            gbInput.Controls.Add(lblExcelHelp);

            GroupBox gbOutput = new GroupBox();
            gbOutput.Text = "Output";
            gbOutput.Location = new Point(16, 202);
            gbOutput.Size = new Size(728, 76);
            Controls.Add(gbOutput);

            Label lblOutput = new Label();
            lblOutput.Text = "Output-Ordner:";
            lblOutput.Location = new Point(14, 30);
            lblOutput.Size = new Size(100, 22);
            gbOutput.Controls.Add(lblOutput);

            _txtOutputFolder = new TextBox();
            _txtOutputFolder.Location = new Point(110, 27);
            _txtOutputFolder.Size = new Size(490, 23);
            gbOutput.Controls.Add(_txtOutputFolder);

            Button btnBrowseOutput = new Button();
            btnBrowseOutput.Text = "Durchsuchen...";
            btnBrowseOutput.Location = new Point(610, 26);
            btnBrowseOutput.Size = new Size(100, 25);
            btnBrowseOutput.Click += delegate { BrowseOutputFolder(); };
            gbOutput.Controls.Add(btnBrowseOutput);

            GroupBox gbExport = new GroupBox();
            gbExport.Text = "Exportumfang pro Patient";
            gbExport.Location = new Point(16, 286);
            gbExport.Size = new Size(350, 224);
            Controls.Add(gbExport);

            _cbCbct = new CheckBox();
            _cbCbct.Text = "CBCT-Serien";
            _cbCbct.Location = new Point(16, 28);
            _cbCbct.Size = new Size(300, 22);
            gbExport.Controls.Add(_cbCbct);

            _cbRtImage = new CheckBox();
            _cbRtImage.Text = "RTIMAGE mit ACQUIRED_DOSE";
            _cbRtImage.Location = new Point(16, 55);
            _cbRtImage.Size = new Size(300, 22);
            gbExport.Controls.Add(_cbRtImage);

            _cbPlanSets = new CheckBox();
            _cbPlanSets.Text = "Plan-Sets exportieren";
            _cbPlanSets.Location = new Point(16, 82);
            _cbPlanSets.Size = new Size(300, 22);
            _cbPlanSets.CheckedChanged += delegate { UpdateDependentStates(); };
            gbExport.Controls.Add(_cbPlanSets);

            _cbCt = new CheckBox();
            _cbCt.Text = "  CT";
            _cbCt.Location = new Point(34, 110);
            _cbCt.Size = new Size(250, 22);
            gbExport.Controls.Add(_cbCt);

            _cbRs = new CheckBox();
            _cbRs.Text = "  RTSTRUCT";
            _cbRs.Location = new Point(34, 136);
            _cbRs.Size = new Size(250, 22);
            gbExport.Controls.Add(_cbRs);

            _cbRp = new CheckBox();
            _cbRp.Text = "  RTPLAN";
            _cbRp.Location = new Point(34, 162);
            _cbRp.Size = new Size(250, 22);
            gbExport.Controls.Add(_cbRp);

            _cbRd = new CheckBox();
            _cbRd.Text = "  RTDOSE";
            _cbRd.Location = new Point(34, 188);
            _cbRd.Size = new Size(250, 22);
            gbExport.Controls.Add(_cbRd);

            GroupBox gbFilter = new GroupBox();
            gbFilter.Text = "Filter / Debug";
            gbFilter.Location = new Point(394, 286);
            gbFilter.Size = new Size(350, 224);
            Controls.Add(gbFilter);

            _cbOnlyTreated = new CheckBox();
            _cbOnlyTreated.Text = "Nur behandelte Pläne exportieren (Default)";
            _cbOnlyTreated.Location = new Point(16, 28);
            _cbOnlyTreated.Size = new Size(310, 22);
            gbFilter.Controls.Add(_cbOnlyTreated);

            Label lblTreated = new Label();
            lblTreated.Text = "Gilt nur für Plan-Sets. CBCT/RTIMAGE werden patientenweit gesucht.";
            lblTreated.Location = new Point(35, 52);
            lblTreated.Size = new Size(300, 38);
            gbFilter.Controls.Add(lblTreated);

            _cbDebugLimit = new CheckBox();
            _cbDebugLimit.Text = "Debug-Zähler aktivieren";
            _cbDebugLimit.Location = new Point(16, 100);
            _cbDebugLimit.Size = new Size(220, 22);
            _cbDebugLimit.CheckedChanged += delegate { UpdateDependentStates(); };
            gbFilter.Controls.Add(_cbDebugLimit);

            Label lblDebug = new Label();
            lblDebug.Text = "Max. Patienten:";
            lblDebug.Location = new Point(35, 130);
            lblDebug.Size = new Size(100, 22);
            gbFilter.Controls.Add(lblDebug);

            _numDebugMax = new NumericUpDown();
            _numDebugMax.Minimum = 1;
            _numDebugMax.Maximum = 100000;
            _numDebugMax.Location = new Point(140, 127);
            _numDebugMax.Size = new Size(90, 23);
            gbFilter.Controls.Add(_numDebugMax);

            Label lblDebugHelp = new Label();
            lblDebugHelp.Text = "Nur wenn aktiv, bricht das Skript nach dieser Anzahl Patienten ab.";
            lblDebugHelp.Location = new Point(35, 158);
            lblDebugHelp.Size = new Size(300, 42);
            gbFilter.Controls.Add(lblDebugHelp);

            Button btnDescription = new Button();
            btnDescription.Text = "Beschreibung...";
            btnDescription.Location = new Point(16, 528);
            btnDescription.Size = new Size(125, 32);
            btnDescription.Click += delegate { ShowDescription(); };
            Controls.Add(btnDescription);

            Button btnStart = new Button();
            btnStart.Text = "Start";
            btnStart.Location = new Point(528, 596);
            btnStart.Size = new Size(100, 32);
            btnStart.Click += delegate { StartClicked(); };
            Controls.Add(btnStart);

            Button btnCancel = new Button();
            btnCancel.Text = "Abbrechen";
            btnCancel.Location = new Point(644, 596);
            btnCancel.Size = new Size(100, 32);
            btnCancel.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
            Controls.Add(btnCancel);

            AcceptButton = btnStart;
            CancelButton = btnCancel;
        }

        private void LoadFromSettings()
        {
            _txtInputExcel.Text = _settings.LastExcelFile ?? "";
            _txtOutputFolder.Text = _settings.ExportBase ?? "";
            _numExcelColumn.Value = Clamp(_settings.ExcelIdColumn, (int)_numExcelColumn.Minimum, (int)_numExcelColumn.Maximum);
            _numExcelStartRow.Value = Clamp(_settings.ExcelStartRow, (int)_numExcelStartRow.Minimum, (int)_numExcelStartRow.Maximum);
            _cbCbct.Checked = _settings.ExportCbct;
            _cbRtImage.Checked = _settings.ExportAcquiredDoseRtImage;
            _cbPlanSets.Checked = _settings.ExportPlanSets;
            _cbCt.Checked = _settings.ExportPlanCt;
            _cbRs.Checked = _settings.ExportPlanRtStruct;
            _cbRp.Checked = _settings.ExportPlanRtPlan;
            _cbRd.Checked = _settings.ExportPlanRtDose;
            _cbOnlyTreated.Checked = _settings.OnlyTreatedPlans;
            _cbDebugLimit.Checked = _settings.DebugLimitEnabled;
            _numDebugMax.Value = Clamp(_settings.DebugMaxPatients, (int)_numDebugMax.Minimum, (int)_numDebugMax.Maximum);
            UpdateColumnPreview();
        }

        private void ApplyToSettings()
        {
            _settings.LastExcelFile = (_txtInputExcel.Text ?? "").Trim();
            _settings.ExportBase = (_txtOutputFolder.Text ?? "").Trim();
            _settings.ExcelIdColumn = (int)_numExcelColumn.Value;
            _settings.ExcelStartRow = (int)_numExcelStartRow.Value;
            _settings.ExportCbct = _cbCbct.Checked;
            _settings.ExportAcquiredDoseRtImage = _cbRtImage.Checked;
            _settings.ExportPlanSets = _cbPlanSets.Checked;
            _settings.ExportPlanCt = _cbCt.Checked;
            _settings.ExportPlanRtStruct = _cbRs.Checked;
            _settings.ExportPlanRtPlan = _cbRp.Checked;
            _settings.ExportPlanRtDose = _cbRd.Checked;
            _settings.OnlyTreatedPlans = _cbOnlyTreated.Checked;
            _settings.DebugLimitEnabled = _cbDebugLimit.Checked;
            _settings.DebugMaxPatients = (int)_numDebugMax.Value;

            if (!string.IsNullOrWhiteSpace(_settings.LastExcelFile))
            {
                try
                {
                    string d = Path.GetDirectoryName(_settings.LastExcelFile);
                    if (!string.IsNullOrWhiteSpace(d))
                        _settings.ExcelInitialDirectory = d;
                }
                catch { }
            }
        }

        private void StartClicked()
        {
            if (string.IsNullOrWhiteSpace(_txtInputExcel.Text) || !File.Exists(_txtInputExcel.Text.Trim()))
            {
                if (!BrowseInputExcel())
                    return;
            }

            if (string.IsNullOrWhiteSpace(_txtOutputFolder.Text))
            {
                MessageBox.Show(this, "Bitte einen Output-Ordner angeben.", "Fehlender Output-Ordner", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ApplyToSettings();
            ResultSettings = _settings;
            DialogResult = DialogResult.OK;
            Close();
        }

        private bool BrowseInputExcel()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Excel-Datei mit Patient-IDs auswählen";
                ofd.Filter = "Excel-Dateien (*.xlsx)|*.xlsx|Alle Dateien (*.*)|*.*";
                ofd.CheckFileExists = true;
                ofd.CheckPathExists = true;
                ofd.InitialDirectory = GetInitialExcelDirectory();

                if (ofd.ShowDialog(this) != DialogResult.OK)
                    return false;

                _txtInputExcel.Text = ofd.FileName;
                return true;
            }
        }

        private void BrowseOutputFolder()
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Output-Ordner auswählen";
                if (!string.IsNullOrWhiteSpace(_txtOutputFolder.Text) && Directory.Exists(_txtOutputFolder.Text.Trim()))
                    fbd.SelectedPath = _txtOutputFolder.Text.Trim();

                if (fbd.ShowDialog(this) == DialogResult.OK)
                    _txtOutputFolder.Text = fbd.SelectedPath;
            }
        }

        private string GetInitialExcelDirectory()
        {
            string fromFile = _txtInputExcel.Text;
            if (!string.IsNullOrWhiteSpace(fromFile))
            {
                try
                {
                    string d = Path.GetDirectoryName(fromFile.Trim());
                    if (!string.IsNullOrWhiteSpace(d) && Directory.Exists(d))
                        return d;
                }
                catch { }
            }

            if (!string.IsNullOrWhiteSpace(_settings.ExcelInitialDirectory) && Directory.Exists(_settings.ExcelInitialDirectory))
                return _settings.ExcelInitialDirectory;

            return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        }

        private void UpdateDependentStates()
        {
            bool planEnabled = _cbPlanSets != null && _cbPlanSets.Checked;
            if (_cbCt != null) _cbCt.Enabled = planEnabled;
            if (_cbRs != null) _cbRs.Enabled = planEnabled;
            if (_cbRp != null) _cbRp.Enabled = planEnabled;
            if (_cbRd != null) _cbRd.Enabled = planEnabled;
            if (_cbOnlyTreated != null) _cbOnlyTreated.Enabled = planEnabled;
            if (_numDebugMax != null) _numDebugMax.Enabled = _cbDebugLimit != null && _cbDebugLimit.Checked;
        }

        private void UpdateColumnPreview()
        {
            if (_lblColumnPreview != null)
                _lblColumnPreview.Text = "= " + ColumnNumberToName((int)_numExcelColumn.Value);
        }

        private void ShowDescription()
        {
            MessageBox.Show(this, BuildDescriptionText(), "Beschreibung der Exportoptionen", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private string BuildDescriptionText()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Input / Excel");
            sb.AppendLine("Die Patient-ID wird aus dem ersten Worksheet gelesen. Die Spalte ist numerisch angegeben: 1 = A, 2 = B usw. Die Startzeile ist die erste Zeile mit Patientendaten, meist 2 nach einer Header-Zeile.");
            sb.AppendLine();
            sb.AppendLine("CBCT");
            sb.AppendLine("Sucht patientenweit StructureSets, deren ID kV oder CBCT enthält, nimmt die zugehörige Bildserie und exportiert diese CT-Serie per SeriesInstanceUID.");
            sb.AppendLine();
            sb.AppendLine("RTIMAGE mit ACQUIRED_DOSE");
            sb.AppendLine("Sucht patientenweit RTIMAGE-Serien und exportiert nur Bilder, deren ImageType ACQUIRED_DOSE enthält. kV-Positionierungsbilder werden übersprungen. UID-Duplikate werden innerhalb eines Laufs nicht erneut exportiert.");
            sb.AppendLine();
            sb.AppendLine("Plan-Sets");
            sb.AppendLine("Exportiert planbezogen CT, RTSTRUCT, RTPLAN und RTDOSE in getrennte Unterordner. Die Unterschalter bestimmen, welche Modalitäten tatsächlich exportiert werden.");
            sb.AppendLine();
            sb.AppendLine("Nur behandelte Pläne");
            sb.AppendLine("Default. Betrifft nur Plan-Sets. Das Skript versucht über ESAPI-Properties wie TreatmentSessions, TreatmentRecords oder Delivered-Fraction-Zähler zu erkennen, ob ein Plan behandelt wurde. Kann die ESAPI-Version diese Information nicht liefern, wird der Plan bei aktivem Filter übersprungen.");
            sb.AppendLine();
            sb.AppendLine("Debug-Zähler");
            sb.AppendLine("Nur wenn aktiv, stoppt der Lauf nach der angegebenen Anzahl Patienten. Ohne Aktivierung werden alle Patient-IDs aus der Excel verarbeitet.");
            sb.AppendLine();
            sb.AppendLine("Persistenz");
            sb.AppendLine("Nach Start werden Input-Datei, letzter Excel-Ordner, Output-Ordner und alle Checkboxen in settings.ini gespeichert.");
            return sb.ToString();
        }

        private static int Clamp(int value, int min, int max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }

        private static string ColumnNumberToName(int columnNumber)
        {
            if (columnNumber < 1)
                columnNumber = 1;

            string name = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                name = Convert.ToChar('A' + modulo) + name;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return name;
        }
    }

    internal sealed class MultiTextWriter : TextWriter
    {
        private readonly TextWriter[] _writers;

        public MultiTextWriter(params TextWriter[] writers)
        {
            _writers = writers.Where(w => w != null).ToArray();
        }

        public override Encoding Encoding => Encoding.UTF8;

        public override void WriteLine(string value)
        {
            string line = value == null
                ? ""
                : $"[{DateTime.Now:HH:mm:ss}] {value}";

            foreach (var w in _writers)
                w.WriteLine(line);
        }

        public override void Write(string value)
        {
            foreach (var w in _writers)
                w.Write(value);
        }

        public override void Write(char value)
        {
            foreach (var w in _writers)
                w.Write(value);
        }
    }

    internal sealed class Settings
    {
        // DICOM daemon connection – configure all site-specific values in settings.ini
        public string Aet { get; set; } = "DCMTK";
        public string Aec { get; set; } = "DICOM_ESAPI";
        public string Aem { get; set; } = "ScriptExport";
        public string DicomHost { get; set; } = "";
        public int DicomPort { get; set; } = 51402;
        public string EsapiImportBase { get; set; } = "";

        public List<string> DcmtkPaths { get; set; } = new List<string>
        {
            @"Assets\DCMTK\bin",   // extract Assets\DCMTK.zip → Assets\DCMTK\bin\movescu.exe
            @"C:\dcmtk\bin"        // common system-wide install path
        };

        public string ExportBase { get; set; } = "";

        public string ExcelInitialDirectory { get; set; } = "";

        public string LastExcelFile { get; set; } = "";
        public int ExcelIdColumn { get; set; } = 1;
        public int ExcelStartRow { get; set; } = 2;

        // Hauptschalter
        public bool ExportCbct { get; set; } = false;
        public bool ExportAcquiredDoseRtImage { get; set; } = true;
        public bool ExportPlanSets { get; set; } = true;

        // Plan-Set-Unterschalter
        public bool ExportPlanCt { get; set; } = true;
        public bool ExportPlanRtStruct { get; set; } = true;
        public bool ExportPlanRtPlan { get; set; } = true;
        public bool ExportPlanRtDose { get; set; } = true;

        public bool OnlyTreatedPlans { get; set; } = true;
        public bool DebugLimitEnabled { get; set; } = false;
        public int DebugMaxPatients { get; set; } = 5;
        public int MaxPatients { get; set; } = int.MaxValue; // Legacy: nicht mehr aktiv, wenn DEBUG_LIMIT_ENABLED=FALSE
        public bool ScriptExportUsesPatientSubfolder { get; set; } = false;
        public int ScriptExportWriteWaitMs { get; set; } = 2000;
        public int MoveScuTimeoutMs { get; set; } = 180000;
        public bool VerboseMoveScuOutput { get; set; } = false;
        public bool PauseAtEnd { get; set; } = true;

        public static Settings Load(string iniPath)
        {
            var s = new Settings();

            if (!File.Exists(iniPath))
                return s;

            foreach (string raw in File.ReadAllLines(iniPath))
            {
                string line = raw.Trim();
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#") || line.StartsWith(";"))
                    continue;

                int idx = line.IndexOf('=');
                if (idx <= 0)
                    continue;

                string key = line.Substring(0, idx).Trim().ToUpperInvariant();
                string value = line.Substring(idx + 1).Trim();

                switch (key)
                {
                    case "AET":
                        s.Aet = value;
                        break;
                    case "AEC":
                        s.Aec = value;
                        break;
                    case "AEM":
                        s.Aem = value;
                        break;
                    case "DICOM_HOST":
                        s.DicomHost = value;
                        break;
                    case "DICOM_PORT":
                        int port;
                        if (int.TryParse(value, out port)) s.DicomPort = port;
                        break;
                    case "ESAPI_IMPORT_BASE":
                        s.EsapiImportBase = value;
                        break;
                    case "DCMTK_PATHS":
                        s.DcmtkPaths = value
                            .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(x => x.Trim())
                            .Where(x => x.Length > 0)
                            .ToList();
                        break;
                    case "EXPORT_BASE":
                        s.ExportBase = value;
                        break;
                    case "EXCEL_INITIAL_DIRECTORY":
                        s.ExcelInitialDirectory = value;
                        break;
                    case "LAST_EXCEL_FILE":
                    case "INPUT_EXCEL_FILE":
                        s.LastExcelFile = value;
                        break;
                    case "EXCEL_ID_COLUMN":
                    case "ID_COLUMN":
                        int idCol;
                        if (int.TryParse(value, out idCol) && idCol > 0) s.ExcelIdColumn = idCol;
                        break;
                    case "EXCEL_START_ROW":
                    case "START_ROW":
                        int startRow;
                        if (int.TryParse(value, out startRow) && startRow > 0) s.ExcelStartRow = startRow;
                        break;

                    case "EXPORT_CBCT":
                        s.ExportCbct = ParseBool(value, s.ExportCbct);
                        break;
                    case "EXPORT_ACQUIRED_DOSE_RTIMAGE":
                    case "EXPORT_RTIMAGE":
                        s.ExportAcquiredDoseRtImage = ParseBool(value, s.ExportAcquiredDoseRtImage);
                        break;
                    case "EXPORT_PLAN_SETS":
                    case "EXPORT_PLANSET":
                        s.ExportPlanSets = ParseBool(value, s.ExportPlanSets);
                        break;

                    case "EXPORT_CT":
                        s.ExportPlanCt = ParseBool(value, s.ExportPlanCt);
                        break;
                    case "EXPORT_RTSTRUCT":
                    case "EXPORT_RS":
                        s.ExportPlanRtStruct = ParseBool(value, s.ExportPlanRtStruct);
                        break;
                    case "EXPORT_RTPLAN":
                    case "EXPORT_RP":
                        s.ExportPlanRtPlan = ParseBool(value, s.ExportPlanRtPlan);
                        break;
                    case "EXPORT_RTDOSE":
                    case "EXPORT_RD":
                        s.ExportPlanRtDose = ParseBool(value, s.ExportPlanRtDose);
                        break;

                    case "ONLY_TREATED_PLANS":
                    case "ONLY_TREATED":
                        s.OnlyTreatedPlans = ParseBool(value, s.OnlyTreatedPlans);
                        break;
                    case "DEBUG_LIMIT_ENABLED":
                    case "DEBUG_COUNTER_ENABLED":
                        s.DebugLimitEnabled = ParseBool(value, s.DebugLimitEnabled);
                        break;
                    case "DEBUG_MAX_PATIENTS":
                    case "DEBUG_COUNTER":
                        int debugN;
                        if (int.TryParse(value, out debugN) && debugN > 0) s.DebugMaxPatients = debugN;
                        break;

                    case "COUNTER":
                    case "MAX_PATIENTS":
                        int n;
                        if (int.TryParse(value, out n) && n > 0)
                        {
                            s.MaxPatients = n;
                            s.DebugMaxPatients = n;
                        }
                        break;
                    case "SCRIPTEXPORT_USES_PATIENT_SUBFOLDER":
                        s.ScriptExportUsesPatientSubfolder = ParseBool(value, s.ScriptExportUsesPatientSubfolder);
                        break;
                    case "SCRIPTEXPORT_WRITE_WAIT_MS":
                        int writeWait;
                        if (int.TryParse(value, out writeWait) && writeWait >= 0) s.ScriptExportWriteWaitMs = writeWait;
                        break;
                    case "MOVESCU_TIMEOUT_MS":
                        int timeout;
                        if (int.TryParse(value, out timeout)) s.MoveScuTimeoutMs = timeout;
                        break;
                    case "VERBOSE_MOVESCU_OUTPUT":
                        s.VerboseMoveScuOutput = ParseBool(value, s.VerboseMoveScuOutput);
                        break;
                    case "PAUSE_AT_END":
                        s.PauseAtEnd = ParseBool(value, s.PauseAtEnd);
                        break;
                }
            }

            if (s.DcmtkPaths == null || s.DcmtkPaths.Count == 0)
                s.DcmtkPaths = new List<string> { @"Assets\DCMTK\bin", @"C:\dcmtk\bin" };

            return s;
        }

        public void Save(string iniPath)
        {
            using (StreamWriter sw = new StreamWriter(iniPath, false, Encoding.UTF8))
            {
                sw.WriteLine("# ExcelPatientDicomExportMG settings");
                sw.WriteLine("# Diese Datei wird beim Start über das Optionsfenster aktualisiert.");
                sw.WriteLine();
                sw.WriteLine("AET=" + Aet);
                sw.WriteLine("AEC=" + Aec);
                sw.WriteLine("AEM=" + Aem);
                sw.WriteLine("DICOM_HOST=" + DicomHost);
                sw.WriteLine("DICOM_PORT=" + DicomPort);
                sw.WriteLine("ESAPI_IMPORT_BASE=" + EsapiImportBase);
                sw.WriteLine("DCMTK_PATHS=" + string.Join(";", DcmtkPaths ?? new List<string>()));
                sw.WriteLine("SCRIPTEXPORT_USES_PATIENT_SUBFOLDER=" + BoolText(ScriptExportUsesPatientSubfolder));
                sw.WriteLine("SCRIPTEXPORT_WRITE_WAIT_MS=" + ScriptExportWriteWaitMs);
                sw.WriteLine();
                sw.WriteLine("EXPORT_BASE=" + ExportBase);
                sw.WriteLine("EXCEL_INITIAL_DIRECTORY=" + ExcelInitialDirectory);
                sw.WriteLine("LAST_EXCEL_FILE=" + LastExcelFile);
                sw.WriteLine("EXCEL_ID_COLUMN=" + ExcelIdColumn);
                sw.WriteLine("EXCEL_START_ROW=" + ExcelStartRow);
                sw.WriteLine();
                sw.WriteLine("EXPORT_CBCT=" + BoolText(ExportCbct));
                sw.WriteLine("EXPORT_ACQUIRED_DOSE_RTIMAGE=" + BoolText(ExportAcquiredDoseRtImage));
                sw.WriteLine("EXPORT_PLAN_SETS=" + BoolText(ExportPlanSets));
                sw.WriteLine("EXPORT_CT=" + BoolText(ExportPlanCt));
                sw.WriteLine("EXPORT_RTSTRUCT=" + BoolText(ExportPlanRtStruct));
                sw.WriteLine("EXPORT_RTPLAN=" + BoolText(ExportPlanRtPlan));
                sw.WriteLine("EXPORT_RTDOSE=" + BoolText(ExportPlanRtDose));
                sw.WriteLine("ONLY_TREATED_PLANS=" + BoolText(OnlyTreatedPlans));
                sw.WriteLine();
                sw.WriteLine("DEBUG_LIMIT_ENABLED=" + BoolText(DebugLimitEnabled));
                sw.WriteLine("DEBUG_MAX_PATIENTS=" + DebugMaxPatients);
                sw.WriteLine();
                sw.WriteLine("MOVESCU_TIMEOUT_MS=" + MoveScuTimeoutMs);
                sw.WriteLine("VERBOSE_MOVESCU_OUTPUT=" + BoolText(VerboseMoveScuOutput));
                sw.WriteLine("PAUSE_AT_END=" + BoolText(PauseAtEnd));
            }
        }

        private static string BoolText(bool value)
        {
            return value ? "TRUE" : "FALSE";
        }

        private static bool ParseBool(string value, bool defaultValue)
        {
            if (string.IsNullOrWhiteSpace(value))
                return defaultValue;

            string v = value.Trim().ToUpperInvariant();

            if (v == "TRUE" || v == "1" || v == "YES" || v == "JA" || v == "J")
                return true;

            if (v == "FALSE" || v == "0" || v == "NO" || v == "NEIN" || v == "N")
                return false;

            return defaultValue;
        }
    }
}
