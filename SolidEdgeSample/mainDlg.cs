using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using KSM_SolidEdge.Properties;

namespace KSM_SolidEdge
{
    public partial class mainDlg : Form
    {
        private bool isRunning = false;
        private CancellationTokenSource cts;

        private DataTable dtLog;
        private double _movingAverageSeconds = -1; // 이동 평균 저장용
        private int _lastRunSuccessCount;
        private int _lastRunFailCount;

        private AppConfig config;

        // 경로 설정: 실행 폴더 기준 상위 루트 설정
        private static readonly string AppDirectory = AppDomain.CurrentDomain.BaseDirectory;
        private readonly string RootDirectory = Directory.GetParent(AppDirectory.TrimEnd('\\'))?.FullName ?? AppDirectory;

        private string TemplatePath => Path.Combine(RootDirectory, "Template", "KSM_Quick.dft");
        private string LogDirectory => Path.Combine(RootDirectory, "Log");

        public mainDlg()
        {
            InitializeComponent();
            config = AppConfig.Load(); // 설정 로드
            ApplyExcelRetryPanelFromAppConfig();
            this.Load += mainDlg_Load;
            this.FormClosing += mainDlg_FormClosing;
        }

        #region [초기화 및 종료]
        private async void mainDlg_Load(object sender, EventArgs e)
        {
            LoadSavedPathsToUi();
            progressInit();
            imageTypeInit();
            InitLogDataTable();
            UpdateRecentRowsHintLabel();
            btnCancel.Enabled = false;

            lblStatus.Text = "대기 중...";
        }

        private void mainDlg_FormClosing(object sender, FormClosingEventArgs e)
        {
            SavePathsFromUi();

            SolidEdgeFramework.Application _seApp = SolidEdgeConnector.GetInstance(false);

            try
            {
                if (_seApp != null)
                {
                    // 1. 열려 있는 문서가 있다면 모두 저장하지 않고 닫기 (안전장치)
                    if (SolidEdgeConnector.App.Documents.Count > 0)
                    {
                        SolidEdgeConnector.App.Documents.Close();
                    }

                    // 2. 솔리드 엣지 자체를 종료 (화면에 보이게 하지 않음)                    
                    SolidEdgeConnector.App.Quit();

                    // 3. COM 객체 해제
                    Marshal.ReleaseComObject(SolidEdgeConnector.App);
                    SolidEdgeConnector.App = null;
                }
            }
            catch { }


        }

        /// <summary>
        /// STA 워커가 UI 컨트롤에 동기 <see cref="Control.Invoke(System.Action)"/>할 수 있으므로,
        /// UI 스레드에서 STA 작업 완료를 기다릴 때 메시지 큐를 처리합니다. (FormClosing / finally 대기 시 교착 방지)
        /// </summary>
        private static void WaitForStaTask(Task task)
        {
            while (!task.IsCompleted)
            {
                Application.DoEvents();
                Thread.Sleep(5);
            }
            task.GetAwaiter().GetResult();
        }

        /// <summary>
        /// config.json 의 ShowExcelRetryPanel 이 false 이면 실패 재시도 엑셀 UI를 숨기고 폼 레이아웃을 맞춥니다.
        /// </summary>
        private void ApplyExcelRetryPanelFromAppConfig()
        {
            bool show = config.ShowExcelRetryPanel;
            lblResultExcel.Visible = show;
            txtResultExcelPath.Visible = show;
            btnResultExcelBrowse.Visible = show;

            const int grpSourceTop = 10;
            const int gap = 6;
            int grpSourceH = show ? 108 : 80;
            grpSource.Height = grpSourceH;

            int grpOutputTop = grpSourceTop + grpSourceH + gap;
            grpOutput.Top = grpOutputTop;

            int groupBox1Top = grpOutputTop + grpOutput.Height + gap;
            groupBox1.Top = groupBox1Top;

            int groupBox2Top = groupBox1Top + groupBox1.Height + gap;
            groupBox2.Top = groupBox2Top;

            int buttonsTop = groupBox2Top + groupBox2.Height + 10;
            btnExecute.Top = buttonsTop;
            btnReset.Top = buttonsTop;
            btnCancel.Top = buttonsTop;
            progressBar1.Top = buttonsTop + 3;
            lblStatus.Top = buttonsTop + 29;
            ClientSize = new Size(ClientSize.Width, buttonsTop + 62);
        }
        #endregion        

        #region [메인 프로세스]
        private async void btnExecute_Click(object sender, EventArgs e)
        {
            if (isRunning) return;

            if (!File.Exists(TemplatePath))
            {
                MessageBox.Show("템플릿 파일을 찾을 수 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string source = txtSourcePath.Text;
            string output = txtOutputPath.Text;
            string type = cboImageType.Text.ToLower();

            config = AppConfig.Load();
            ApplyExcelRetryPanelFromAppConfig();
            UpdateRecentRowsHintLabel();

            string resultExcelPath = (config.ShowExcelRetryPanel ? txtResultExcelPath.Text : string.Empty).Trim();

            if (string.IsNullOrEmpty(source))
            {
                WriteLog("[오류] 소스 폴더 경로가 비어 있습니다.");
                Invoke((Action)(() => MessageBox.Show(this, "소스 폴더 경로를 지정해 주세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error)));
                return;
            }

            if (string.IsNullOrEmpty(output))
            {
                WriteLog("[오류] 출력 폴더 경로가 비어 있습니다.");
                Invoke((Action)(() => MessageBox.Show(this, "출력 폴더 경로를 지정해 주세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error)));
                return;
            }

            lblStatus.Text = "시스템 준비 중 (Solid Edge 연결 중)...";

            if (!Directory.Exists(output)) Directory.CreateDirectory(output);

            SavePathsFromUi();

            SetUIState(true);

            cts = new CancellationTokenSource();           
            dtLog.Clear(); // 새 작업 시작 시 로그 초기화
            _movingAverageSeconds = -1;

            try
            {
                // --- [추가] 절전 모드 방지 시작 ---
                SleepModeController.PreventSleep();
                WriteLog("[시스템] 절전 모드 진입 방지 활성화");

                await Task.Factory.StartNew(() =>
                {
                    if (SolidEdgeConnector.App == null)
                        SolidEdgeConnector.App = SolidEdgeConnector.GetInstance();

                    SolidEdgeConnector.App.DisplayAlerts = false;
                    SolidEdgeConnector.App.ScreenUpdating = false;
                    SolidEdgeConnector.App.Visible = false;

                    
                    RunProcess(source, output, type, resultExcelPath, cts.Token);
                }, cts.Token, TaskCreationOptions.DenyChildAttach, StaTaskScheduler.Instance);

                int totalDone = _lastRunSuccessCount + _lastRunFailCount;
                string summary = totalDone > 0
                    ? $"성공: {_lastRunSuccessCount}건\r\n실패: {_lastRunFailCount}건\r\n(총 {totalDone}건)"
                    : "처리된 파일이 없습니다.";
                DialogResult result = MessageBox.Show(this,
                    $"모든 변환 작업이 완료되었습니다.\r\n\r\n{summary}\r\n\r\n출력 폴더를 여시겠습니까?",
                    "완료",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (result == DialogResult.Yes)
                {
                    // 출력 폴더 자동 열기
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = output,
                        UseShellExecute = true,
                        Verb = "open"
                    });
                }
            }
            catch (OperationCanceledException)
            {
                this.Cursor = Cursors.Default;
                MessageBox.Show("작업이 취소되었습니다.");
            }
            catch (Exception ex)
            {
                WriteLog($"[치명적 오류] {ex.Message}");
                MessageBox.Show($"오류 발생: {ex.Message}");
            }
            finally
            {
                SleepModeController.AllowSleep();
                WriteLog("[시스템] 절전 모드 진입 방지 해제");

                SetUIState(false);

                try
                {
                    var restoreUi = Task.Factory.StartNew(() =>
                    {
                        if (SolidEdgeConnector.App != null)
                        {
                            SolidEdgeConnector.App.ScreenUpdating = true;
                            SolidEdgeConnector.App.DisplayAlerts = true;
                        }
                    }, CancellationToken.None, TaskCreationOptions.DenyChildAttach, StaTaskScheduler.Instance);

                    WaitForStaTask(restoreUi);
                }
                catch { }

                progressInit();
            }
        }

        private void btnReset_Click(object sender, EventArgs e) => dataGridView1.Rows.Clear();

        /// <summary>
        /// 저장된 Result 엑셀(ProcessLog 시트)에서 Result 값이 '실패'인 행의 파일명을 수집합니다.
        /// </summary>
        private static HashSet<string> LoadFailedFilenamesFromResultExcel(string excelPath)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var wb = new XLWorkbook(excelPath))
            {
                IXLWorksheet ws = wb.Worksheet("ProcessLog");
                if (ws == null)
                {
                    if (!wb.Worksheets.Any())
                        throw new InvalidOperationException("워크시트가 없습니다.");
                    ws = wb.Worksheets.ElementAt(0);
                }

                var firstRow = ws.FirstRowUsed();
                if (firstRow == null)
                    return set;

                int colFile = -1, colResult = -1;
                foreach (var cell in firstRow.CellsUsed())
                {
                    string h = cell.GetString().Trim();
                    if (h.Equals("FileName", StringComparison.OrdinalIgnoreCase) || h.Equals("파일명", StringComparison.OrdinalIgnoreCase))
                        colFile = cell.Address.ColumnNumber;
                    if (h.Equals("Result", StringComparison.OrdinalIgnoreCase) || h.Equals("결과", StringComparison.OrdinalIgnoreCase))
                        colResult = cell.Address.ColumnNumber;
                }

                if (colFile < 0 || colResult < 0)
                    throw new InvalidOperationException("FileName(또는 파일명), Result(또는 결과) 열을 찾을 수 없습니다.");

                var lastRow = ws.LastRowUsed();
                if (lastRow == null)
                    return set;

                int headerR = firstRow.RowNumber();
                int endR = lastRow.RowNumber();
                for (int r = headerR + 1; r <= endR; r++)
                {
                    string result = ws.Cell(r, colResult).GetString().Trim();

                    if (result != "실패")
                        continue;
                    string fn = ws.Cell(r, colFile).GetString().Trim();
                    if (string.IsNullOrEmpty(fn))
                        continue;
                    set.Add(Path.GetFileName(fn));
                }
            }

            return set;
        }

        /// <summary>소스 폴더에서 .asm / .par 파일 경로를 이름순으로 반환합니다.</summary>
        private static string[] EnumerateAsmParInDirectory(string sourceDirectory)
        {
            return Directory.GetFiles(sourceDirectory, "*.*")
                .Where(f => f.EndsWith(".asm", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".par", StringComparison.OrdinalIgnoreCase))
                .OrderBy(f => f)
                .ToArray();
        }

        private void RunProcess(string source, string output, string type, string resultExcelPath, CancellationToken token)
        {
            source = (source ?? string.Empty).Trim();
            resultExcelPath = (resultExcelPath ?? string.Empty).Trim();
            _lastRunSuccessCount = 0;
            _lastRunFailCount = 0;

            string[] files;

            if (string.IsNullOrEmpty(resultExcelPath))
            {
                // --- 소스 폴더만 지정: 해당 폴더의 .asm / .par 전체 변환 ---
                WriteLog("[모드] 소스 폴더 전체 변환");
                files = EnumerateAsmParInDirectory(source);
                if (files.Length == 0)
                {
                    WriteLog("[알림] 소스 폴더에 .asm 또는 .par 파일이 없습니다.");
                    return;
                }
            }
            else
            {
                // --- 결과 엑셀 지정: Result 열이 '실패'인 파일만 소스 폴더에서 골라 재변환 ---
                if (!File.Exists(resultExcelPath))
                {
                    WriteLog($"[엑셀] 파일 없음: {resultExcelPath}");
                    Invoke((Action)(() => MessageBox.Show(this, "결과 엑셀 파일을 찾을 수 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error)));
                    return;
                }

                HashSet<string> failedNames;
                try
                {
                    failedNames = LoadFailedFilenamesFromResultExcel(resultExcelPath);
                }
                catch (Exception ex)
                {
                    WriteLog($"[엑셀] 읽기 오류: {ex.Message}");
                    Invoke((Action)(() => MessageBox.Show(this, $"결과 엑셀을 읽을 수 없습니다.\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error)));
                    return;
                }

                if (failedNames.Count == 0)
                {
                    WriteLog("[엑셀] Result 열이 '실패'인 행이 없습니다.");
                    Invoke((Action)(() => MessageBox.Show(this, "엑셀에서 결과(Result)가 '실패'인 행을 찾지 못했습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information)));
                    return;
                }

                string[] allInSource = EnumerateAsmParInDirectory(source);
                files = allInSource.Where(f => failedNames.Contains(Path.GetFileName(f))).ToArray();
                WriteLog($"[모드] 엑셀 실패 재시도 — 대상 파일 {files.Length}개 (엑셀 실패 행 {failedNames.Count}개, 소스 내 .asm/.par {allInSource.Length}개)");

                if (files.Length == 0)
                {
                    Invoke((Action)(() => MessageBox.Show(this, "소스 폴더에 엑셀에 나온 실패 파일과 일치하는 .asm/.par 파일이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information)));
                    return;
                }
            }

            int total = files.Length;
            Invoke((Action)(() => {
                lblStatus.Text = "변환 시작...";
                dataGridView1.Rows.Clear();
                progressBar1.Maximum = total;
            }));

            DateTime startTime = DateTime.Now;
            int no = 1;            

            foreach (string file in files)
            {
                token.ThrowIfCancellationRequested();

                // --- [최적화] 100개마다 UI 그리드 초기화 ---
                if (no % config.GridClearInterval == 1)
                {
                    Invoke((Action)(() => dataGridView1.Rows.Clear()));
                }

                string fileName = Path.GetFileName(file);
                int sheetCount = 0;

                DataGridViewRow row = (DataGridViewRow)Invoke(new Func<DataGridViewRow>(() => AddProcessingRow(no, fileName)));

                SolidEdgeFramework.SolidEdgeDocument seDoc = null;
                SolidEdgeDraft.DraftDocument draftDoc = null;

                try
                {
                    // 최적화된 파일 열기 (ReadOnly, IgnoreLinks)
                    seDoc = (SolidEdgeFramework.SolidEdgeDocument)SolidEdgeConnector.App.Documents.Open(file, Type.Missing, true, true);
                    draftDoc = (SolidEdgeDraft.DraftDocument)SolidEdgeConnector.App.Documents.Add("SolidEdge.DraftDocument", TemplatePath);

                    draftDoc.PopulateQuicksheetTemplate(seDoc.FullName);
                    draftDoc.UpdateAll(false);

                    dynamic sections = draftDoc.Sections;
                    dynamic workingSection = sections.WorkingSection;
                    string nameWithoutExt = Path.GetFileNameWithoutExtension(file);

                    foreach (var item in workingSection.Sheets)
                    {
                        token.ThrowIfCancellationRequested();
                        SolidEdgeDraft.Sheet sheet = (SolidEdgeDraft.Sheet)item;

                        if (sheet.Name == "Sheet1")
                        {
                            Marshal.ReleaseComObject(sheet);
                            continue;
                        }

                        sheet.Activate();
                        SolidEdgeDraft.SheetWindow sw = (SolidEdgeDraft.SheetWindow)SolidEdgeConnector.App.ActiveWindow;

                        // 1. [추가] 이미지를 저장하기 전에 도면을 화면 중앙에 꽉 차게 정렬합니다.
                        sw.Fit();

                        string saveName = $"{nameWithoutExt}_{sheet.Name}.{type}";
                        string fullPath = Path.Combine(output, saveName);
                        int w_size = config.ImageSizeWidth;
                        int h_size = config.ImageSizeHeight;

                        sw.SaveAsImage(fullPath, w_size, h_size, Type.Missing, 24,
                                       SolidEdgeFramework.SeImageQualityType.seImageQualityHigh, true);

                        sheetCount++;
                        Marshal.ReleaseComObject(sw);
                        Marshal.ReleaseComObject(sheet);
                    }

                    Invoke((Action)(() => UpdateRow(row, true, sheetCount)));
                    WriteLog($"[성공] {fileName} ({sheetCount} sheets)");

                    AddLogToDataTable(row.Cells["colNo"].Value.ToString(), row.Cells["colStatus"].Value.ToString(), row.Cells["colFileName"].Value.ToString(), row.Cells["colResult"].Value.ToString(), "");
                    _lastRunSuccessCount++;
                }
                catch (Exception ex)
                {
                    WriteLog($"[실패] {fileName}: {ex.Message}");
                    Invoke((Action)(() => UpdateRow(row, false, 0)));
                    
                    AddLogToDataTable(row.Cells["colNo"].Value.ToString(), row.Cells["colStatus"].Value.ToString(), row.Cells["colFileName"].Value.ToString(), row.Cells["colResult"].Value.ToString(), ex.Message);
                    _lastRunFailCount++;
                }
                finally
                {                                       

                    if (draftDoc != null) { draftDoc.Close(false); Marshal.ReleaseComObject(draftDoc); }
                    if (seDoc != null) { seDoc.Close(false); Marshal.ReleaseComObject(seDoc); }

                    Invoke((Action)(() => UpdateProgress(no, total, startTime)));

                    // --- 대량 작업 관리 로직 ---
                    if (no % config.ReleaseMemoryInterval == 0) ReleaseMemory(); //메모리 정리
                    if (no % config.RestartSolidEdgeInterval == 0 && no < total) RestartSolidEdge(); //SE 재시작

                    no++;
                }
            }
            ExportDataTableToExcel();
        }
        #endregion

        #region [유틸리티 및 UI 업데이트]
        private void progressInit()
        {
            progressBar1.Value = 0;
            lblStatus.Text = "대기 중...";
        }

        private void imageTypeInit()
        {
            cboImageType.Items.Clear();
            cboImageType.Items.AddRange(new string[] { "tif", "png", "jpg" });
            cboImageType.SelectedItem = "tif";
            cboImageType.DropDownStyle = ComboBoxStyle.DropDownList;
        }

        private void UpdateProgress(int current, int total, DateTime startTime)
        {
            if (current == 0) return;

            progressBar1.Value = current;

            // 1. 현재 파일까지 걸린 전체 시간
            TimeSpan elapsed = DateTime.Now - startTime;
            double currentFileSeconds = elapsed.TotalSeconds / current;

            // 2. 이동 평균 계산 (최근 속도에 20% 가중치)
            // 초기에 -1일 때는 현재 속도로 초기화
            if (_movingAverageSeconds < 0)
                _movingAverageSeconds = currentFileSeconds;
            else
                _movingAverageSeconds = (_movingAverageSeconds * 0.8) + (currentFileSeconds * 0.2);

            // 3. 남은 시간 계산 (이동 평균 적용)
            double remainingSeconds = _movingAverageSeconds * (total - current);
            TimeSpan eta = TimeSpan.FromSeconds(remainingSeconds);

            // 4. 시간 표시 포맷 (1시간이 넘을 경우 대비)
            string etaStr = eta.TotalHours >= 1
                ? $"{(int)eta.TotalHours:00}:{eta.Minutes:00}:{eta.Seconds:00}"
                : $"{eta.Minutes:00}:{eta.Seconds:00}";

            lblStatus.Text = $"진행: {current}/{total} | 예상 남은 시간: {etaStr}";
        }

        private DataGridViewRow AddProcessingRow(int no, string fileName)
        {
            dataGridView1.Rows.Insert(0, no, "⏳", fileName, "처리 중");
            return dataGridView1.Rows[0];
        }

        private void UpdateRow(DataGridViewRow row, bool success, int count)
        {
            row.Cells["colStatus"].Value = success ? "✔" : "❌";
            row.Cells["colResult"].Value = success ? "완료" : "실패";
            row.DefaultCellStyle.BackColor = success ? Color.AliceBlue : Color.MistyRose;
        }

        private void SetUIState(bool isProcessing)
        {
            // UI 스레드에서 실행 보장
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => SetUIState(isProcessing)));
                return;
            }

            isRunning = isProcessing;

            // 실행 버튼 제어
            btnExecute.Enabled = !isProcessing;
            btnExecute.BackColor = isProcessing ? Color.DarkGray : Color.LightSteelBlue; // 실행 중엔 회색, 아니면 하늘색(예시)
            btnExecute.Text = isProcessing ? "변환 중..." : "변환 실행";

            // 재시도 버튼 제어
            //btnRetryFailed.Enabled = !isProcessing;
            //btnRetryFailed.BackColor = isProcessing ? Color.DarkGray : Color.Orange;

            // 중단 버튼은 실행 중에만 활성화
            btnCancel.Enabled = isProcessing;
            btnCancel.BackColor = isProcessing ? Color.IndianRed : Color.LightGray;

            // 마우스 커서 상태
            Cursor targetCursor = isProcessing ? Cursors.WaitCursor : Cursors.Default;

            this.Cursor = targetCursor;           // 폼 전체 커서
            dataGridView1.Cursor = targetCursor;  // 그리드뷰 전용 커서 (명시적 지정)            
        }

        private void WriteLog(string message)
        {
            try
            {
                if (!Directory.Exists(LogDirectory)) Directory.CreateDirectory(LogDirectory);
                string logFile = Path.Combine(LogDirectory, $"Log_{DateTime.Now:yyyyMMdd}.txt");
                File.AppendAllText(logFile, $"[{DateTime.Now:HH:mm:ss}] {message}\r\n");
            }
            catch { }
        }

        private void AddLogToDataTable(string no, string status, string fileName, string result, string exMessage)
        {
            // DataTable은 스레드 안전하게 처리하거나, 여기서는 List에 담았다가 나중에 변환해도 됨
            // 루프 내부에서 lock 없이 처리하려면 아래와 같이 작성
            DataRow dr = dtLog.NewRow();
            dr["No"] = no;
            dr["Status"] = status;
            dr["FileName"] = fileName;
            dr["Result"] = result;
            dr["Message"] = exMessage;
            dr["Timestamp"] = DateTime.Now;
            dtLog.Rows.Add(dr);
        }

        private void ExportDataTableToExcel()
        {
            try
            {
                if (!Directory.Exists(LogDirectory)) Directory.CreateDirectory(LogDirectory);
                string filePath = Path.Combine(LogDirectory, $"Result_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

                using (var wb = new XLWorkbook())
                {
                    // DataTable을 한 번에 시트로 변환 (매우 빠름)
                    var ws = wb.Worksheets.Add(dtLog, "ProcessLog");
                    ws.Columns().AdjustToContents();
                    wb.SaveAs(filePath);
                }
                WriteLog($"[System] Excel Export 완료: {filePath}");
            }
            catch (Exception ex)
            {
                WriteLog($"[Excel Error] {ex.Message}");
            }
        }

        private void InitLogDataTable()
        {
            dtLog = new DataTable("ProcessLog");
            dtLog.Columns.Add("No", typeof(int));
            dtLog.Columns.Add("Status", typeof(string));
            dtLog.Columns.Add("FileName", typeof(string));
            dtLog.Columns.Add("Result", typeof(string));
            dtLog.Columns.Add("Message", typeof(string));
            dtLog.Columns.Add("Timestamp", typeof(DateTime));
        }

        private void btnSourceBrowse_Click(object sender, EventArgs e) { using (var dlg = new FolderBrowserDialog()) if (dlg.ShowDialog() == DialogResult.OK) txtSourcePath.Text = dlg.SelectedPath; }
        private void btnOutputBrowse_Click(object sender, EventArgs e) { using (var dlg = new FolderBrowserDialog()) if (dlg.ShowDialog() == DialogResult.OK) txtOutputPath.Text = dlg.SelectedPath; }
        private void btnResultExcelBrowse_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel (*.xlsx)|*.xlsx|모든 파일 (*.*)|*.*";
                dlg.Title = "실패 재시도용 결과 엑셀 선택";
                if (dlg.ShowDialog() == DialogResult.OK)
                    txtResultExcelPath.Text = dlg.FileName;
            }
        }
        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(this, "진행 중인 변환 작업을 중단하시겠습니까?", "작업 취소",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
                return;
            cts?.Cancel();
        }

        private void LoadSavedPathsToUi()
        {
            try
            {
                var s = Settings.Default;
                if (!string.IsNullOrWhiteSpace(s.LastSourcePath))
                    txtSourcePath.Text = s.LastSourcePath;
                if (!string.IsNullOrWhiteSpace(s.LastOutputPath))
                    txtOutputPath.Text = s.LastOutputPath;
            }
            catch { }
        }

        private void SavePathsFromUi()
        {
            try
            {
                var s = Settings.Default;
                s.LastSourcePath = txtSourcePath.Text.Trim();
                s.LastOutputPath = txtOutputPath.Text.Trim();
                s.Save();
            }
            catch { }
        }

        private void UpdateRecentRowsHintLabel()
        {
            if (lblRecentRowsHint == null)
                return;
            int n = config != null && config.GridClearInterval > 0 ? config.GridClearInterval : 100;
            lblRecentRowsHint.Text =
                $"최근 약 {n}건만 표시됩니다. 목록은 주기적으로 비워지며, 전체 기록은 완료 후 Log 폴더의 결과 Excel에서 확인할 수 있습니다.";
        }
        #endregion

        #region [피로도 및 메모리 관리 함수]
        // 1. 메모리 누수 방지 (가비지 컬렉션 강제 수행)
        private void ReleaseMemory()
        {
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                WriteLog("[시스템] 가비지 컬렉션 및 메모리 정리 완료");
            }
            catch { }
        }

        // 2. 프로세스 피로도 관리 (Solid Edge 재시작)
        private void RestartSolidEdge()
        {
            try
            {
                WriteLog("[시스템] Solid Edge 재시작");

                //정상 종료 시도
                if (SolidEdgeConnector.App != null)
                {
                    try { SolidEdgeConnector.App.Quit(); } catch { }
                    Marshal.ReleaseComObject(SolidEdgeConnector.App);
                    SolidEdgeConnector.App = null;
                }

                // 3초 대기 후에도 Solid Edge 호스트(Edge.exe)가 남아 있으면 해당 프로세스만 강제 종료
                int waitCount = 0;
                while (waitCount < 3)
                {
                    Process[] procs = GetSolidEdgeHostProcesses();
                    if (procs.Length == 0) break;

                    Thread.Sleep(1000);
                    waitCount++;

                    if (waitCount >= 3)
                    {
                        foreach (var p in procs)
                        {
                            try
                            {
                                p.Kill();
                                p.WaitForExit(1000);
                                WriteLog("[시스템] 응답 없는 Solid Edge(Edge.exe) 프로세스 강제 종료 완료.");
                            }
                            catch { }
                            finally
                            {
                                try { p.Dispose(); } catch { }
                            }
                        }
                    }
                    else
                    {
                        foreach (var p in procs)
                        {
                            try { p.Dispose(); } catch { }
                        }
                    }
                }

                //OS가 프로세스 테이블을 정리할 아주 짧은 시간
                Thread.Sleep(1000);

                //새 세션 기동
                SolidEdgeConnector.App = SolidEdgeConnector.GetInstance();
                SolidEdgeConnector.App.DisplayAlerts = false;
                SolidEdgeConnector.App.ScreenUpdating = false;
                SolidEdgeConnector.App.Visible = false;

                WriteLog("[시스템] Solid Edge 새 인스턴스 준비 완료.");
            }
            catch (Exception ex)
            {
                WriteLog($"[오류] 재시작 실패: {ex.Message}");
                // 여기서 실패하더라도 다음 파일 루프에서 GetInstance()가 재시도
                SolidEdgeConnector.App = null;
            }
        }

        /// <summary>
        /// Solid Edge 메인 호스트는 설치 경로의 Edge.exe입니다. 실행 파일 경로에 Solid Edge 설치 디렉터리가 포함된 경우만 반환합니다.
        /// (Chromium 기반 Microsoft Edge는 msedge.exe이며 ProcessName이 다릅니다.)
        /// </summary>
        private static Process[] GetSolidEdgeHostProcesses()
        {
            var list = new List<Process>();
            foreach (var p in Process.GetProcessesByName("Edge"))
            {
                try
                {
                    string path = p.MainModule?.FileName;
                    if (string.IsNullOrEmpty(path))
                    {
                        p.Dispose();
                        continue;
                    }

                    bool isSolidEdgeHost =
                        path.IndexOf(@"\Siemens\Solid Edge", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        path.IndexOf(@"\Solid Edge\", StringComparison.OrdinalIgnoreCase) >= 0;

                    if (isSolidEdgeHost)
                        list.Add(p);
                    else
                        p.Dispose();
                }
                catch
                {
                    try { p.Dispose(); } catch { }
                }
            }

            return list.ToArray();
        }

        private static bool WaitForStaTaskWithTimeout(Task task, int timeoutMilliseconds)
        {
            Stopwatch sw = Stopwatch.StartNew();
            while (!task.IsCompleted)
            {
                Application.DoEvents();
                Thread.Sleep(5);

                if (sw.ElapsedMilliseconds > timeoutMilliseconds)
                    return false; // 시간 초과
            }
            task.GetAwaiter().GetResult();
            return true; // 정상 완료
        }

        /// <summary>
        /// 미처 닫히지 않고 남아있는 Solid Edge 호스트 프로세스를 강제 종료합니다.
        /// </summary>
        private void KillSolidEdgeHostProcessesGracefully()
        {
            try
            {
                Process[] procs = GetSolidEdgeHostProcesses();
                foreach (var p in procs)
                {
                    try
                    {
                        Trace.WriteLine($"[Kill] Solid Edge 호스트 프로세스 종료 시도: PID={p.Id}, Path={p.MainModule?.FileName}");
                        p.Kill();
                        p.Dispose();
                    }
                    catch { }
                }
            }
            catch { }
        }
        #endregion
    }
}