using System;
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

namespace KSM_SolidEdge
{
    public partial class mainDlg : Form
    {
        private bool isRunning = false;
        private CancellationTokenSource cts;

        private DataTable dtLog;
        private double _movingAverageSeconds = -1; // 이동 평균 저장용

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
            this.Load += mainDlg_Load;
            this.FormClosing += mainDlg_FormClosing;
        }

        #region [초기화 및 종료]
        private async void mainDlg_Load(object sender, EventArgs e)
        {
            progressInit();
            imageTypeInit();
            InitLogDataTable();
            btnCancel.Enabled = false;

            lblStatus.Text = "시스템 준비 중 (Solid Edge 연결)...";

            await Task.Run(() => {
                try
                {
                    SolidEdgeConnector.GetInstance();
                    
                    if (SolidEdgeConnector.App != null)
                    {
                        Debug.WriteLine("Solid Edge 연결");
                        SolidEdgeConnector.App.DisplayAlerts = false;
                        SolidEdgeConnector.App.ScreenUpdating = true;
                        SolidEdgeConnector.App.Visible = false;
                    }
                }
                catch (Exception ex)
                {
                    WriteLog($"[준비 실패] SE 연결 실패: {ex.Message}");
                }
            });

            lblStatus.Text = "대기 중...";
        }

        private void mainDlg_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (SolidEdgeConnector.App != null)
                {
                    // 1. 열려 있는 문서가 있다면 모두 저장하지 않고 닫기 (안전장치)
                    if (SolidEdgeConnector.App.Documents.Count > 0)
                    {
                        SolidEdgeConnector.App.Documents.Close();
                    }

                    // 2. 솔리드 엣지 자체를 종료 (화면에 보이게 하지 않음)
                    SolidEdgeConnector.App.DisplayAlerts = true;
                    SolidEdgeConnector.App.Quit();

                    // 3. COM 객체 해제
                    Marshal.ReleaseComObject(SolidEdgeConnector.App);
                    SolidEdgeConnector.App = null;
                }
            }
            catch { }
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

            // 작업 시작 전 SE 상태 체크
            if (SolidEdgeConnector.App == null)
            {
                SolidEdgeConnector.App = SolidEdgeConnector.GetInstance();
            }

            // 최적화 설정 적용
            SolidEdgeConnector.App.DisplayAlerts = false;
            SolidEdgeConnector.App.ScreenUpdating = false;
            SolidEdgeConnector.App.Visible = false;

            SetUIState(true);

            // 실행 시점에 최신 설정 반영
            config = AppConfig.Load();

            cts = new CancellationTokenSource();
            string source = txtSourcePath.Text;
            string output = txtOutputPath.Text;
            string type = cboImageType.Text.ToLower();
            dtLog.Clear(); // 새 작업 시작 시 로그 초기화
            _movingAverageSeconds = -1;

            try
            {
                // --- [추가] 절전 모드 방지 시작 ---
                SleepModeController.PreventSleep();
                WriteLog("[시스템] 절전 모드 진입 방지 활성화");

                await Task.Run(() => RunProcess(source, output, type, cts.Token));
                // 알림창을 띄우고 확인을 누르면 폴더 열기
                DialogResult result = MessageBox.Show(this, "모든 변환 작업이 완료되었습니다.\n출력 폴더를 여시겠습니까?", "완료"
                    , MessageBoxButtons.YesNo, MessageBoxIcon.Information);

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

                if (SolidEdgeConnector.App != null)
                {
                    SolidEdgeConnector.App.ScreenUpdating = true;
                    SolidEdgeConnector.App.DisplayAlerts = true;
                }
                progressInit();
            }
        }

        private void btnReset_Click(object sender, EventArgs e) => dataGridView1.Rows.Clear();

        private void RunProcess(string source, string output, string type, CancellationToken token)
        {
            if (!Directory.Exists(output)) Directory.CreateDirectory(output);

            string[] files = Directory.GetFiles(source, "*.*")
                .Where(f => f.EndsWith(".asm", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".par", StringComparison.OrdinalIgnoreCase))
                .OrderBy(f => f) // 이름순 정렬 추가
                .ToArray();

            if (files.Length == 0) return;

            int total = files.Length;
            Invoke((Action)(() => {
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

                        string saveName = $"{nameWithoutExt}_{sheet.Name}.{type}";
                        string fullPath = Path.Combine(output, saveName);
                        int w_size = config.ImageSizeWidth;
                        int h_size = config.ImageSizeHeight;

                        sw.SaveAsImage(fullPath, w_size, h_size, Type.Missing, 24,
                                       SolidEdgeFramework.SeImageQualityType.seImageQualityHigh, false);

                        sheetCount++;
                        Marshal.ReleaseComObject(sw);
                        Marshal.ReleaseComObject(sheet);
                    }

                    Invoke((Action)(() => UpdateRow(row, true, sheetCount)));
                    WriteLog($"[성공] {fileName} ({sheetCount} sheets)");

                    AddLogToDataTable(row.Cells["colNo"].Value.ToString(), row.Cells["colStatus"].Value.ToString(), row.Cells["colFileName"].Value.ToString(), row.Cells["colResult"].Value.ToString(), "");
                }
                catch (Exception ex)
                {
                    WriteLog($"[실패] {fileName}: {ex.Message}");
                    Invoke((Action)(() => UpdateRow(row, false, 0)));
                    
                    AddLogToDataTable(row.Cells["colNo"].Value.ToString(), row.Cells["colStatus"].Value.ToString(), row.Cells["colFileName"].Value.ToString(), row.Cells["colResult"].Value.ToString(), ex.Message);
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
        private void btnCancel_Click(object sender, EventArgs e) { cts?.Cancel(); }
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

                //3초만 기다려보고 안 죽으면 즉시 Kill
                int waitCount = 0;
                while (waitCount < 3)
                {
                    var procs = Process.GetProcessesByName("Edge");
                    if (procs.Length == 0) break; // 스스로 잘 죽었으면 루프 탈출

                    Thread.Sleep(1000);
                    waitCount++;

                    // 3초가 됐는데도 살아있다면 강제 종료 단행
                    if (waitCount >= 3)
                    {
                        foreach (var p in procs)
                        {
                            try
                            {
                                p.Kill();
                                p.WaitForExit(1000); // 완전히 사라질 때까지 1초만 더 대기
                                WriteLog("[시스템] 응답 없는 프로세스 강제 종료 완료.");
                            }
                            catch { }
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
        #endregion
    }
}