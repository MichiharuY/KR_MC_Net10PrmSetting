using Microsoft.VisualBasic;
using System.Collections;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace KR_Net10PrmSetting
{
    public partial class FormMain : Form
    {

        // 変数
        public NET10Address? m_NET10Addr = null; // 収録アドレス管理クラス
        public NET10Control? m_NET10Ctrl = null; // NET10通信管理クラス

        // NET10アドレス情報(Excelファイル読出用)
        public struct ExcelInfo
        {
            public List<int> Row;       // Excelシートの行番号
            public List<string> Item;   // 項目名
            public List<short> Addr;    // アドレス
            public List<short> Data;    // 書込データ
        }
        private string m_ExcelPath = Properties.Settings.Default.ExcelPath; // 設定ファイルから取得
        private ExcelInfo[] m_Excel = new ExcelInfo[2]; // Excel情報
        private Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        private Microsoft.Office.Interop.Excel.Workbook? xlWorkbook = null;
        private Microsoft.Office.Interop.Excel.Worksheet? xlWorksheet = null;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public FormMain()
        {
            InitializeComponent();

            m_Excel[0] = new ExcelInfo();
            m_Excel[1] = new ExcelInfo();
        }

        /// <summary>
        /// デストラクタ
        /// </summary>
        ~FormMain()
        {
            if (xlWorksheet != null)
            {
                Marshal.ReleaseComObject(xlWorksheet);
            }

            if (xlWorkbook != null)
            {
                // Excel Close and Release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
            }

            // Excel Quit and Release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        /// <summary>
        /// フォームロード
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormMain_Load(object sender, EventArgs e)
        {
            this.StartPosition = FormStartPosition.CenterScreen;

            ////////////////////////////////////////
            // NET10通信の初期化
            m_NET10Ctrl = new NET10Control();
            if (m_NET10Ctrl.StartComm() == false)
            {
                MessageBox.Show("NET10通信の初期化に失敗しました。", "NET10通信初期化エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// フォームアクティブ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormMain_Activated(object sender, EventArgs e)
        {
            // Excelのパラメータ一覧を取得しリスト作成
            MakeParamCombo();
        }

        /// <summary>
        /// パラメータのコンボリストを作成する
        /// </summary>
        private void MakeParamCombo()
        {
            lblStatus.Text = "Excelファイル読込中...";

            // パラメータ名、アドレス、設定値、Excelのセル位置
            ExcelRead(m_ExcelPath, 1, 2, 2, 144, 5);

            cmbItemName1.Items.Clear();
            for (int i = 0; i < m_Excel[0].Item.Count; i++)
            {
                cmbItemName1.Items.Add(m_Excel[0].Item[i]);
            }

            ExcelRead(m_ExcelPath, 2, 2, 2, 145, 6);

            cmbItemName2.Items.Clear();
            for (int i = 0; i < m_Excel[1].Item.Count; i++)
            {
                cmbItemName2.Items.Add(m_Excel[1].Item[i]);
            }

            cmbItemName1.SelectedIndex = 0;
            cmbItemName2.SelectedIndex = 0;

            lblStatus.Text = "";
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (m_NET10Ctrl != null)
            {
                if (m_NET10Ctrl.EndComm() == false)
                {
                    MessageBox.Show("NET10の通信停止に失敗しました。", "NET10通信停止エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// エクセルファイルの指定したシートを2次元配列に読み込む.
        /// </summary>
        /// <param name="filePath">エクセルファイルのパス</param>
        /// <param name="sheetIndex">シートの番号 (1, 2, 3, ...)</param>
        /// <param name="startRow">最初の行 (>= 1)</param>
        /// <param name="startColmn">最初の列 (>= 1)</param>
        /// <param name="lastRow">最後の行</param>
        /// <param name="lastColmn">最後の列</param>
        /// <returns>シート情報を格納した2次元文字配列. ただしファイル読み込みに失敗したときには null.</returns>
        public ArrayList ExcelRead(string filePath, int sheetIndex,
                              int startRow, int startColmn,
                              int lastRow, int lastColmn)
        {
            var arrOut = new ArrayList();

            m_Excel[sheetIndex - 1].Row = new List<int>();
            m_Excel[sheetIndex - 1].Item = new List<string>();
            m_Excel[sheetIndex - 1].Addr = new List<short>();
            m_Excel[sheetIndex - 1].Data = new List<short>();

            // ワークブックを開く
            if (!ExcelOpen(filePath)) { return arrOut; }

            if (xlWorkbook != null)
            {
                xlWorksheet = xlWorkbook.Sheets[sheetIndex];
                xlWorksheet.Select();


                for (int r = startRow; r <= lastRow; r++)
                {
                    // 一行読み込む
                    var row = new ArrayList();
                    for (int c = startColmn; c <= lastColmn; c++)
                    {
                        var cell = xlWorksheet.Cells[r, c];

                        if (cell == null || cell.Value == null)
                        {
                            row.Add("");
                        }
                        else
                        {
                            row.Add(cell.Value);
                        }
                    }

                    string item = "";
                    short addr = 0;
                    if (row[0] != null)
                    {
                        item = row[0].ToString();
                    }
                    if (row[1] != null)
                    {
                        if (row[1].ToString() != "")
                        {
                            addr = Convert.ToInt16(row[1].ToString().Substring(1, 4), 16);
                        }
                    }

                    short data = 0;
                    if (sheetIndex == 1)
                    {
                        // 制御パラメータ
                        if (row[3] != null)
                        {
                            if (row[3].ToString() != "")
                            {
                                data = Convert.ToInt16(row[3].ToString(), 10);
                            }
                        }
                    }
                    else
                    {
                        // 警報パラメータ
                        if (row[4] != null)
                        {
                            if (row[4].ToString() != "")
                            {
                                data = Convert.ToInt16(row[4].ToString(), 10);
                            }
                        }

                    }

                    //リストに格納
                    m_Excel[sheetIndex - 1].Row.Add(r);
                    m_Excel[sheetIndex - 1].Item.Add(item);
                    m_Excel[sheetIndex - 1].Addr.Add(addr);
                    m_Excel[sheetIndex - 1].Data.Add(data);

                    arrOut.Add(row);
                }

                // ワークシートを閉じる
                Marshal.ReleaseComObject(xlWorksheet);
                xlWorksheet = null;
            }

            // ワークブックとエクセルのプロセスを閉じる
            ExcelClose();

            return arrOut;
        }

        /// <summary>
        /// エクセルファイルの指定したシートのセルにデータを書き込む
        /// </summary>
        /// <param name="filePath">エクセルファイルのパス</param>
        /// <param name="sheetIndex">シートの番号 (1, 2, 3, ...)</param>
        /// <param name="row">行 (>= 1)</param>
        /// <param name="colmn">列 (>= 1)</param>
        /// <param name="data">書込データ</param>
        /// <returns>実行結果</returns>
        public bool ExcelWrite(string filePath, int sheetIndex, int row, int colmn, string data)
        {
            lblStatus.Text = "Excelファイルへ保存中...";

            // ワークブックを開く
            if (!ExcelOpen(filePath)) { return false; }

            if (xlWorkbook != null)
            {
                xlWorksheet = xlWorkbook.Sheets[sheetIndex];
                xlWorksheet.Select();

                xlWorksheet.Cells[row, colmn] = data;

                // 保存する
                xlWorkbook.Save();

                // ワークシートを閉じる
                Marshal.ReleaseComObject(xlWorksheet);
                xlWorksheet = null;
            }

            // ワークブックとエクセルのプロセスを閉じる
            ExcelClose();

            lblStatus.Text = "";

            return true;
        }

        /// <summary>
        /// 指定されたパスのExcel Workbookを開く
        /// </summary>
        /// <param name="filePath">Excel Workbookのパス(相対パスでも絶対パスでもOK)</param>
        /// <returns>Excel Workbookのオープンに成功したら true. それ以外 false.</returns>
        protected bool ExcelOpen(string filePath)
        {
            if (!System.IO.File.Exists(filePath))
            {
                return false;
            }

            try
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlApp.Visible = false;

                // filePath が相対パスのとき例外が発生するので fullPath に変換
                string fullPath = System.IO.Path.GetFullPath(filePath);
                xlWorkbook = xlApp.Workbooks.Open(fullPath);
            }
            catch
            {
                ExcelClose();
                return false;
            }

            return true;
        }

        /// <summary>
        /// 開いているWorkbookとExcelのプロセスを閉じる.
        /// </summary>
        protected void ExcelClose()
        {
            if (xlWorkbook != null)
            {
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                xlWorkbook = null;
            }

            if (xlApp != null)
            {
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                xlApp = null;
            }
        }

        /// <summary>
        /// パラメータ変更（制御パラメータ）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbItemName1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 現在のNet10のアドレスを取得
            txtAddress1.Text = "W" + Convert.ToString(m_Excel[0].Addr[cmbItemName1.SelectedIndex], 16).ToUpper().PadLeft(4, '0');
            txtSetValue1.Text = m_Excel[0].Data[cmbItemName1.SelectedIndex].ToString();


            if (m_NET10Ctrl != null)
            {
                // アドレスの初期化
                m_NET10Addr = new NET10Address();
                m_NET10Addr.m_SendAddrInfo.Addr = new List<short>();
                m_NET10Addr.m_SendAddrInfo.Data = new List<short>();
                short addr = Convert.ToInt16(txtAddress1.Text.ToString().Substring(1), 16);
                m_NET10Addr.m_SendAddrInfo.Addr.Add(addr);
                m_NET10Addr.m_SendAddrInfo.Data.Add(0);

                // Net10から受信
                short ret = m_NET10Ctrl.DataReceive(m_NET10Addr.m_SendAddrInfo);
                txtNet10Value1.Text = ret.ToString();
            }
        }

        /// <summary>
        /// パラメータ変更（警報パラメータ）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbItemName2_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 現在のNet10のアドレスを取得
            txtAddress2.Text = "W" + Convert.ToString(m_Excel[1].Addr[cmbItemName2.SelectedIndex], 16).ToUpper().PadLeft(4, '0');
            txtSetValue2.Text = m_Excel[1].Data[cmbItemName2.SelectedIndex].ToString();

            if (m_NET10Ctrl != null)
            {
                // アドレスの初期化
                m_NET10Addr = new NET10Address();
                m_NET10Addr.m_SendAddrInfo.Addr = new List<short>();
                m_NET10Addr.m_SendAddrInfo.Data = new List<short>();
                short addr = Convert.ToInt16(txtAddress2.Text.ToString().Substring(1), 16);
                m_NET10Addr.m_SendAddrInfo.Addr.Add(addr);
                m_NET10Addr.m_SendAddrInfo.Data.Add(0);

                // Net10から受信
                short ret = m_NET10Ctrl.DataReceive(m_NET10Addr.m_SendAddrInfo);
                txtNet10Value2.Text = ret.ToString();
            }
        }

        /// <summary>
        /// 制御パラメータの入力値をNet10の指定アドレスに送信
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSet1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("NET10へデータ設定してもよろしいですか？", "データ設定確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            if (txtAddress1.Text == "")
            {
                MessageBox.Show("書込アドレスがExcelに登録されていません。", "アドレス登録エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtAddress1.Focus();
                return;
            }

            if (txtSetValue1.Text == "")
            {
                MessageBox.Show("書込データが入力されていません。", "データ入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtSetValue1.Focus();
                return;
            }

            if (m_NET10Ctrl != null)
            {
                // アドレスの初期化
                m_NET10Addr = new NET10Address();
                m_NET10Addr.m_SendAddrInfo.Addr = new List<short>();
                m_NET10Addr.m_SendAddrInfo.Data = new List<short>();
                short addr = Convert.ToInt16(txtAddress1.Text.ToString().Substring(1), 16);
                short data = Convert.ToInt16(txtSetValue1.Text.ToString(), 10);
                m_NET10Addr.m_SendAddrInfo.Addr.Add(addr);
                m_NET10Addr.m_SendAddrInfo.Data.Add(data);

                // Net10へ送信
                if (m_NET10Ctrl.DataSend(m_NET10Addr.m_SendAddrInfo) == 0)
                {
                    ExcelWrite(m_ExcelPath, 1, m_Excel[0].Row[cmbItemName1.SelectedIndex], 5, txtSetValue1.Text);

                    if (m_NET10Ctrl != null)
                    {
                        // Net10から受信
                        short ret = m_NET10Ctrl.DataReceive(m_NET10Addr.m_SendAddrInfo);
                        txtNet10Value1.Text = ret.ToString();
                    }

                    MessageBox.Show("正常に送信しました。", "データ送信", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("送信に失敗しました。", "データ送信エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// 警報パラメータの入力値をNet10の指定アドレスに送信
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSet2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("NET10へデータ設定してもよろしいですか？", "データ設定確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            if (txtAddress2.Text == "")
            {
                MessageBox.Show("書込アドレスがExcelに登録されていません。", "アドレス登録エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtAddress2.Focus();
                return;
            }

            if (txtSetValue2.Text == "")
            {
                MessageBox.Show("書込データが入力されていません。", "データ入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtSetValue2.Focus();
                return;
            }

            if (m_NET10Ctrl != null)
            {
                // アドレスの初期化
                m_NET10Addr = new NET10Address();
                m_NET10Addr.m_SendAddrInfo.Addr = new List<short>();
                m_NET10Addr.m_SendAddrInfo.Data = new List<short>();
                short addr = Convert.ToInt16(txtAddress2.Text.ToString().Substring(1), 16);
                short data = Convert.ToInt16(txtSetValue2.Text.ToString(), 10);
                m_NET10Addr.m_SendAddrInfo.Addr.Add(addr);
                m_NET10Addr.m_SendAddrInfo.Data.Add(data);

                // Net10へ送信
                if (m_NET10Ctrl.DataSend(m_NET10Addr.m_SendAddrInfo) == 0)
                {
                    ExcelWrite(m_ExcelPath, 2, m_Excel[1].Row[cmbItemName2.SelectedIndex], 6, txtSetValue2.Text);

                    if (m_NET10Ctrl != null)
                    {
                        // Net10から受信
                        short ret = m_NET10Ctrl.DataReceive(m_NET10Addr.m_SendAddrInfo);
                        txtNet10Value2.Text = ret.ToString();
                    }

                    MessageBox.Show("正常に送信しました。", "データ送信", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("送信に失敗しました。", "データ送信エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// EXCELデータ再読み込み
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExcelReload_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Excelの設定データを再読み込みしてもよろしいですか？", "データ再読込確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                m_Excel = new ExcelInfo[2]; // Excel情報

                // Excelのパラメータ一覧を取得しリスト作成
                MakeParamCombo();
            }

        }
    }
}