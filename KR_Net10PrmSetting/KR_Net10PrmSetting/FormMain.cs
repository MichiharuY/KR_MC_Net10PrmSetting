using Microsoft.VisualBasic;
using System.Collections;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace KR_Net10PrmSetting
{
    public partial class FormMain : Form
    {

        // �ϐ�
        public NET10Address? m_NET10Addr = null; // ���^�A�h���X�Ǘ��N���X
        public NET10Control? m_NET10Ctrl = null; // NET10�ʐM�Ǘ��N���X

        // NET10�A�h���X���(Excel�t�@�C���Ǐo�p)
        public struct ExcelInfo
        {
            public List<int> Row;       // Excel�V�[�g�̍s�ԍ�
            public List<string> Item;   // ���ږ�
            public List<short> Addr;    // �A�h���X
            public List<short> Data;    // �����f�[�^
        }
        private string m_ExcelPath = Properties.Settings.Default.ExcelPath; // �ݒ�t�@�C������擾
        private ExcelInfo[] m_Excel = new ExcelInfo[2]; // Excel���
        private Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        private Microsoft.Office.Interop.Excel.Workbook? xlWorkbook = null;
        private Microsoft.Office.Interop.Excel.Worksheet? xlWorksheet = null;

        /// <summary>
        /// �R���X�g���N�^
        /// </summary>
        public FormMain()
        {
            InitializeComponent();

            m_Excel[0] = new ExcelInfo();
            m_Excel[1] = new ExcelInfo();
        }

        /// <summary>
        /// �f�X�g���N�^
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
        /// �t�H�[�����[�h
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormMain_Load(object sender, EventArgs e)
        {
            this.StartPosition = FormStartPosition.CenterScreen;

            ////////////////////////////////////////
            // NET10�ʐM�̏�����
            m_NET10Ctrl = new NET10Control();
            if (m_NET10Ctrl.StartComm() == false)
            {
                MessageBox.Show("NET10�ʐM�̏������Ɏ��s���܂����B", "NET10�ʐM�������G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// �t�H�[���A�N�e�B�u
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormMain_Activated(object sender, EventArgs e)
        {
            // Excel�̃p�����[�^�ꗗ���擾�����X�g�쐬
            MakeParamCombo();
        }

        /// <summary>
        /// �p�����[�^�̃R���{���X�g���쐬����
        /// </summary>
        private void MakeParamCombo()
        {
            lblStatus.Text = "Excel�t�@�C���Ǎ���...";

            // �p�����[�^���A�A�h���X�A�ݒ�l�AExcel�̃Z���ʒu
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
                    MessageBox.Show("NET10�̒ʐM��~�Ɏ��s���܂����B", "NET10�ʐM��~�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// �G�N�Z���t�@�C���̎w�肵���V�[�g��2�����z��ɓǂݍ���.
        /// </summary>
        /// <param name="filePath">�G�N�Z���t�@�C���̃p�X</param>
        /// <param name="sheetIndex">�V�[�g�̔ԍ� (1, 2, 3, ...)</param>
        /// <param name="startRow">�ŏ��̍s (>= 1)</param>
        /// <param name="startColmn">�ŏ��̗� (>= 1)</param>
        /// <param name="lastRow">�Ō�̍s</param>
        /// <param name="lastColmn">�Ō�̗�</param>
        /// <returns>�V�[�g�����i�[����2���������z��. �������t�@�C���ǂݍ��݂Ɏ��s�����Ƃ��ɂ� null.</returns>
        public ArrayList ExcelRead(string filePath, int sheetIndex,
                              int startRow, int startColmn,
                              int lastRow, int lastColmn)
        {
            var arrOut = new ArrayList();

            m_Excel[sheetIndex - 1].Row = new List<int>();
            m_Excel[sheetIndex - 1].Item = new List<string>();
            m_Excel[sheetIndex - 1].Addr = new List<short>();
            m_Excel[sheetIndex - 1].Data = new List<short>();

            // ���[�N�u�b�N���J��
            if (!ExcelOpen(filePath)) { return arrOut; }

            if (xlWorkbook != null)
            {
                xlWorksheet = xlWorkbook.Sheets[sheetIndex];
                xlWorksheet.Select();


                for (int r = startRow; r <= lastRow; r++)
                {
                    // ��s�ǂݍ���
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
                        // ����p�����[�^
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
                        // �x��p�����[�^
                        if (row[4] != null)
                        {
                            if (row[4].ToString() != "")
                            {
                                data = Convert.ToInt16(row[4].ToString(), 10);
                            }
                        }

                    }

                    //���X�g�Ɋi�[
                    m_Excel[sheetIndex - 1].Row.Add(r);
                    m_Excel[sheetIndex - 1].Item.Add(item);
                    m_Excel[sheetIndex - 1].Addr.Add(addr);
                    m_Excel[sheetIndex - 1].Data.Add(data);

                    arrOut.Add(row);
                }

                // ���[�N�V�[�g�����
                Marshal.ReleaseComObject(xlWorksheet);
                xlWorksheet = null;
            }

            // ���[�N�u�b�N�ƃG�N�Z���̃v���Z�X�����
            ExcelClose();

            return arrOut;
        }

        /// <summary>
        /// �G�N�Z���t�@�C���̎w�肵���V�[�g�̃Z���Ƀf�[�^����������
        /// </summary>
        /// <param name="filePath">�G�N�Z���t�@�C���̃p�X</param>
        /// <param name="sheetIndex">�V�[�g�̔ԍ� (1, 2, 3, ...)</param>
        /// <param name="row">�s (>= 1)</param>
        /// <param name="colmn">�� (>= 1)</param>
        /// <param name="data">�����f�[�^</param>
        /// <returns>���s����</returns>
        public bool ExcelWrite(string filePath, int sheetIndex, int row, int colmn, string data)
        {
            lblStatus.Text = "Excel�t�@�C���֕ۑ���...";

            // ���[�N�u�b�N���J��
            if (!ExcelOpen(filePath)) { return false; }

            if (xlWorkbook != null)
            {
                xlWorksheet = xlWorkbook.Sheets[sheetIndex];
                xlWorksheet.Select();

                xlWorksheet.Cells[row, colmn] = data;

                // �ۑ�����
                xlWorkbook.Save();

                // ���[�N�V�[�g�����
                Marshal.ReleaseComObject(xlWorksheet);
                xlWorksheet = null;
            }

            // ���[�N�u�b�N�ƃG�N�Z���̃v���Z�X�����
            ExcelClose();

            lblStatus.Text = "";

            return true;
        }

        /// <summary>
        /// �w�肳�ꂽ�p�X��Excel Workbook���J��
        /// </summary>
        /// <param name="filePath">Excel Workbook�̃p�X(���΃p�X�ł���΃p�X�ł�OK)</param>
        /// <returns>Excel Workbook�̃I�[�v���ɐ��������� true. ����ȊO false.</returns>
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

                // filePath �����΃p�X�̂Ƃ���O����������̂� fullPath �ɕϊ�
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
        /// �J���Ă���Workbook��Excel�̃v���Z�X�����.
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
        /// �p�����[�^�ύX�i����p�����[�^�j
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbItemName1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // ���݂�Net10�̃A�h���X���擾
            txtAddress1.Text = "W" + Convert.ToString(m_Excel[0].Addr[cmbItemName1.SelectedIndex], 16).ToUpper().PadLeft(4, '0');
            txtSetValue1.Text = m_Excel[0].Data[cmbItemName1.SelectedIndex].ToString();


            if (m_NET10Ctrl != null)
            {
                // �A�h���X�̏�����
                m_NET10Addr = new NET10Address();
                m_NET10Addr.m_SendAddrInfo.Addr = new List<short>();
                m_NET10Addr.m_SendAddrInfo.Data = new List<short>();
                short addr = Convert.ToInt16(txtAddress1.Text.ToString().Substring(1), 16);
                m_NET10Addr.m_SendAddrInfo.Addr.Add(addr);
                m_NET10Addr.m_SendAddrInfo.Data.Add(0);

                // Net10�����M
                short ret = m_NET10Ctrl.DataReceive(m_NET10Addr.m_SendAddrInfo);
                txtNet10Value1.Text = ret.ToString();
            }
        }

        /// <summary>
        /// �p�����[�^�ύX�i�x��p�����[�^�j
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbItemName2_SelectedIndexChanged(object sender, EventArgs e)
        {
            // ���݂�Net10�̃A�h���X���擾
            txtAddress2.Text = "W" + Convert.ToString(m_Excel[1].Addr[cmbItemName2.SelectedIndex], 16).ToUpper().PadLeft(4, '0');
            txtSetValue2.Text = m_Excel[1].Data[cmbItemName2.SelectedIndex].ToString();

            if (m_NET10Ctrl != null)
            {
                // �A�h���X�̏�����
                m_NET10Addr = new NET10Address();
                m_NET10Addr.m_SendAddrInfo.Addr = new List<short>();
                m_NET10Addr.m_SendAddrInfo.Data = new List<short>();
                short addr = Convert.ToInt16(txtAddress2.Text.ToString().Substring(1), 16);
                m_NET10Addr.m_SendAddrInfo.Addr.Add(addr);
                m_NET10Addr.m_SendAddrInfo.Data.Add(0);

                // Net10�����M
                short ret = m_NET10Ctrl.DataReceive(m_NET10Addr.m_SendAddrInfo);
                txtNet10Value2.Text = ret.ToString();
            }
        }

        /// <summary>
        /// ����p�����[�^�̓��͒l��Net10�̎w��A�h���X�ɑ��M
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSet1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("NET10�փf�[�^�ݒ肵�Ă���낵���ł����H", "�f�[�^�ݒ�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            if (txtAddress1.Text == "")
            {
                MessageBox.Show("�����A�h���X��Excel�ɓo�^����Ă��܂���B", "�A�h���X�o�^�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtAddress1.Focus();
                return;
            }

            if (txtSetValue1.Text == "")
            {
                MessageBox.Show("�����f�[�^�����͂���Ă��܂���B", "�f�[�^���̓G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtSetValue1.Focus();
                return;
            }

            if (m_NET10Ctrl != null)
            {
                // �A�h���X�̏�����
                m_NET10Addr = new NET10Address();
                m_NET10Addr.m_SendAddrInfo.Addr = new List<short>();
                m_NET10Addr.m_SendAddrInfo.Data = new List<short>();
                short addr = Convert.ToInt16(txtAddress1.Text.ToString().Substring(1), 16);
                short data = Convert.ToInt16(txtSetValue1.Text.ToString(), 10);
                m_NET10Addr.m_SendAddrInfo.Addr.Add(addr);
                m_NET10Addr.m_SendAddrInfo.Data.Add(data);

                // Net10�֑��M
                if (m_NET10Ctrl.DataSend(m_NET10Addr.m_SendAddrInfo) == 0)
                {
                    ExcelWrite(m_ExcelPath, 1, m_Excel[0].Row[cmbItemName1.SelectedIndex], 5, txtSetValue1.Text);

                    if (m_NET10Ctrl != null)
                    {
                        // Net10�����M
                        short ret = m_NET10Ctrl.DataReceive(m_NET10Addr.m_SendAddrInfo);
                        txtNet10Value1.Text = ret.ToString();
                    }

                    MessageBox.Show("����ɑ��M���܂����B", "�f�[�^���M", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("���M�Ɏ��s���܂����B", "�f�[�^���M�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// �x��p�����[�^�̓��͒l��Net10�̎w��A�h���X�ɑ��M
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSet2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("NET10�փf�[�^�ݒ肵�Ă���낵���ł����H", "�f�[�^�ݒ�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            if (txtAddress2.Text == "")
            {
                MessageBox.Show("�����A�h���X��Excel�ɓo�^����Ă��܂���B", "�A�h���X�o�^�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtAddress2.Focus();
                return;
            }

            if (txtSetValue2.Text == "")
            {
                MessageBox.Show("�����f�[�^�����͂���Ă��܂���B", "�f�[�^���̓G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtSetValue2.Focus();
                return;
            }

            if (m_NET10Ctrl != null)
            {
                // �A�h���X�̏�����
                m_NET10Addr = new NET10Address();
                m_NET10Addr.m_SendAddrInfo.Addr = new List<short>();
                m_NET10Addr.m_SendAddrInfo.Data = new List<short>();
                short addr = Convert.ToInt16(txtAddress2.Text.ToString().Substring(1), 16);
                short data = Convert.ToInt16(txtSetValue2.Text.ToString(), 10);
                m_NET10Addr.m_SendAddrInfo.Addr.Add(addr);
                m_NET10Addr.m_SendAddrInfo.Data.Add(data);

                // Net10�֑��M
                if (m_NET10Ctrl.DataSend(m_NET10Addr.m_SendAddrInfo) == 0)
                {
                    ExcelWrite(m_ExcelPath, 2, m_Excel[1].Row[cmbItemName2.SelectedIndex], 6, txtSetValue2.Text);

                    if (m_NET10Ctrl != null)
                    {
                        // Net10�����M
                        short ret = m_NET10Ctrl.DataReceive(m_NET10Addr.m_SendAddrInfo);
                        txtNet10Value2.Text = ret.ToString();
                    }

                    MessageBox.Show("����ɑ��M���܂����B", "�f�[�^���M", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("���M�Ɏ��s���܂����B", "�f�[�^���M�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// EXCEL�f�[�^�ēǂݍ���
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExcelReload_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Excel�̐ݒ�f�[�^���ēǂݍ��݂��Ă���낵���ł����H", "�f�[�^�ēǍ��m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                m_Excel = new ExcelInfo[2]; // Excel���

                // Excel�̃p�����[�^�ꗗ���擾�����X�g�쐬
                MakeParamCombo();
            }

        }
    }
}