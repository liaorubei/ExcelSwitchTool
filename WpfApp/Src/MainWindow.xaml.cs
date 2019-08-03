using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApp
{
    /// <summary>
    /// 表格转换工具
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<String> datas = new List<string>();
        private List<List<string>> exportData;

        public MainWindow()
        {
            InitializeComponent();
            this.DataGrid数据.LoadingRow += (o, s) => { s.Row.Header = s.Row.GetIndex() + 1; };
        }

        private void Button_选择表格_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "*.xls;*.xlsx|*.xls;*.xlsx";
            if (true == dialog.ShowDialog())
            {
                String fileName = dialog.FileName;
                this.主界面.IsEnabled = false;
                Task.Run(() =>
                {
                    String result = String.Empty;
                    try
                    {
                        IWorkbook workbook = null;
                        if (fileName.EndsWith(".xls"))
                        {
                            workbook = new HSSFWorkbook(new FileStream(fileName, FileMode.Open));
                        }
                        else
                        {
                            workbook = new XSSFWorkbook(new FileStream(fileName, FileMode.Open));
                        }

                        var sheet = workbook.GetSheetAt(0);
                        var lastRowNum = sheet.LastRowNum;

                        //清除原始数据
                        this.datas.Clear();

                        //填充最新数据
                        for (int i = 0; i <= lastRowNum; i++)
                        {
                            var row = sheet.GetRow(i);
                            var value = row.GetCell(0).StringCellValue;
                            datas.Add(value);
                        }

                        workbook.Close();
                    }
                    catch (Exception ex)
                    {
                        result = "打开或读取失败" + ex.Message;
                    }

                    this.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (String.IsNullOrEmpty(result))
                        {
                            this.DataGrid数据.ItemsSource = datas.Select(i => new { Data = i });
                        }
                        else
                        {
                            MessageBox.Show(result, "提示");
                        }
                        this.主界面.IsEnabled = true;
                    }));
                });
            }
        }

        private void Button_开始生成_Click(object sender, RoutedEventArgs e)
        {
            Int32 rowCount = Convert.ToInt32(this.TextBoxRowNum.Text);

            Task.Run(() =>
            {
                DataTable dt = new DataTable();

                List<List<String>> rows = new List<List<string>>();
                this.exportData = rows;
                for (int i = 0; i < rowCount; i++)
                {
                    rows.Add(new List<string>());
                }

                int count = 0;
                while (count < this.datas.Count)
                {
                    foreach (var row in rows)
                    {
                        if (count < this.datas.Count)
                        {
                            row.Add(this.datas[count]);
                            count++;
                        }
                    }
                }

                for (int i = 0; i < rows[0].Count; i++)
                {
                    dt.Columns.Add(i.ToString());
                }

                foreach (var row in rows)
                {
                    var dataRow = dt.NewRow();
                    for (int i = 0; i < row.Count; i++)
                    {
                        dataRow[i.ToString()] = row[i];
                    }

                    dt.Rows.Add(dataRow);
                }

                this.Dispatcher.BeginInvoke(new Action(() =>
                {
                    this.DataGrid数据.Columns.Clear();
                    this.DataGrid数据.ItemsSource = null;
                    this.DataGrid数据.ItemsSource = dt.DefaultView;
                }));
            });
        }

        private void Button_导出数据_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "*.xlsx|*.xlsx";
            if (true == dialog.ShowDialog())
            {
                if (this.exportData != null)
                {
                    var workbook = new XSSFWorkbook();
                    var sheet = workbook.CreateSheet("导出数据");
                    for (int i = 0; i < this.exportData.Count; i++)
                    {
                        var row = sheet.CreateRow(i);
                        for (int j = 0; j < this.exportData[i].Count; j++)
                        {
                            var cell = row.CreateCell(j);
                            cell.SetCellValue(this.exportData[i][j]);
                        }
                    }

                    workbook.Write(new FileStream(dialog.FileName, FileMode.Create));
                    workbook.Close();
                }
            }
        }
    }
}