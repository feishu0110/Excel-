using System;
using System.Collections.Generic;
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
using System.Windows.Forms;
using System.Windows.Interop;
using Microsoft.Office.Interop.Excel;
using Application =Microsoft.Office.Interop.Excel.Application;
using System.IO;
using MessageBox= System.Windows.Forms.MessageBox;
using Path = System.IO.Path;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace 合并多个Excel
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnSelFolder_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog folderDialog = new System.Windows.Forms.FolderBrowserDialog();
            folderDialog.SelectedPath = "C:\\"; // 默认选择C盘根目录  
            folderDialog.Description = "请选择一个文件夹：";
            folderDialog.ShowNewFolderButton = true; // 允许创建新文件夹  

            if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // 用户选择了一个文件夹，可以在这里进行相应的操作  
                
              txtFolder.Text = folderDialog.SelectedPath;
            }
        }

        private void Combine_Click(object sender, RoutedEventArgs e)
        {
            if (txtFolder.Text == "")
            {
                MessageBox.Show("请先选择文件夹", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            progressBar.Visibility = Visibility.Visible;
            double n = 1;
            progressBar.Dispatcher.Invoke(new Action<System.Windows.DependencyProperty, object>(progressBar.SetValue), System.Windows.Threading.DispatcherPriority.Background, 
               System.Windows.Controls.ProgressBar.ValueProperty, n);
            Application excelApp = new Application();
            if (IsNumeric(txtRow.Text) == false)
            {
                MessageBox.Show("行号必须为数字","提示",MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            int row = Convert.ToInt16(txtRow.Text);
           
            int k = 1;
            Workbook workbookCombine = null;
            Worksheet SheetCombine = null;
            string[] allFiles = Directory.GetFiles(txtFolder.Text);
            string[] filteredFiles = allFiles
            .Where(file => ((file.EndsWith(".xls") || file.EndsWith(".xlsx"))&&!file.EndsWith("合并.xls") && !file.EndsWith("合并_1.xls") && !file.EndsWith("合并_2.xls")
             && !file.EndsWith("合并_3.xls") && !file.EndsWith("合并_4.xls") && !file.EndsWith("合并_5.xls")))
            .ToArray();
            string strLog = "";
            foreach (string filename in filteredFiles)// 遍历文件夹中的所有Excel文件  
            {
                strLog += "正在读取：" + filename + Environment.NewLine;             
                txtLog.Text = strLog;
                n = (float)(k)/ filteredFiles.Length*100;
                progressBar.Dispatcher.Invoke(new Action<System.Windows.DependencyProperty, object>(progressBar.SetValue), System.Windows.Threading.DispatcherPriority.Background,
                   System.Windows.Controls.ProgressBar.ValueProperty, n);
                if (k == 1)
                {
                    string destinationFilePath = GenerateUniqueFileName(txtFolder.Text+"\\合并.xls", txtFolder.Text);                  // 复制 Excel 文件
                  

                    // 复制文件
                    File.Copy(filename, destinationFilePath);
                    workbookCombine= excelApp.Workbooks.Open(destinationFilePath);
                    SheetCombine = GetSheetByName(txtSheet.Text, workbookCombine);
                    if (SheetCombine == null)
                    {
                        MessageBox.Show("所选文件夹的第一个Excel中未找到Sheet:" + txtSheet.Text, "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        workbookCombine.Close(false);
                        excelApp.Quit();
                        excelApp = null;
                        return;
                    }
                    if(SheetCombine.Name!= txtSheet.Text)
                    {
                        MessageBox.Show("未找到Sheet:" + txtSheet.Text+"，将使用Excel第一个工作表。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                Workbook workbook1 = excelApp.Workbooks.Open(filename);
                Worksheet xlSheet = GetSheetByName(txtSheet.Text, workbook1);
                if (xlSheet != null)
                {
                    for (int i = 1; i <= 100; i++)
                    {
                        SheetCombine.Cells[row + k - 1,i].Value = xlSheet.Cells[row, i].Value;
                    }
                   
                }
                workbook1.Close(false);
                strLog += "已完成：" + filename + Environment.NewLine;
                k++;
                txtLog.Text = strLog;
            }
            
            if (k == 1)
            {
                
                MessageBox.Show("所选文件夹未找到Excel文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                excelApp.Quit();
                excelApp = null;
            }
            else {

                MessageBox.Show("完成", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                excelApp.Visible = true;
            }
            progressBar.Visibility = Visibility.Hidden;
        }
        static bool IsNumeric(string value)
        {
            int result;
            return Int32.TryParse(value, out result);
        }
        private Worksheet GetSheetByName(string sheetName,Workbook workbook)
        {
            if (workbook.Sheets.Count > 0) // 判断工作簿是否至少包含一个工作表  
            {
                foreach (Worksheet worksheet in workbook.Sheets) // 遍历所有工作表  
                {
                    if (worksheet.Name == sheetName) // 判断工作表名称是否为"Sheet1"  
                    {
                       
                        return worksheet; // 如果找到，则结束循环并输出结果  
                       
                    }
                }
                return workbook.Sheets[1];
            }
           
            return null;
        }
        static string GenerateUniqueFileName(string filePath, string destinationFolder)
        {
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
            string fileExtension = Path.GetExtension(filePath);

            int suffix = 1;
            string destinationFilePath = Path.Combine(destinationFolder, $"{fileNameWithoutExtension}_{suffix}{fileExtension}");

            // 如果文件名已存在，则添加后缀
            while (File.Exists(destinationFilePath))
            {
                suffix++;
                destinationFilePath = Path.Combine(destinationFolder, $"{fileNameWithoutExtension}_{suffix}{fileExtension}");
            }

            return destinationFilePath;
        }
    }
}
