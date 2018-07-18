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
using OfficeOpenXml;
using System.IO;
 
using System.Net.Mail;

namespace examm
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        
        public MainWindow()
        {
            InitializeComponent();

        }
        public void Excel()
        {
            string path = @"C:/Users/акылбекова/Documents/Visual Studio 2015/Projects/examm/examm/bin/sample/xlsx";
            FileInfo fi = new FileInfo(path);
            using (ExcelPackage package =
                new ExcelPackage(fi, true))
            {
                ExcelWorksheet workshet =
                    package.Workbook.Worksheets["Exam"];

                model.newEquipment d = new model.newEquipment();
                //if (d.Where(o => o.CreateDate >= Calend1.SelectedDate && o.LastDate <= Calend2.SelectedDate).Count() < 1)
                //{
                //    TextBlocks.Text = "NO Such data";
                //    return;
                //}
                for (int i = 0; i < workshet.Dimension.End.Row; i++)
                {
                    model.newEquipment hh = new model.newEquipment();
                    model.TableEvaluationSysStatus vv = new model.TableEvaluationSysStatus();
                    model.TrackEvaluationPart ff = new model.TrackEvaluationPart();
                    workshet.Cells[i, 1].Value =  hh.intEquipmentID;
                    workshet.Cells[i, 2].Value = hh.intModelID;
                    workshet.Cells[i, 3].Value = vv.intEvaluationSysStatusId;
                    workshet.Cells[i, 4].Value = hh.CreateDate;
                    float EttR = 0;
                    float Eltr = 0;
                    workshet.Cells[i, 5].Value = EttR;
                    workshet.Cells[i, 6].Value = Eltr;
                    workshet.Cells[i, 7].Value = vv.intEvaluationSysStatusId;
                    workshet.Cells[i, 8].Value = ff.intEvalutionId.Value;
                    

                }



                //FI
                try
                {
                    FileStream fs = File.Create(path);
                    fs.Close();
                    package.SaveAs(fs);

                }
                catch (Exception ex)
                {
                     TextBlocks.Text = ex.Message;
                }
                
            }
        }
      
        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            MessageBox.Show(textBox.Text);
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            Excel();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            SubWindow subWindow = new SubWindow();
            subWindow.Show();

            string to =  ;
            string from = "info@shag.com";
            MailMessage message = new MailMessage(from, to);
            message.Subject = "Using the new SMTP client.";
            message.Body = @"Using this new feature, you can send an e-mail message from an application very easily.";
            SmtpClient client = new SmtpClient(server);
          
            client.UseDefaultCredentials = true;
            
            try
            {
                client.Send(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught in CreateTestMessage2(): {0}",
                            ex.ToString());
            }
    }
}
