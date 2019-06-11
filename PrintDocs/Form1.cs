using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Printing;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PrintDocs
{


    public partial class Form1 : Form
    {

        public class PrintDoc
        {
            public string fileName;
            public string fullName;

            public PrintDoc()
            {
                fileName = "";
                fullName = "";
            }

        }

        string[] allfiles;
        BindingList<PrintDoc> fileList = new BindingList<PrintDoc>();



        public Form1()
        {
            InitializeComponent();
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            //DialogResult dr = openFileDialog1.ShowDialog();

            //string[] s = openFileDialog1.FileName.Split('.');

            //if (dr.ToString() == "OK")

            //{

            //    if (s.Length > 1)

            //        if (s[1] == "doc" || s[1] == "docx" || s[1] == "jpg")

            //            txtFileName.Text = openFileDialog1.FileName;

            //        else

            //            MessageBox.Show("Please select doc,docx,jpeg file !!");

            //}
        }

        private void BtnPrint_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {

                string printFileName = row.Cells[0].Value.ToString();
                string printFile = row.Cells[1].Value.ToString();
                //Using below code we can print any document


                ProcessStartInfo info = new ProcessStartInfo(printFile.Trim());

                info.Verb = "Print";

                info.CreateNoWindow = true;

                info.WindowStyle = ProcessWindowStyle.Hidden;

                Process.Start(info);


                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();

                bool success = checkPrintSuccess(printFile);
                if (success) moveFiles(printFileName, printFile);
            }


        }

        private void Form1_Load(object sender, EventArgs e)
        {
            allfiles = Directory.GetFiles(@"Z:\TEST\staffs\Nini\letters\", "*.*", SearchOption.AllDirectories);
            bindingSource1 = new BindingSource();

            foreach (var file in allfiles)
            {
                FileInfo info = new FileInfo(file);
                // Do something with the Folder or just add them to a list via nameoflist.add();
                PrintDoc pq = new PrintDoc();
                pq.fileName = info.Name;
                pq.fullName = info.FullName;

                fileList.Add(pq);
                addPrintQueueToList(pq);

            }


        }

        private void addPrintQueueToList(PrintDoc pq)
        {
            dataGridView1.Rows.Add();
            DataGridViewRow row = dataGridView1.Rows[dataGridView1.Rows.Count - 1];

            row.Cells[FileName.Index].Value = pq.fileName;
            row.Cells[FullName.Index].Value = pq.fullName;
        }


        private void moveFiles(string printDoc, string printFilePath)
        {
            //foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            //{

            //string fileName = row.Cells[0].Value.ToString();
            //string sourceFile = @"" + row.Cells[1].Value.ToString();
            string sourceFile = @"" + printFilePath;
            string targetPath = @"Z:\TEST\staffs\Nini\letters\SUDAH PRINT\";

            string destFile = Path.Combine(targetPath, printDoc);

            if (!Directory.Exists(targetPath))
            {
                Directory.CreateDirectory(targetPath);
            }

            File.Move(sourceFile, destFile);

            //}

            MessageBox.Show("done");

        }


        private bool checkPrintSuccess(string docName)
        {

            bool success = false;
            PrintServer myPrintServer = new PrintServer();

            // List the print server's queues
            PrintQueueCollection myPrintQueues = myPrintServer.GetPrintQueues();
            String printQueueNames = "My Print Queues:\n\n";
            string printJob = "JOBS:\n\n";


            foreach (PrintQueue pq in myPrintQueues)
            {
                printQueueNames += "\t" + pq.Name + "\n";

                //HP LaserJet Pro MFP M127fw

                if (pq.Name == "HP LaserJet Pro MFP M127fw")
                {
                    try
                    {
                        PrintJobInfoCollection jobs = pq.GetPrintJobInfoCollection();


                        if (jobs != null)
                        {

                            foreach (PrintSystemJobInfo job in jobs)
                            {
                                printJob += "\t" + job.Name + "\n";

                                if (job.Name == docName && (((job.JobStatus & PrintJobStatus.Completed) == PrintJobStatus.Completed) || ((job.JobStatus & PrintJobStatus.Printed) == PrintJobStatus.Printed)))
                                {
                                    success = true;
                                }
                                else
                                    success = false;
                            }// end for each print job 
                        }
                    }
                    catch (NullReferenceException e)
                    {
                        success = true;

                    }

                    break;
                }
            }
            Console.WriteLine(printQueueNames);

            MessageBox.Show(printJob);
            Console.WriteLine("\nPress Return to continue.");
            Console.ReadLine();

            return success;
        }

        // Check for possible trouble states of a print job using the flags of the JobStatus property 
        internal static void SpotTroubleUsingJobAttributes(PrintSystemJobInfo theJob)
        {
            if ((theJob.JobStatus & PrintJobStatus.Blocked) == PrintJobStatus.Blocked)
            {
                Console.WriteLine("The job is blocked.");
            }
            if (((theJob.JobStatus & PrintJobStatus.Completed) == PrintJobStatus.Completed)
                ||
                ((theJob.JobStatus & PrintJobStatus.Printed) == PrintJobStatus.Printed))
            {
                Console.WriteLine("The job has finished. Have user recheck all output bins and be sure the correct printer is being checked.");
            }
            if (((theJob.JobStatus & PrintJobStatus.Deleted) == PrintJobStatus.Deleted)
                ||
                ((theJob.JobStatus & PrintJobStatus.Deleting) == PrintJobStatus.Deleting))
            {
                Console.WriteLine("The user or someone with administration rights to the queue has deleted the job. It must be resubmitted.");
            }
            if ((theJob.JobStatus & PrintJobStatus.Error) == PrintJobStatus.Error)
            {
                Console.WriteLine("The job has errored.");
            }
            if ((theJob.JobStatus & PrintJobStatus.Offline) == PrintJobStatus.Offline)
            {
                Console.WriteLine("The printer is offline. Have user put it online with printer front panel.");
            }
            if ((theJob.JobStatus & PrintJobStatus.PaperOut) == PrintJobStatus.PaperOut)
            {
                Console.WriteLine("The printer is out of paper of the size required by the job. Have user add paper.");
            }

            if (((theJob.JobStatus & PrintJobStatus.Paused) == PrintJobStatus.Paused)
                ||
                ((theJob.HostingPrintQueue.QueueStatus & PrintQueueStatus.Paused) == PrintQueueStatus.Paused))
            {
                //HandlePausedJob(theJob);
                //HandlePausedJob is defined in the complete example.
            }

            if ((theJob.JobStatus & PrintJobStatus.Printing) == PrintJobStatus.Printing)
            {
                Console.WriteLine("The job is printing now.");
            }
            if ((theJob.JobStatus & PrintJobStatus.Spooling) == PrintJobStatus.Spooling)
            {
                Console.WriteLine("The job is spooling now.");
            }
            if ((theJob.JobStatus & PrintJobStatus.UserIntervention) == PrintJobStatus.UserIntervention)
            {
                Console.WriteLine("The printer needs human intervention.");
            }

        }



    }


}

