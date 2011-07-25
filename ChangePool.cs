/*
 ===========================================================================
 Copyright (c) 2010 BrickRed Technologies Limited

 Permission is hereby granted, free of charge, to any person obtaining a copy
 of this software and associated documentation files (the "Software"), to deal
 in the Software without restriction, including without limitation the rights
 to use, copy, modify, merge, publish, distribute, sub-license, and/or sell
 copies of the Software, and to permit persons to whom the Software is
 furnished to do so, subject to the following conditions:

 The above copyright notice and this permission notice shall be included in
 all copies or substantial portions of the Software.

 THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 THE SOFTWARE.
 ===========================================================================
 */


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.SharePoint.Administration;
using System.Threading;
using System.Security;

namespace BrickRed.OpenSource.SharePointApplicationPoolChanger
{
    public partial class ChangePool : Form
    {
        SPFarm farm;
        SPWebService service;
        BackgroundWorker LoadWorker;
        BackgroundWorker PoolWorker;
        BackgroundWorker ChangeWorker;

        public ChangePool()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {


            LoadWorker = new BackgroundWorker();
            LoadWorker.WorkerReportsProgress = true;
            LoadWorker.WorkerSupportsCancellation = true;

            LoadWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(LoadWorker_RunWorkerCompleted);
            LoadWorker.ProgressChanged += new ProgressChangedEventHandler(LoadWorker_ProgressChanged);
            LoadWorker.DoWork += new DoWorkEventHandler(LoadWorker_DoWork);
            LoadWorker.RunWorkerAsync();


            ChangeWorker = new BackgroundWorker();

            ChangeWorker.WorkerReportsProgress = true;
            ChangeWorker.WorkerSupportsCancellation = true;

            ChangeWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ChangeWorker_RunWorkerCompleted);
            ChangeWorker.ProgressChanged += new ProgressChangedEventHandler(ChangeWorker_ProgressChanged);
            ChangeWorker.DoWork += new DoWorkEventHandler(ChangeWorker_DoWork);



            PoolWorker = new BackgroundWorker();

            PoolWorker.WorkerReportsProgress = true;
            PoolWorker.WorkerSupportsCancellation = true;

            PoolWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(PoolWorker_RunWorkerCompleted);
            PoolWorker.ProgressChanged += new ProgressChangedEventHandler(PoolWorker_ProgressChanged);
            PoolWorker.DoWork += new DoWorkEventHandler(PoolWorker_DoWork);


        }

        #region "Pool Worker"
        void PoolWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            PoolWorker.ReportProgress(0, "Loading existing applicaiton pools. Please wait... ");

            foreach (SPApplicationPool appPool in service.ApplicationPools)
            {
                PoolWorker.ReportProgress(1, appPool.Name);
            }

            PoolWorker.ReportProgress(0, "application pools loaded successfully...");

            if (PoolWorker.CancellationPending)
            {
                e.Cancel = true;
            }
        }

        void PoolWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            switch (e.ProgressPercentage)
            {
                case 1:
                    comboBoxExistingPool.Items.Add(Convert.ToString(e.UserState));
                    break;
                default:
                    labelMessage.Text = Convert.ToString(e.UserState);
                    break;
            }
        }

        void PoolWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("Pool Cancelled...");
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error. Details: " + (e.Error as Exception).ToString());
            }
            else
            {
                labelMessage.Text = "Application Pool loaded Successfully.";
            }
        }
        #endregion

        #region "Load Worker"

        void LoadWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            switch (e.ProgressPercentage)
            {
                case 1:
                    comboBoxWebApplications.Items.Add(Convert.ToString(e.UserState));
                    break;
                default:
                    labelMessage.Text = Convert.ToString(e.UserState);
                    break;
            }
        }

        void LoadWorker_DoWork(object sender, DoWorkEventArgs e)
        {

            LoadWorker.ReportProgress(0, "Initializing farm. Please wait... ");
            farm = SPFarm.Local;

            LoadWorker.ReportProgress(0, "Initializing service context. Please wait... ");
            service = farm.Services.GetValue<SPWebService>();

            LoadWorker.ReportProgress(0, "Loading web applicaitons. Please wait... ");

            foreach (SPWebApplication webApp in service.WebApplications)
            {
                LoadWorker.ReportProgress(1, webApp.Name);
            }

            LoadWorker.ReportProgress(0, "Web application loaded successfully...");

            if (LoadWorker.CancellationPending)
            {
                e.Cancel = true;
            }
        }

        void LoadWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("Loading Cancelled...");
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error. Details: " + (e.Error as Exception).ToString());
            }
            else
            {
                labelMessage.Text = "Web Applications loaded successfully...";
            }
        }
        #endregion

        #region "Change Worker"
        void ChangeWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            string applicationPoolName;
            string[] values = Convert.ToString(e.Argument).Split((char)127);
            SPApplicationPool objSPApplicaitionPool;
            string webApplicationName = values[2];

            if (values[0] == "1")
            {
                ChangeWorker.ReportProgress(0, "Fetching application pool. Please wait... ");
                applicationPoolName = values[1];
                objSPApplicaitionPool = service.ApplicationPools[applicationPoolName];

            }
            else
            {
                ChangeWorker.ReportProgress(0, "Creating new application pool. Please wait... ");
                applicationPoolName = textBoxNewAppPool.Text.Trim();
                objSPApplicaitionPool = new SPApplicationPool(applicationPoolName, service);

                switch (Convert.ToInt32(values[1]))
                {
                    case 0:
                        objSPApplicaitionPool.CurrentIdentityType = IdentityType.NetworkService;
                        break;
                    case 1:
                        objSPApplicaitionPool.CurrentIdentityType = IdentityType.SpecificUser;
                        objSPApplicaitionPool.Username = textBoxUserName.Text.Trim();

                        SecureString secureString = new SecureString();
                        string myPassword = textBoxPassword.Text;
                        foreach (char c in myPassword)
                        {
                            secureString.AppendChar(c);
                        }
                        secureString.MakeReadOnly();
                        objSPApplicaitionPool.SetPassword(secureString);
                        break;
                }


                ChangeWorker.ReportProgress(0, "Updating application pool. Please wait... ");
                objSPApplicaitionPool.Update();

                ChangeWorker.ReportProgress(0, "Deploying application pool. Please wait... ");
                objSPApplicaitionPool.Deploy();
            }



            //now set this to the web applicaiton
            ChangeWorker.ReportProgress(0, "Fetching Web application. Please wait... ");
            SPWebApplication objSPWebApplication = service.WebApplications[webApplicationName];

            ChangeWorker.ReportProgress(0, "Fetching application pool. Please wait... ");
            objSPApplicaitionPool = service.ApplicationPools[applicationPoolName];

            ChangeWorker.ReportProgress(0, "Setting new applicaiton pool. Please wait... ");
            objSPWebApplication.ApplicationPool = objSPApplicaitionPool;

            ChangeWorker.ReportProgress(0, "Updating web application. Please wait... ");
            objSPWebApplication.Update();

            ChangeWorker.ReportProgress(0, "Provisioning web application. Please wait... ");
            objSPWebApplication.ProvisionGlobally();

            ChangeWorker.ReportProgress(1, applicationPoolName);

        }

        void ChangeWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            switch (e.ProgressPercentage)
            {
                case 1:
                    labelAppPool.Text = Convert.ToString(e.UserState);
                    break;
                default:
                    labelMessage.Text = Convert.ToString(e.UserState);
                    break;
            }

        }

        void ChangeWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("Change Cancelled...");
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error. Details: " + (e.Error as Exception).ToString());
            }
            else
            {
                labelMessage.Text = "Application Pool Changed Successfully.";
            }
        }
        #endregion


        private void comboBoxWebApplications_SelectedIndexChanged(object sender, EventArgs e)
        {
            labelAppPool.Text = service.WebApplications[Convert.ToString(comboBoxWebApplications.SelectedItem)].ApplicationPool.Name;
            labelMessage.Text = "";
        }

        private void buttonAssignAppPool_Click(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    if (comboBoxWebApplications.SelectedIndex != -1)
                    {
                        if (!string.IsNullOrEmpty(textBoxNewAppPool.Text))
                        {
                            if (service.ApplicationPools[textBoxNewAppPool.Text.Trim()] == null)
                            {
                                switch (comboBoxIdentity.SelectedIndex)
                                {
                                    case 0:
                                        ChangeWorker.RunWorkerAsync(string.Format("{0}{1}{2}{3}{4}", tabControl1.SelectedIndex, (char)127, comboBoxIdentity.SelectedIndex, (char)127, comboBoxWebApplications.SelectedItem));
                                        break;
                                    case 1:
                                        if (!string.IsNullOrEmpty(textBoxUserName.Text) && !string.IsNullOrEmpty(textBoxPassword.Text))
                                        {
                                            ChangeWorker.RunWorkerAsync(string.Format("{0}{1}{2}{3}{4}", tabControl1.SelectedIndex, (char)127, comboBoxIdentity.SelectedIndex, (char)127, comboBoxWebApplications.SelectedItem));
                                        }
                                        else
                                        {
                                            MessageBox.Show("Please enter username / password.");
                                        }
                                        break;
                                    default:
                                        MessageBox.Show("Please select identity.");
                                        break;
                                }
                            }
                            else
                            {
                                MessageBox.Show("This application pool already exists. Please use existing pool tab for assigning existing pool or change pool name");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Please enter new application pool name.");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select web application.");
                    }
                    break;


                case 1:
                    if (comboBoxExistingPool.SelectedIndex != -1)
                    {
                        ChangeWorker.RunWorkerAsync(string.Format("{0}{1}{2}{3}{4}", tabControl1.SelectedIndex, (char)127, comboBoxExistingPool.SelectedItem, (char)127, comboBoxWebApplications.SelectedItem));
                    }
                    else
                    {
                        MessageBox.Show("Please select existing pool.");
                    }
                    break;

            }
        }

        private void comboBoxIdentity_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxIdentity.SelectedIndex == 1)
            {
                groupBoxCredentials.Visible = true;
            }
            else
            {
                groupBoxCredentials.Visible = false;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:

                    break;
                case 1:
                    comboBoxExistingPool.Items.Clear();
                    PoolWorker.RunWorkerAsync();
                    break;
            }
        }

    }
}
