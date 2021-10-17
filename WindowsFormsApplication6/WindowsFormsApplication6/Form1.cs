using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using System.Diagnostics;
using System.IO;
using SolidWorks.Interop.cosworks;
using System.Collections;
using System.Runtime.InteropServices;
using eDrawings.Interop.EModelViewControl;
using EModelView;

namespace WindowsFormsApplication6
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            comboBox1.Text = "Сталь";
            comboBox2.Text = "Сталь";
            comboBox3.Text = "Сталь";
            comboBox4.Text = "Сталь";
            comboBox5.Text = "Статический";
            comboBox6.Text = "Об/мин";
            comboBox7.Text = "Об/мин^2";


            textBox14.Visible = false;
            textBox15.Visible = false;
            textBox16.Visible = false;
            textBox17.Visible = false;
            textBox18.Visible = false;
            textBox19.Visible = false;
            textBox20.Visible = false;
            textBox21.Visible = false;
            textBox22.Visible = false;
            textBox23.Visible = false;
            textBox24.Visible = false;
            textBox25.Visible = false;
            textBox26.Visible = false;
            textBox27.Visible = false;
            textBox28.Visible = false;
            textBox29.Visible = false;
            textBox30.Visible = false;
            textBox31.Visible = false;
            textBox32.Visible = false;
            textBox33.Visible = false;
            textBox1.Text = "55";
            textBox3.Text = "155";
            textBox4.Text = "280";
            textBox5.Text = "130";
            textBox6.Text = "55";
            textBox2.Text = "40";
            textBox7.Text = "60";
            textBox8.Text = "160";
            textBox9.Text = "125";
            textBox10.Text = "240";
            textBox11.Text = "60";
            textBox12.Text = "160";
            textBox13.Text = "50";
            tabPage5.Parent = null;
            tabPage6.Parent = null;
            tabPage7.Parent = null;
            tabPage8.Parent = null;
            tabPage9.Parent = null;
            tabPage10.Parent = null;
            tabPage11.Parent = null;
            tabPage12.Parent = null;
            tabPage13.Parent = null;
            tabPage14.Parent = null;
            tabPage15.Parent = null;
            tabPage16.Parent = null;
            tabPage17.Parent = null;
        }
        SldWorks SwApp;
        IModelDoc2 swModel;
        ModelDoc2 Part = default(ModelDoc2);

        int errors = 0;
        int warnings = 0;
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        object PIDpodsh11 = null;
        object PIDpodsh22 = null;
        object PIDring = null;
        object PIDload = null;
        object PIDforce1 = null;
        object PIDforce2 = null;
        object PIDforce3 = null;
        object PIDforce4 = null;


        private void getReference()
        {
            Part = SwApp.ActiveDoc;
            object obj = null;
            bool boolstatus = false;


            boolstatus = Part.Extension.SelectByID2("", "FACE", -0.209637845877239, 1.32960929494743E-02, 8.69897345223194E-02, false, 0, null, 0);
            obj = ((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
            PIDpodsh11 = Part.Extension.GetPersistReference3(obj);

            boolstatus = Part.Extension.SelectByID2("", "FACE", 0.204445254567361, 6.74646597611286E-02, 5.65023865275407E-02, false, 0, null, 0);
            obj = ((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
            PIDpodsh22 = Part.Extension.GetPersistReference3(obj);

            boolstatus = Part.Extension.SelectByID2("", "FACE", 1.99824949561389E-04, 9.64164394427485E-02, 8.41894898735518E-02, false, 0, null, 0);
            obj = ((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
            PIDring = Part.Extension.GetPersistReference3(obj);

            boolstatus = Part.Extension.SelectByID2("", "FACE", -0.502825589827751, 2.67489972691237E-02, 2.64244421907165E-02, false, 0, null, 0);
            obj = ((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
            PIDload = Part.Extension.GetPersistReference3(obj);

            boolstatus = Part.Extension.SelectByID2("", "FACE", -0.464999999999918, 1.72261506710925E-03, 1.78543327359648E-03, false, 0, null, 0);
            obj = ((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
            PIDforce1 = Part.Extension.GetPersistReference3(obj);

            boolstatus = Part.Extension.SelectByID2("", "FACE", -0.509999999999934, -1.15601135849204E-02, 4.55654671370098E-02, false, 0, null, 0);
            obj = ((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
            PIDforce2 = Part.Extension.GetPersistReference3(obj);

            boolstatus = Part.Extension.SelectByID2("", "FACE", 0.439999999999827, -1.87938504719298E-03, -8.56435392444155E-03, false, 0, null, 0);
            obj = ((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
            PIDforce3 = Part.Extension.GetPersistReference3(obj);

            boolstatus = Part.Extension.SelectByID2("", "FACE", 0.489999999999839, -5.93841323888284E-02, 7.41454703458082E-02, false, 0, null, 0);
            obj = ((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
            PIDforce4 = Part.Extension.GetPersistReference3(obj);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Form FormName = new Form4();

            FormName.Show();

            //убиваем процесс
            Process[] processes = Process.GetProcessesByName("SLDWORKS");
            foreach (Process process in processes)
            {
                process.CloseMainWindow();
                process.Kill();
            }

            //запускаем
            Guid myGuid = new Guid("F16137AD-8EE8-4D2A-8CAC-DFF5D1F67522");
            object processSw = System.Activator.CreateInstance(System.Type.GetTypeFromCLSID(myGuid));


            SwApp = (SldWorks)processSw;
            Part = SwApp.OpenDoc6("D:/диплом/ротор/ротор.SLDASM", 2, 0, "", errors, warnings);
            SwApp.Visible = true;

            button2.Visible = true;
            button1.Visible = false;
            button3.Visible = true;

            string writePathPod11 = @"D:/диплом\ротор\pod11.txt";

            using (StreamWriter sw = new StreamWriter(writePathPod11, false, System.Text.Encoding.Default))
            {
                sw.WriteLine("\"pod\"=84");
            }
            string writePathPod22 = @"D:/диплом\ротор\pod22.txt";

            using (StreamWriter sw = new StreamWriter(writePathPod22, false, System.Text.Encoding.Default))
            {
                sw.WriteLine("\"pod\"=84");
            }
            string writePathRing = @"D:/диплом\ротор\ring.txt";

            using (StreamWriter sw = new StreamWriter(writePathRing, false, System.Text.Encoding.Default))
            {
                sw.WriteLine("\"ring\"=124");

            }
            string writePath = @"D:/диплом\ротор\globalperem.txt";

            using (StreamWriter sw = new StreamWriter(writePath, false, System.Text.Encoding.Default))
            {
                sw.WriteLine("\"otvnar\"=20");
                sw.WriteLine("\"dlinotv\"=55");
            }

            string writePath1 = @"D:/диплом\ротор\val.txt";

            using (StreamWriter sw = new StreamWriter(writePath1, false, System.Text.Encoding.Default))
            {
                sw.WriteLine("\"otvnar\"=20");
                sw.WriteLine("\"dlinaotv\"=55");
                sw.WriteLine("\"l2\"=155");
                sw.WriteLine("\"l3\"=280");
                sw.WriteLine("\"l4\"=130");
                sw.WriteLine("\"l5\"=55");
                sw.WriteLine("\"d2\"=30");
                sw.WriteLine("\"d4\"=62.5");
                sw.WriteLine("\"d5\"=120");
                sw.WriteLine("\"d6\"=30");
                sw.WriteLine("\"d8\"=25");

            }

            string writePath2 = @"D:/диплом\ротор\podsh1.txt";

            using (StreamWriter sw = new StreamWriter(writePath2, false, System.Text.Encoding.Default))
            {
                sw.WriteLine("\"d2\"=30");
                sw.WriteLine("\"d3\"=80");
            }

            string writePath3 = @"D:/диплом\ротор\podsh2.txt";
            using (StreamWriter sw = new StreamWriter(writePath3, false, System.Text.Encoding.Default))
            {
                sw.WriteLine("\"d6\"=30");
                sw.WriteLine("\"d7\"=80");
            }

            string writePath4 = @"D:/диплом\ротор\kompresor.txt";
            using (StreamWriter sw = new StreamWriter(writePath4, false, System.Text.Encoding.Default))
            {
                sw.WriteLine("\"otvna\"=25");
                sw.WriteLine("\"l1\"=55");

            }

            bool boolstatus = false;
            AssemblyDoc swAssembly = null;
            swModel = SwApp.ActiveDoc;
            boolstatus = swModel.Extension.SelectByID2("Повернуть1@валл-1@ротор", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
            swModel.FeatEdit();
            swModel.ClearSelection2(true);
            swModel.SketchManager.InsertSketch(true);
            boolstatus = swModel.Extension.SelectByID2("Повернуть1@левоеколесо-1@ротор", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
            swModel.FeatEdit();
            swModel.ClearSelection2(true);
            swModel.SketchManager.InsertSketch(true);
            swModel.InsertSketch();
            swAssembly = ((AssemblyDoc)(swModel));
            swAssembly.EditAssembly();
            getReference();
            int fileerror = SwApp.LoadAddIn("D:\\проги\\Solid\\SOLIDWORKS\\Simulation\\cosworks.dll");


            //заполняем комбобокс индексами анализов 
            CosmosWorks COSMOSWORKS;
            CwAddincallback CWObject;
            CWModelDoc ActDoc;
            CWStudyManager StudyMngr;
            

            CWObject = (CwAddincallback)SwApp.GetAddInObject("SldWorks.Simulation");
            COSMOSWORKS = CWObject.CosmosWorks;

            ActDoc = COSMOSWORKS.ActiveDoc;


            StudyMngr = ActDoc.StudyManager;
            int scount;
            scount = StudyMngr.StudyCount;
            if (scount != 0)
            {
                for (int i = 0; i < scount; i++)
                {
                    comboBox8.Items.Add(StudyMngr.GetStudy(i).Name);
                }
            }
            FormName.Close();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            string l1, l2, l3, l4, l5, r1, r2, r3, r4, r5, r6, r7, r8;
            double d1, d2, d3, d4, d5, d6, d7, d8;
            if (textBox1.Text == "")
            {
                textBox1.Text = "0";
            }
            if (textBox2.Text == "")
            {
                textBox2.Text = "0";
            }
            if (textBox3.Text == "")
            {
                textBox3.Text = "0";
            }
            if (textBox4.Text == "")
            {
                textBox4.Text = "0";
            }
            if (textBox5.Text == "")
            {
                textBox5.Text = "0";
            }
            if (textBox6.Text == "")
            {
                textBox6.Text = "0";
            }
            if (textBox7.Text == "")
            {
                textBox7.Text = "0";
            }
            if (textBox8.Text == "")
            {
                textBox8.Text = "0";
            }
            if (textBox9.Text == "")
            {
                textBox9.Text = "0";
            }
            if (textBox10.Text == "")
            {
                textBox10.Text = "0";
            }
            if (textBox11.Text == "")
            {
                textBox11.Text = "0";
            }
            if (textBox12.Text == "")
            {
                textBox12.Text = "0";
            }
            if (textBox13.Text == "")
            {
                textBox13.Text = "0";
            }
            l1 = textBox1.Text;//55 (30<=x<=80)
            d1 = Convert.ToDouble(textBox2.Text) / 2;//20 (8<=x<=48)
            r1 = Convert.ToString(d1);
            r1 = r1.Replace(" ", "").Replace(",", ".");


            l2 = textBox3.Text;//150
            l3 = textBox4.Text;//280
            l4 = textBox5.Text;//130
            l5 = textBox6.Text;//55
            d2 = Convert.ToDouble(textBox7.Text) / 2;//30
            d3 = Convert.ToDouble(textBox8.Text) / 2;//80
            d4 = Convert.ToDouble(textBox9.Text) / 2;//62.5
            d5 = Convert.ToDouble(textBox10.Text) / 2;//120
            d6 = Convert.ToDouble(textBox11.Text) / 2;//30
            d7 = Convert.ToDouble(textBox12.Text) / 2;//80
            d8 = Convert.ToDouble(textBox13.Text) / 2;//25

            r2 = Convert.ToString(d2);
            r2 = r2.Replace(" ", "").Replace(",", ".");
            r3 = Convert.ToString(d3);
            r3 = r3.Replace(" ", "").Replace(",", ".");
            r4 = Convert.ToString(d4);
            r4 = r4.Replace(" ", "").Replace(",", ".");
            r5 = Convert.ToString(d5);
            r5 = r5.Replace(" ", "").Replace(",", ".");
            r6 = Convert.ToString(d6);
            r6 = r6.Replace(" ", "").Replace(",", ".");
            r7 = Convert.ToString(d7);
            r7 = r7.Replace(" ", "").Replace(",", ".");
            r8 = Convert.ToString(d8);
            r8 = r8.Replace(" ", "").Replace(",", ".");

            if (Convert.ToDouble(l1) < 41.25 || Convert.ToDouble(l1) > 68.75 || (d2 - d1) < 6 || (d6 - d8) < 5 || Convert.ToDouble(l2) < 116.25 || Convert.ToDouble(l2) > 193.75 || Convert.ToDouble(l3) < 210 || Convert.ToDouble(l3) > 350 || Convert.ToDouble(l4) < 97.5 || Convert.ToDouble(l4) > 162.5 || Convert.ToDouble(l5) < 41.25 || Convert.ToDouble(l5) > 68.75 || d1 < 15 || d1 > 25 || d2 < 22.5 || d2 > 37.5 || d3 < 60 || d3 > 100 || d4 < 46.875 || d4 > 78.125 || d5 < 90 || d5 > 150 || d6 < 22.5 || d6 > 37.5 || d7 < 60 || d7 > 100 || d8 < 18.75 || d8 > 31.25)
            {
                Form FormName = new Form3();

                FormName.Show();
            }
            else
            {

                string writePathPod11 = @"D:/диплом\ротор\pod11.txt";

                using (StreamWriter sw = new StreamWriter(writePathPod11, false, System.Text.Encoding.Default))
                {
                    sw.WriteLine("\"pod\"=" + Convert.ToString(d3 + 4).Replace(" ", "").Replace(",", "."));
                }
                string writePathPod22 = @"D:/диплом\ротор\pod22.txt";

                using (StreamWriter sw = new StreamWriter(writePathPod22, false, System.Text.Encoding.Default))
                {
                    sw.WriteLine("\"pod\"=" + Convert.ToString(d7 + 4).Replace(" ", "").Replace(",", "."));
                }
                string writePathRing = @"D:/диплом\ротор\ring.txt";

                using (StreamWriter sw = new StreamWriter(writePathRing, false, System.Text.Encoding.Default))
                {
                    sw.WriteLine("\"ring\"=" + Convert.ToString(d5 + 4).Replace(" ", "").Replace(",", "."));

                }
                string writePath = @"D:/диплом\ротор\globalperem.txt";

                using (StreamWriter sw = new StreamWriter(writePath, false, System.Text.Encoding.Default))
                {
                    sw.WriteLine("\"otvnar\"=" + r1);
                    sw.WriteLine("\"dlinotv\"=" + l1.Replace(",", "."));
                }

                string writePath1 = @"D:/диплом\ротор\val.txt";

                using (StreamWriter sw = new StreamWriter(writePath1, false, System.Text.Encoding.Default))
                {
                    sw.WriteLine("\"otvnar\"=" + r1);
                    sw.WriteLine("\"dlinaotv\"=" + l1.Replace(",", "."));
                    sw.WriteLine("\"l2\"=" + l2.Replace(",", "."));
                    sw.WriteLine("\"l3\"=" + l3.Replace(",", "."));
                    sw.WriteLine("\"l4\"=" + l4.Replace(",", "."));
                    sw.WriteLine("\"l5\"=" + l5.Replace(",", "."));
                    sw.WriteLine("\"d2\"=" + r2);
                    sw.WriteLine("\"d4\"=" + r4);
                    sw.WriteLine("\"d5\"=" + r5);
                    sw.WriteLine("\"d6\"=" + r6);
                    sw.WriteLine("\"d8\"=" + r8);

                }

                string writePath2 = @"D:/диплом\ротор\podsh1.txt";

                using (StreamWriter sw = new StreamWriter(writePath2, false, System.Text.Encoding.Default))
                {
                    sw.WriteLine("\"d2\"=" + r2);
                    sw.WriteLine("\"d3\"=" + r3);
                }

                string writePath3 = @"D:/диплом\ротор\podsh2.txt";
                using (StreamWriter sw = new StreamWriter(writePath3, false, System.Text.Encoding.Default))
                {
                    sw.WriteLine("\"d6\"=" + r6);
                    sw.WriteLine("\"d7\"=" + r7);
                }

                string writePath4 = @"D:/диплом\ротор\kompresor.txt";
                using (StreamWriter sw = new StreamWriter(writePath4, false, System.Text.Encoding.Default))
                {
                    sw.WriteLine("\"otvna\"=" + r8);
                    sw.WriteLine("\"l1\"=" + l5.Replace(",", "."));

                }

                bool boolstatus = false;
                AssemblyDoc swAssembly = null;
                swModel = SwApp.ActiveDoc;
                boolstatus = swModel.Extension.SelectByID2("Повернуть1@валл-1@ротор", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
                swModel.FeatEdit();
                swModel.ClearSelection2(true);
                swModel.SketchManager.InsertSketch(true);
                boolstatus = swModel.Extension.SelectByID2("Повернуть1@левоеколесо-1@ротор", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
                swModel.FeatEdit();
                swModel.ClearSelection2(true);
                swModel.SketchManager.InsertSketch(true);
                swModel.InsertSketch();
                swAssembly = ((AssemblyDoc)(swModel));
                swAssembly.EditAssembly();

                /* swModel.AssemblyPartToggle();
                 swModel.EditAssembly();*/
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != ','))
            {
                e.Handled = true;
            }

        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox13_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void информацияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form FormName = new Form2();
            FormName.Show();
        }


        private void button3_Click(object sender, EventArgs e)
        {
            Form FormName = new Form5();

            FormName.Show();
            dynamic COSMOSWORKS = default(dynamic);
            dynamic COSMOSObject = default(dynamic);
            CWModelDoc ActDoc = default(CWModelDoc);
            CWStudyManager StudyMngr = default(CWStudyManager);
            CWStudy Study = default(CWStudy);
            CWSolidManager SolidMgr = default(CWSolidManager);
            CWSolidComponent SolidComponent = default(CWSolidComponent);
            CWSolidBody SolidBody = default(CWSolidBody);
            string SName = null;
            int errors = 0;
            int warnings = 0;
            int errorCode = 0;
            int errCode = 0;
            int CompCount = 0;
            int j = 0;
            CWMaterial CWMat = null;
            int iApp = 0;
            bool bApp = false;

            // Determine host SOLIDWORKS version
            int swVersion = Convert.ToInt32(SwApp.RevisionNumber().Substring(0, 2));
            // Calculate the correct Simulation ProgID for this version of SOLIDWORKS
            int cwVersion = swVersion - 15;
            String cwProgID = String.Format("SldWorks.Simulation.{0}", cwVersion);
            Debug.Print(cwProgID);

            COSMOSObject = SwApp.GetAddInObject(cwProgID);

            COSMOSWORKS = COSMOSObject.CosmosWorks;

            ActDoc = (CWModelDoc)COSMOSWORKS.ActiveDoc;

            StudyMngr = (CWStudyManager)ActDoc.StudyManager;

            switch (comboBox5.Text)
            {
                case "Статический":
                    Study = (CWStudy)StudyMngr.CreateNewStudy(textBox36.Text, (int)swsAnalysisStudyType_e.swsAnalysisStudyTypeStatic, (int)swsMeshType_e.swsMeshTypeSolid, out errCode);
                    break;
                case "Колебания":
                    Study = (CWStudy)StudyMngr.CreateNewStudy(textBox36.Text, (int)swsAnalysisStudyType_e.swsAnalysisStudyTypeFrequency, (int)swsMeshType_e.swsMeshTypeSolid, out errCode);
                    break;
                case "Гармонический":
                    Study = (CWStudy)StudyMngr.CreateNewStudy3(textBox36.Text, (int)swsAnalysisStudyType_e.swsAnalysisStudyTypeDynamic, (int)swsDynamicAnalysisSubType_e.swsDynamicAnalysisSubTypeHarmonic, out errCode);
                    break;
                case "Переходные процессы":
                    Study = (CWStudy)StudyMngr.CreateNewStudy3(textBox36.Text, (int)swsAnalysisStudyType_e.swsAnalysisStudyTypeDynamic, (int)swsDynamicAnalysisSubType_e.swsDynamicAnalysisSubTypeTransient, out errCode);
                    break;
            }
            SolidMgr = (CWSolidManager)Study.SolidManager;

            CompCount = SolidMgr.ComponentCount;

            switch (comboBox1.Text)
            {
                case "Сталь":
                    for (j = 7; j <= 8; j++)
                    {
                        SolidComponent = SolidMgr.GetComponentAt(j, out errorCode);

                        SName = SolidComponent.ComponentName;

                        SolidBody = (CWSolidBody)SolidComponent.GetSolidBodyAt(0, out errCode);

                        CWMat = (CWMaterial)SolidBody.GetDefaultMaterial();
                        CWMat.MaterialUnits = 0;

                        CWMat.MaterialName = "Сталь";
                        CWMat.SetPropertyByName("EX", 210000000000.0, 0);
                        CWMat.SetPropertyByName("NUXY", 0.3, 0);
                        CWMat.SetPropertyByName("DENS", 7800, 0);
                        CWMat.SetPropertyByName("SIGYLD", 600000000, 0);
                        errCode = SolidBody.SetSolidBodyMaterial(CWMat);

                        SolidBody = null;
                        SolidComponent = null;
                    }
                    /*for (j = 1; j <= (CompCount - 1); j++)
                    {
                        SolidComponent = (CWSolidComponent)SolidMgr.GetComponentAt(j, out errorCode);
                        SName = SolidComponent.ComponentName;
                        SolidBody = (CWSolidBody)SolidComponent.GetSolidBodyAt(0, out errCode);
                        if (errCode != 0)
                            iApp = (int)SolidBody.SetLibraryMaterial("D:\\проги\\Solid\\SOLIDWORKS\\lang\\english\\sldmaterials\\solidworks materials.sldmat", "Gray Cast Iron (SN)");
                        bApp = System.Convert.ToBoolean(iApp);

                        SolidBody = null;
                    }*/
                    break;
                case "Алюминий":

                    for (j = 7; j <= 8; j++)
                    {
                        SolidComponent = SolidMgr.GetComponentAt(j, out errorCode);

                        SName = SolidComponent.ComponentName;

                        SolidBody = (CWSolidBody)SolidComponent.GetSolidBodyAt(0, out errCode);

                        CWMat = (CWMaterial)SolidBody.GetDefaultMaterial();
                        CWMat.MaterialUnits = 0;

                        CWMat.MaterialName = "Алюминий";
                        CWMat.SetPropertyByName("EX", 72000000000.0, 0);
                        CWMat.SetPropertyByName("NUXY", 0.33, 0);
                        CWMat.SetPropertyByName("DENS", 2450, 0);
                        CWMat.SetPropertyByName("SIGYLD", 250000000, 0);
                        errCode = SolidBody.SetSolidBodyMaterial(CWMat);

                        SolidBody = null;
                        SolidComponent = null;
                    }

                    break;

                case "Ручной ввод":
                    double ex1, nuxy1, dens1, sigyld1;
                    string name1;

                    ex1 = Convert.ToDouble(textBox14.Text);
                    nuxy1 = Convert.ToDouble(textBox15.Text);
                    dens1 = Convert.ToDouble(textBox16.Text);
                    sigyld1 = Convert.ToDouble(textBox17.Text);
                    name1 = textBox30.Text;

                    for (j = 7; j <= 8; j++)
                    {
                        SolidComponent = SolidMgr.GetComponentAt(j, out errorCode);

                        SName = SolidComponent.ComponentName;

                        SolidBody = (CWSolidBody)SolidComponent.GetSolidBodyAt(0, out errCode);

                        CWMat = (CWMaterial)SolidBody.GetDefaultMaterial();
                        CWMat.MaterialUnits = 0;

                        CWMat.MaterialName = name1;
                        CWMat.SetPropertyByName("EX", ex1, 0);
                        CWMat.SetPropertyByName("NUXY", nuxy1, 0);
                        CWMat.SetPropertyByName("DENS", dens1, 0);
                        CWMat.SetPropertyByName("SIGYLD", sigyld1, 0);

                        errCode = SolidBody.SetSolidBodyMaterial(CWMat);

                        SolidBody = null;
                        SolidComponent = null;
                    }

                    break;
            }


            switch (comboBox2.Text)
            {
                case "Сталь":


                    SolidComponent = SolidMgr.GetComponentAt(6, out errorCode);

                    SName = SolidComponent.ComponentName;


                    SolidBody = (CWSolidBody)SolidComponent.GetSolidBodyAt(0, out errCode);

                    CWMat = (CWMaterial)SolidBody.GetDefaultMaterial();
                    CWMat.MaterialUnits = 0;

                    CWMat.MaterialName = "Сталь";
                    CWMat.SetPropertyByName("EX", 210000000000.0, 0);
                    CWMat.SetPropertyByName("NUXY", 0.3, 0);
                    CWMat.SetPropertyByName("DENS", 7800, 0);
                    CWMat.SetPropertyByName("SIGYLD", 600000000, 0);
                    errCode = SolidBody.SetSolidBodyMaterial(CWMat);

                    SolidBody = null;
                    SolidComponent = null;


                    break;
                case "Алюминий":


                    SolidComponent = SolidMgr.GetComponentAt(6, out errorCode);

                    SName = SolidComponent.ComponentName;


                    SolidBody = (CWSolidBody)SolidComponent.GetSolidBodyAt(0, out errCode);

                    CWMat = (CWMaterial)SolidBody.GetDefaultMaterial();
                    CWMat.MaterialUnits = 0;

                    CWMat.MaterialName = "Алюминий";
                    CWMat.SetPropertyByName("EX", 72000000000.0, 0);
                    CWMat.SetPropertyByName("NUXY", 0.33, 0);
                    CWMat.SetPropertyByName("DENS", 2450, 0);
                    CWMat.SetPropertyByName("SIGYLD", 250000000, 0);
                    errCode = SolidBody.SetSolidBodyMaterial(CWMat);

                    SolidBody = null;
                    SolidComponent = null;


                    break;

                case "Ручной ввод":
                    double ex2, nuxy2, dens2, sigyld2;
                    string name2;

                    ex2 = Convert.ToDouble(textBox18.Text);
                    nuxy2 = Convert.ToDouble(textBox19.Text);
                    dens2 = Convert.ToDouble(textBox20.Text);
                    sigyld2 = Convert.ToDouble(textBox21.Text);
                    name2 = textBox31.Text;


                    SolidComponent = SolidMgr.GetComponentAt(6, out errorCode);

                    SName = SolidComponent.ComponentName;

                    SolidBody = (CWSolidBody)SolidComponent.GetSolidBodyAt(0, out errCode);

                    CWMat = (CWMaterial)SolidBody.GetDefaultMaterial();
                    CWMat.MaterialUnits = 0;

                    CWMat.MaterialName = name2;
                    CWMat.SetPropertyByName("EX", ex2, 0);
                    CWMat.SetPropertyByName("NUXY", nuxy2, 0);
                    CWMat.SetPropertyByName("DENS", dens2, 0);
                    CWMat.SetPropertyByName("SIGYLD", sigyld2, 0);

                    errCode = SolidBody.SetSolidBodyMaterial(CWMat);

                    SolidBody = null;
                    SolidComponent = null;


                    break;
            }


            switch (comboBox3.Text)
            {
                case "Сталь":

                    for (j = 0; j <= 4; j++)
                    {
                        SolidComponent = SolidMgr.GetComponentAt(j, out errorCode);

                        SName = SolidComponent.ComponentName;


                        SolidBody = (CWSolidBody)SolidComponent.GetSolidBodyAt(0, out errCode);

                        CWMat = (CWMaterial)SolidBody.GetDefaultMaterial();
                        CWMat.MaterialUnits = 0;

                        CWMat.MaterialName = "Сталь";
                        CWMat.SetPropertyByName("EX", 210000000000.0, 0);
                        CWMat.SetPropertyByName("NUXY", 0.3, 0);
                        CWMat.SetPropertyByName("DENS", 7800, 0);
                        CWMat.SetPropertyByName("SIGYLD", 600000000, 0);
                        errCode = SolidBody.SetSolidBodyMaterial(CWMat);

                        SolidBody = null;
                        SolidComponent = null;
                    }
                    break;
                case "Алюминий":

                    for (j = 0; j <= 4; j++)
                    {
                        SolidComponent = SolidMgr.GetComponentAt(j, out errorCode);

                        SName = SolidComponent.ComponentName;


                        SolidBody = (CWSolidBody)SolidComponent.GetSolidBodyAt(0, out errCode);

                        CWMat = (CWMaterial)SolidBody.GetDefaultMaterial();
                        CWMat.MaterialUnits = 0;

                        CWMat.MaterialName = "Алюминий";
                        CWMat.SetPropertyByName("EX", 72000000000.0, 0);
                        CWMat.SetPropertyByName("NUXY", 0.33, 0);
                        CWMat.SetPropertyByName("DENS", 2450, 0);
                        CWMat.SetPropertyByName("SIGYLD", 250000000, 0);
                        errCode = SolidBody.SetSolidBodyMaterial(CWMat);

                        SolidBody = null;
                        SolidComponent = null;

                    }
                    break;

                case "Ручной ввод":
                    double ex3, nuxy3, dens3, sigyld3;
                    string name3;

                    ex3 = Convert.ToDouble(textBox22.Text);
                    nuxy3 = Convert.ToDouble(textBox23.Text);
                    dens3 = Convert.ToDouble(textBox24.Text);
                    sigyld3 = Convert.ToDouble(textBox25.Text);
                    name3 = textBox32.Text;
                    for (j = 0; j <= 4; j++)
                    {
                        SolidComponent = SolidMgr.GetComponentAt(j, out errorCode);

                        SName = SolidComponent.ComponentName;

                        SolidBody = (CWSolidBody)SolidComponent.GetSolidBodyAt(0, out errCode);

                        CWMat = (CWMaterial)SolidBody.GetDefaultMaterial();
                        CWMat.MaterialUnits = 0;

                        CWMat.MaterialName = name3;
                        CWMat.SetPropertyByName("EX", ex3, 0);
                        CWMat.SetPropertyByName("NUXY", nuxy3, 0);
                        CWMat.SetPropertyByName("DENS", dens3, 0);
                        CWMat.SetPropertyByName("SIGYLD", sigyld3, 0);

                        errCode = SolidBody.SetSolidBodyMaterial(CWMat);

                        SolidBody = null;
                        SolidComponent = null;
                    }
                    break;
            }


            switch (comboBox4.Text)
            {
                case "Сталь":
                    for (j = 5; j <= 9; j += 4)
                    {
                        SolidComponent = SolidMgr.GetComponentAt(j, out errorCode);

                        SName = SolidComponent.ComponentName;
                        SolidBody = (CWSolidBody)SolidComponent.GetSolidBodyAt(0, out errCode);

                        CWMat = (CWMaterial)SolidBody.GetDefaultMaterial();
                        CWMat.MaterialUnits = 0;

                        CWMat.MaterialName = "Сталь";
                        CWMat.SetPropertyByName("EX", 210000000000.0, 0);
                        CWMat.SetPropertyByName("NUXY", 0.3, 0);
                        CWMat.SetPropertyByName("DENS", 7800, 0);
                        CWMat.SetPropertyByName("SIGYLD", 600000000, 0);
                        errCode = SolidBody.SetSolidBodyMaterial(CWMat);

                        SolidBody = null;
                        SolidComponent = null;
                    }
                    break;
                case "Алюминий":

                    for (j = 5; j <= 9; j += 4)
                    {
                        SolidComponent = SolidMgr.GetComponentAt(j, out errorCode);
                        SName = SolidComponent.ComponentName;
                        SolidBody = (CWSolidBody)SolidComponent.GetSolidBodyAt(0, out errCode);

                        CWMat = (CWMaterial)SolidBody.GetDefaultMaterial();
                        CWMat.MaterialUnits = 0;

                        CWMat.MaterialName = "Алюминий";
                        CWMat.SetPropertyByName("EX", 72000000000.0, 0);
                        CWMat.SetPropertyByName("NUXY", 0.33, 0);
                        CWMat.SetPropertyByName("DENS", 2450, 0);
                        CWMat.SetPropertyByName("SIGYLD", 250000000, 0);
                        errCode = SolidBody.SetSolidBodyMaterial(CWMat);


                        SolidBody = null;
                        SolidComponent = null;
                    }

                    break;
                case "Ручной ввод":
                    double ex4, nuxy4, dens4, sigyld4;
                    string name4;

                    ex4 = Convert.ToDouble(textBox26.Text);
                    nuxy4 = Convert.ToDouble(textBox27.Text);
                    dens4 = Convert.ToDouble(textBox28.Text);
                    sigyld4 = Convert.ToDouble(textBox29.Text);
                    name4 = textBox33.Text;

                    for (j = 5; j <= 9; j += 4)
                    {
                        SolidComponent = SolidMgr.GetComponentAt(j, out errorCode);

                        SName = SolidComponent.ComponentName;

                        SolidBody = (CWSolidBody)SolidComponent.GetSolidBodyAt(0, out errCode);

                        CWMat = (CWMaterial)SolidBody.GetDefaultMaterial();
                        CWMat.MaterialUnits = 0;

                        CWMat.MaterialName = name4;
                        CWMat.SetPropertyByName("EX", ex4, 0);
                        CWMat.SetPropertyByName("NUXY", nuxy4, 0);
                        CWMat.SetPropertyByName("DENS", dens4, 0);
                        CWMat.SetPropertyByName("SIGYLD", sigyld4, 0);

                        errCode = SolidBody.SetSolidBodyMaterial(CWMat);

                        SolidBody = null;
                        SolidComponent = null;
                    }
                    break;

            }

            CWLoadsAndRestraintsManager LBCMgr = default(CWLoadsAndRestraintsManager);
            CWRestraint CWRes1 = default(CWRestraint);
            CWRestraint CWRes2 = default(CWRestraint);
            CWRestraint CWRes3 = default(CWRestraint);
            CWJoints Joints = default(CWJoints);
            // ModelDoc2 Part = default(ModelDoc2);

            object obj = null;
            bool boolstatus = false;
            object PID = null;
            object CWBaseExcitationEntity = null;
            object myComponent;
            object DispArray1;
            object[] DispArray2 = new object[1];
            object[] DispArray3 = new object[1];
            object[] DispArray4 = new object[1];
            object[] DispArray5 = new object[1];
            object RefGeom1;
            object RefGeom2;
            object RefGeom3;
            CWPressure CWFeatObj2 = default(CWPressure);
            CWCentriFugalForce CWCentrifugalLoad;
            CWDampingOptions DampingOptions = default(CWDampingOptions);
            CWForce CWForce;
            LBCMgr = Study.LoadsAndRestraintsManager;

            //////// нагрузка
            switch (comboBox5.Text)
            {
                case "Статический":
                    if (textBox34.Text == "")
                    {
                        textBox34.Text = "0";
                    }
                    if (textBox35.Text == "")
                    {
                        textBox35.Text = "0";
                    }

                    double av, aa;

                    CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PIDload), out errCode);
                    DispArray1 = CWBaseExcitationEntity;

                    CWCentrifugalLoad = LBCMgr.AddCentrifugalForce(DispArray1, out errCode);

                    CWCentrifugalLoad.CFForceBeginEdit();
                    switch (comboBox6.Text)
                    {
                        case "Об/мин":
                            av = Convert.ToDouble(textBox34.Text);
                            av = av / 9.5493;
                            CWCentrifugalLoad.AngularVelocity = av;
                            break;
                        case "Рад/сек":
                            av = Convert.ToDouble(textBox34.Text);
                            CWCentrifugalLoad.AngularVelocity = av;
                            break;
                    }
                    switch (comboBox7.Text)
                    {
                        case "Об/мин^2":
                            aa = Convert.ToDouble(textBox35.Text);
                            aa = aa / 572.9578;
                            CWCentrifugalLoad.AngularAcceleration = aa;
                            break;
                        case "Рад/сек^2":
                            aa = Convert.ToDouble(textBox35.Text);
                            CWCentrifugalLoad.AngularAcceleration = aa;
                            break;
                    }

                    CWCentrifugalLoad.Unit = (int)swsUnitSystem_e.swsUnitSystemSI;
                    CWCentrifugalLoad.CFForceEndEdit();

                   // errCode = CWFeatObj2.PressureEndEdit();
                    break;


                case "Гармонический":
                    if (textBox34.Text == "")
                    {
                        textBox34.Text = "0";
                    }


                    CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PIDforce1), out errCode);
                    DispArray2[0] = CWBaseExcitationEntity;
                    CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PIDforce2), out errCode);
                    DispArray1 = CWBaseExcitationEntity;
                    CWForce = LBCMgr.AddForce(0, (DispArray2), DispArray1, out errCode);
                    CWForce.ForceBeginEdit();
                    double force;
                    force = (1.9 / Convert.ToDouble(textBox34.Text)) * (1570.8 * 1570.8);
                    CWForce.SetForceComponentValues(1, 0, 0, force, 0, 0);
                    object[] curve = new object[] { 2, 0, 0, 5862.226, 1.0 };
                    CWForce.SetTimeCurve(curve, out errCode);
                    CWForce.Unit = (int)swsUnitSystem_e.swsUnitSystemSI;
                    CWForce.ForceEndEdit();

                    CWForce = LBCMgr.AddForce(0, (DispArray2), DispArray1, out errCode);
                    CWForce.ForceBeginEdit();
                    CWForce.SetForceComponentValues(0, 1, 0, 0, force, 0);
                    CWForce.SetTimeCurve(curve, out errCode);
                    CWForce.PhaseAngleUnit = 1;
                    CWForce.PhaseAngle = 90;
                    CWForce.Unit = (int)swsUnitSystem_e.swsUnitSystemSI;
                    CWForce.ForceEndEdit();

                    CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PIDforce3), out errCode);
                    DispArray2[0] = CWBaseExcitationEntity;
                    CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PIDforce4), out errCode);
                    DispArray1 = CWBaseExcitationEntity;
                    CWForce = LBCMgr.AddForce(0, (DispArray2), DispArray1, out errCode);
                    CWForce.ForceBeginEdit();
                    force = (3.5 / Convert.ToDouble(textBox34.Text)) * (1570.8 * 1570.8);
                    CWForce.SetForceComponentValues(1, 0, 0, force, 0, 0);
                    CWForce.SetTimeCurve(curve, out errCode);
                    CWForce.Unit = (int)swsUnitSystem_e.swsUnitSystemSI;
                    CWForce.ForceEndEdit();


                    CWForce = LBCMgr.AddForce(0, (DispArray2), DispArray1, out errCode);
                    CWForce.ForceBeginEdit();
                    CWForce.SetForceComponentValues(0, 1, 0, 0, force, 0);
                    CWForce.SetTimeCurve(curve, out errCode);
                    CWForce.PhaseAngleUnit = 1;
                    CWForce.PhaseAngle = 90;
                    CWForce.Unit = (int)swsUnitSystem_e.swsUnitSystemSI;
                    CWForce.ForceEndEdit();

                    /////демпфирование



                    DampingOptions = Study.DampingOptions;
                    DampingOptions.DampingType = 1;

                    double beta;
                    beta = 0.9 / (Math.PI * 225);

                    DampingOptions.SetRayleighDampingCoefficients(0.0, beta);
                    DampingOptions.ComputeFromMaterialDamping = 0;
                    break;
                case "Переходные процессы":
                    if (textBox34.Text == "")
                    {
                        textBox34.Text = "0";
                    }


                    CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PIDforce1), out errCode);
                    DispArray2[0] = CWBaseExcitationEntity;
                    CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PIDforce2), out errCode);
                    DispArray1 = CWBaseExcitationEntity;
                    CWForce = LBCMgr.AddForce(0, (DispArray2), DispArray1, out errCode);
                    CWForce.ForceBeginEdit();
                    double force1;
                    force1 = (1.9 / Convert.ToDouble(textBox34.Text)) * (1570.8 * 1570.8);
                    CWForce.SetForceComponentValues(1, 0, 0, force1, 0, 0);
                    //object[] curve1 = new object[] { 2, 0, 0, 5862.226, 1.0 };
                    object[] curve1 = new object[801];
                    curve1[0] = 200;
                    curve1[1] = 0;
                    curve1[2] = 0;
                    double t;
                    for (int i=3; i<=800; i+=2)
                    {
                        t= (0.2 * ((i - 1) / 2)) / 400;
                        curve1[i] = t;
                        curve1[i + 1] = 1 * Math.Sin(5860 * t);
                    }
                    /*object[] curve2 = new object[401];
                    curve2[0] = 200;
                    curve2[1] = 0;
                    curve2[2] = 1.0;
                    double t1;
                    for (int i = 3; i <= 400; i += 2)
                    {
                       t1 = (0.5 * ((i - 1) / 2)) / 200;
                       curve2[i] = t1;
                       curve2[i + 1] = 0.802 * Math.Sin(((5860 * t1)+90)* (Math.PI / 180));
                    }*/
                    CWForce.SetTimeCurve(curve1, out errCode);
                    CWForce.Unit = (int)swsUnitSystem_e.swsUnitSystemSI;
                    CWForce.ForceEndEdit();

                    CWForce = LBCMgr.AddForce(0, (DispArray2), DispArray1, out errCode);
                    CWForce.ForceBeginEdit();
                    CWForce.SetForceComponentValues(0, 1, 0, 0, force1, 0);
                    CWForce.SetTimeCurve(curve1, out errCode);
                    CWForce.PhaseAngleUnit = 1;
                    CWForce.PhaseAngle = 90;
                    CWForce.Unit = (int)swsUnitSystem_e.swsUnitSystemSI;
                    CWForce.ForceEndEdit();

                    CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PIDforce3), out errCode);
                    DispArray2[0] = CWBaseExcitationEntity;
                    CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PIDforce4), out errCode);
                    DispArray1 = CWBaseExcitationEntity;
                    CWForce = LBCMgr.AddForce(0, (DispArray2), DispArray1, out errCode);
                    CWForce.ForceBeginEdit();
                    force1 = (3.5 / Convert.ToDouble(textBox34.Text)) * (1570.8 * 1570.8);
                    CWForce.SetForceComponentValues(1, 0, 0, force1, 0, 0);
                    CWForce.SetTimeCurve(curve1, out errCode);
                    CWForce.Unit = (int)swsUnitSystem_e.swsUnitSystemSI;
                    CWForce.ForceEndEdit();


                    CWForce = LBCMgr.AddForce(0, (DispArray2), DispArray1, out errCode);
                    CWForce.ForceBeginEdit();
                    CWForce.SetForceComponentValues(0, 1, 0, 0, force1, 0);
                    CWForce.SetTimeCurve(curve1, out errCode);
                    CWForce.PhaseAngleUnit = 1;
                    CWForce.PhaseAngle = 90;
                    CWForce.Unit = (int)swsUnitSystem_e.swsUnitSystemSI;
                    CWForce.ForceEndEdit();

                    /////демпфирование

                    


                    DampingOptions = Study.DampingOptions;
                    DampingOptions.DampingType = 1;

                    double beta1;
                    beta1 = 0.9 / (Math.PI * 225);

                    DampingOptions.SetRayleighDampingCoefficients(0.0, beta1);
                    DampingOptions.ComputeFromMaterialDamping = 0;
                    break;
            }
        
    
            //////// закрепление 1го подш



            CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PIDpodsh11), out errCode);
            DispArray3[0] = CWBaseExcitationEntity;

            boolstatus = Part.Extension.SelectByID2("Сверху@podsh11-1@ротор", "PLANE", 0, 0, 0, false, 0, null, 0);
            obj = ((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
            PID = Part.Extension.GetPersistReference3(obj);
            CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PID), out errCode);
            RefGeom1 = CWBaseExcitationEntity;

            CWRes1 = LBCMgr.AddRestraint(0, (DispArray3), RefGeom1, out errCode);

            CWRes1.RestraintBeginEdit();
            CWRes1.SetTranslationComponentsValues(0, 0, 1, 0.0, 0.0, 0.0);
            CWRes1.SetRotationComponentsValues(0, 0, 0, 0.0, 0.0, 0.0);
            CWRes1.Unit = 2;
            errCode = CWRes1.RestraintEndEdit();

            /////////////////// закрепление 2го подш

            // boolstatus = Part.Extension.SelectByID2("", "FACE", 0.204445254567361, 6.74646597611286E-02, 5.65023865275407E-02, false, 0, null, 0);

            //obj = ((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
            // PID = Part.Extension.GetPersistReference3(obj);
            CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PIDpodsh22), out errCode);
            DispArray4[0] = CWBaseExcitationEntity;

            boolstatus = Part.Extension.SelectByID2("Сверху@podsh22-1@ротор", "PLANE", 0, 0, 0, false, 0, null, 0);
            obj = ((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
            PID = Part.Extension.GetPersistReference3(obj);
            CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PID), out errCode);
            RefGeom2 = CWBaseExcitationEntity;

            CWRes2 = LBCMgr.AddRestraint(0, (DispArray4), RefGeom2, out errCode);

            CWRes2.RestraintBeginEdit();
            CWRes2.SetTranslationComponentsValues(0, 0, 1, 0.0, 0.0, 0.0);
            CWRes2.SetRotationComponentsValues(0, 0, 0, 0.0, 0.0, 0.0);
            CWRes2.Unit = 2;
            errCode = CWRes2.RestraintEndEdit();
            ////закреп кольца////
            // boolstatus = Part.Extension.SelectByID2("", "FACE", 1.99824949561389E-04, 9.64164394427485E-02, 8.41894898735518E-02, false, 0, null, 0);
            //obj = ((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
            //PID = Part.Extension.GetPersistReference3(obj);
            CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PIDring), out errCode);
            DispArray5[0] = CWBaseExcitationEntity;

            boolstatus = Part.Extension.SelectByID2("Сверху@ring-1@ротор", "PLANE", 0, 0, 0, false, 0, null, 0);
            obj = ((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
            PID = Part.Extension.GetPersistReference3(obj);
            CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PID), out errCode);
            RefGeom3 = CWBaseExcitationEntity;

            CWRes3 = LBCMgr.AddRestraint(0, (DispArray5), RefGeom3, out errCode);

            CWRes3.RestraintBeginEdit();
            CWRes3.SetTranslationComponentsValues(0, 0, 1, 0.0, 0.0, 0.0);
            CWRes3.SetRotationComponentsValues(0, 0, 0, 0.0, 0.0, 0.0);
            CWRes3.Unit = 2;
            errCode = CWRes3.RestraintEndEdit();
            ////////////////////////////////// СЕТКА

            CWBeamManager BeamMgr = default(CWBeamManager);

            CWBeamBody BeamBody = default(CWBeamBody);


            CWMesh Mesh = default(CWMesh);

            int nbrBeamBodies = 0;

            int beamBodyType = 0;

            double ElementSize = 0;

            double Tolerance = 0;

            int p = 0;


            bool keepJointUpdates = true;


            //Get and set beam info

            BeamMgr = (CWBeamManager)Study.BeamManager;

            nbrBeamBodies = BeamMgr.BeamCount;


            BeamBody = null;

            for (p = 0; p <= (nbrBeamBodies - 1); p++)

            {

                BeamBody = (CWBeamBody)BeamMgr.GetBeamBodyAt(j, out errCode);


                beamBodyType = BeamBody.BeamType;




                BeamBody = null;

            }

            //контакты 

            CWContactComponent CWComponentContact = default(CWContactComponent);
            CWContactSet CWContactSet = default(CWContactSet);
            CWContactManager ContactMgr = default(CWContactManager);
            object cont;

            boolstatus = Part.Extension.SelectByID2("ротор.SLDASM", "COMPONENT", 0, 0, 0, false, 0, null, 0);
            obj = ((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
            PID = Part.Extension.GetPersistReference3(obj);
            CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PID), out errCode);
            cont = CWBaseExcitationEntity;

            ContactMgr = Study.ContactManager;
            ContactMgr.DeleteContactComponent("Global Contact");
            CWComponentContact = ContactMgr.CreateContactComponent((int)swsContactType_e.swsContactTypeBonded, (int)swsMeshCompatibility_e.swsMeshCompatibilityCompatible, cont, out errCode);
            // errCode = ContactMgr.SuppressUnsuppressComponentContact(CWComponentContact.ContactName, 0);
            /*
           ContactMgr = (CWContactManager)Study.ContactManager;
           ContactMgr.DeleteContactComponent("Global Contact");

           CWContactSet = (CWContactSet)ContactMgr.CreateContactSet2((int)swsContactType_e.swsContactTypeBonded, 0, cont, 0, out errCode);*/


            ////соединение/////
            /* CWBearingLoad CWBearingLoad;
             CWSpringConnector springConnector;
             CWLoadsAndRestraintsManager lrMngr;
             CWLoadsAndRestraints lr;
             object[] down = new object[1];
             object[] up = new object[1];
            // object down;
            // object up;
             object Vert1;
             object Vert2;
            // object[] varArray1 = new object[2];
             object[] dir = new object[2];


             boolstatus = Part.Extension.SelectByID2("", "FACE", -0.223152884199123, 7.75950402719445E-02, -1.94681721069969E-02, false, 0, null, 0);
             obj = ((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
             PID = Part.Extension.GetPersistReference3(obj);
             CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PID), out errCode);
             down[0] = CWBaseExcitationEntity;


             boolstatus = Part.Extension.SelectByID2("", "FACE", -0.219714728094914, 7.69317360240933E-02, -3.37269623939278E-02, false, 0, null, 0);
             obj = ((SelectionMgr)(Part.SelectionManager)).GetSelectedObject6(1, -1);
             PID = Part.Extension.GetPersistReference3(obj);
             CWBaseExcitationEntity = Part.Extension.GetObjectByPersistReference3((PID), out errCode);
             up[0] = CWBaseExcitationEntity;

             lrMngr = Study.LoadsAndRestraintsManager;

             springConnector = lrMngr.AddSpringConnector((int)swsSpringConnectorType_e.swsSpringConnectoryTypeConcentricCylindricalFaces, (down), (up), ref errCode);

             springConnector.SpringConnectorBeginEdit();
             springConnector.Unit = (int)swsUnit_e.swsUnitEnglish;
             springConnector.NormalRadialStiffnessValue = 800.0;
             springConnector.TangentialOrShearStiffnessValue = 100000.0;

             springConnector.SpringConnectorEndEdit();*/

            //Mesh the part

            Mesh = Study.Mesh;

            Mesh.Quality = (int)swsMeshQuality_e.swsMeshQualityHigh;

            Mesh.GetDefaultElementSizeAndTolerance((int)swsLinearUnit_e.swsLinearUnitMillimeters, out ElementSize, out Tolerance);
            ElementSize = 9.5;
            ElementSize = (double)trackBar1.Value / 100;

            errCode = Study.CreateMesh((int)swsLinearUnit_e.swsLinearUnitMillimeters, ElementSize, Tolerance);

            //errCode = Study.RunAnalysis();

            //заполняем комбобокс индексами анализов 
            
            CwAddincallback CWObject;
            CWObject = (CwAddincallback)SwApp.GetAddInObject("SldWorks.Simulation");
            COSMOSWORKS = CWObject.CosmosWorks;

            ActDoc = COSMOSWORKS.ActiveDoc;
            
            StudyMngr = ActDoc.StudyManager;
            int scount;
            scount = StudyMngr.StudyCount;
            if (scount != 0)
            {
                comboBox8.Items.Clear();
                for (int i = 0; i < scount; i++)
                {
                    comboBox8.Items.Add(StudyMngr.GetStudy(i).Name);
                }
            }
            FormName.Close();
        }

        /*CWMat.MaterialName = "Алюминий";
                    CWMat.SetPropertyByName("EX", 210000000000.0, 0);
                    CWMat.SetPropertyByName("NUXY", 0.28, 0);
                    CWMat.SetPropertyByName("GXY", 79000000000.0, 0);
                    CWMat.SetPropertyByName("DENS", 7700, 0);
                    CWMat.SetPropertyByName("SIGXT", 723825600, 0);
                    CWMat.SetPropertyByName("SIGYLD", 620422000, 0);
                    CWMat.SetPropertyByName("ALPX", 1.3E-05, 0);
                    CWMat.SetPropertyByName("KX", 50, 0);
                    CWMat.SetPropertyByName("C", 460, 0);
                    errCode = SolidBody.SetSolidBodyMaterial(CWMat);*/
        ////////////

        ///\/////////////////////////////////////////
        ///
        //////////////////////////

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox1.Text)
            {
                case "Сталь":
                    textBox14.Visible = false;
                    textBox15.Visible = false;
                    textBox16.Visible = false;
                    textBox17.Visible = false;
                    textBox30.Visible = false;

                    break;
                case "Алюминий":
                    textBox14.Visible = false;
                    textBox15.Visible = false;
                    textBox16.Visible = false;
                    textBox17.Visible = false;
                    textBox30.Visible = false;
                    break;

                case "Ручной ввод":
                    textBox14.Visible = true;
                    textBox15.Visible = true;
                    textBox16.Visible = true;
                    textBox17.Visible = true;
                    textBox30.Visible = true;
                    break;
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox2.Text)
            {
                case "Сталь":
                    textBox18.Visible = false;
                    textBox19.Visible = false;
                    textBox20.Visible = false;
                    textBox21.Visible = false;
                    textBox31.Visible = false;

                    break;
                case "Алюминий":
                    textBox18.Visible = false;
                    textBox19.Visible = false;
                    textBox20.Visible = false;
                    textBox21.Visible = false;
                    textBox31.Visible = false;
                    break;

                case "Ручной ввод":
                    textBox18.Visible = true;
                    textBox19.Visible = true;
                    textBox20.Visible = true;
                    textBox21.Visible = true;
                    textBox31.Visible = true;
                    break;
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox3.Text)
            {
                case "Сталь":
                    textBox22.Visible = false;
                    textBox23.Visible = false;
                    textBox24.Visible = false;
                    textBox25.Visible = false;
                    textBox32.Visible = false;

                    break;
                case "Алюминий":
                    textBox22.Visible = false;
                    textBox23.Visible = false;
                    textBox24.Visible = false;
                    textBox25.Visible = false;
                    textBox32.Visible = false;
                    break;

                case "Ручной ввод":
                    textBox22.Visible = true;
                    textBox23.Visible = true;
                    textBox24.Visible = true;
                    textBox25.Visible = true;
                    textBox32.Visible = true;
                    break;
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox4.Text)
            {
                case "Сталь":
                    textBox26.Visible = false;
                    textBox27.Visible = false;
                    textBox28.Visible = false;
                    textBox29.Visible = false;
                    textBox33.Visible = false;
                    break;
                case "Алюминий":
                    textBox26.Visible = false;
                    textBox27.Visible = false;
                    textBox28.Visible = false;
                    textBox29.Visible = false;
                    textBox33.Visible = false;

                    break;
                case "Ручной ввод":
                    textBox26.Visible = true;
                    textBox27.Visible = true;
                    textBox28.Visible = true;
                    textBox29.Visible = true;
                    textBox33.Visible = true;
                    break;
            }
        }

       
        

        private void textBox14_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox15_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox16_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox17_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox18_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox19_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox20_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox21_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox22_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox23_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox24_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox25_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox26_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox27_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox28_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox29_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox34_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }

        private void textBox35_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
               (e.KeyChar != ','))
            {
                e.Handled = true;
            }
        }
        

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            double elsize;
            trackBar1.Minimum = 500;
            trackBar1.Maximum = 950;
           
            elsize =((double)trackBar1.Value/100);
            label25.Text =Convert.ToString(elsize)+"мм";
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox5.Text)
            {
                case "Статический":
                    label27.Visible = true;
                    label26.Visible = true;
                    textBox35.Visible = true;
                    textBox34.Visible = true;
                    comboBox6.Visible = true;
                    comboBox7.Visible = true;
                    label26.Text = "Угловая скорость";
                    label26.Font = new Font("Times New Roman", 12);
                    break;
                case "Колебания":
                    label27.Visible = false;
                    textBox35.Visible = false;
                    textBox34.Visible = false;
                    comboBox6.Visible = false;
                    comboBox7.Visible = false;
                    label26.Visible = false;
                    break;
                case "Гармонический":
                    label26.Visible = true;
                    textBox34.Visible = true;
                    label27.Visible = false;
                    textBox35.Visible = false;
                    comboBox6.Visible = false;
                    comboBox7.Visible = false;
                    label26.Text = "Параметр небаланса как часть массы колес";
                    label26.Font = new Font("Times New Roman", 10);
                    break;
                case "Переходные процессы":
                    label26.Visible = true;
                    textBox34.Visible = true;
                    label27.Visible = false;
                    textBox35.Visible = false;
                    comboBox6.Visible = false;
                    comboBox7.Visible = false;
                    label26.Text = "Параметр небаланса как часть массы колес";
                    label26.Font = new Font("Times New Roman", 10);
                    break;
            }
        }
        int k = 1;

        private void button4_Click(object sender, EventArgs e)
        {
            int fileerror = SwApp.UnloadAddIn("D:\\проги\\Solid\\SOLIDWORKS\\Simulation\\cosworks.dll");
            fileerror = SwApp.LoadAddIn("D:\\проги\\Solid\\SOLIDWORKS\\Simulation\\cosworks.dll");

            if (comboBox8.SelectedIndex == -1)
            {
                comboBox8.SelectedIndex = comboBox8.Items.Count - 1;
            }

            if (pictureBox2.Image != null)
            {
                pictureBox2.Image.Dispose();
                pictureBox2.Image = null;
                pictureBox2.Load();
                pictureBox2.Update();

            }
            if (pictureBox3.Image != null)
            {
                pictureBox3.Image.Dispose();
                pictureBox3.Image = null;
                pictureBox3.Load("D:\\диплом\\photo\\qwe.png");
                pictureBox3.Update();
                System.IO.File.Delete(@"D:\\диплом\\photo\\StressPlot.jpg");

            }
            if (pictureBox4.Image != null)
            {
                pictureBox4.Image.Dispose();
                pictureBox4.Image = null;
                pictureBox4.Load("D:\\диплом\\photo\\qwe.png");
                pictureBox4.Update();
                System.IO.File.Delete(@"D:\\диплом\\photo\\DispPlot.jpg");

            }
            if (pictureBox5.Image != null)
            {
                pictureBox5.Image.Dispose();
                pictureBox5.Image = null;
                pictureBox5.Load("D:\\диплом\\photo\\qwe.png");
                pictureBox5.Update();
                System.IO.File.Delete(@"D:\\диплом\\photo\\StrainPlot.jpg");
            }
            if (pictureBox6.Image != null)
            {
                pictureBox6.Image.Dispose();
                pictureBox6.Image = null;
                pictureBox6.Load("D:\\диплом\\photo\\qwe.png");
                pictureBox6.Update();
                System.IO.File.Delete(@"D:\\диплом\\photo\\AmplitudePlot1.jpg");
            }
            if (pictureBox7.Image != null)
            {
                pictureBox7.Image.Dispose();
                pictureBox7.Image = null;
                pictureBox7.Load("D:\\диплом\\photo\\qwe.png");
                pictureBox7.Update();
                System.IO.File.Delete(@"D:\\диплом\\photo\\AmplitudePlot2.jpg");
            }
            if (pictureBox8.Image != null)
            {
                pictureBox8.Image.Dispose();
                pictureBox8.Image = null;
                pictureBox8.Load("D:\\диплом\\photo\\qwe.png");
                pictureBox8.Update();
                System.IO.File.Delete(@"D:\\диплом\\photo\\AmplitudePlot3.jpg");
            }
            if (pictureBox9.Image != null)
            {
                pictureBox9.Image.Dispose();
                pictureBox9.Image = null;
                pictureBox9.Load("D:\\диплом\\photo\\qwe.png");
                pictureBox9.Update();
                System.IO.File.Delete(@"D:\\диплом\\photo\\AmplitudePlot4.jpg");
            }
            if (pictureBox10.Image != null)
            {
                pictureBox10.Image.Dispose();
                pictureBox10.Image = null;
                pictureBox10.Load("D:\\диплом\\photo\\qwe.png");
                pictureBox10.Update();
                System.IO.File.Delete(@"D:\\диплом\\photo\\AmplitudePlot5.jpg");
            }
            if (pictureBox11.Image != null)
            {
                pictureBox11.Image.Dispose();
                pictureBox11.Image = null;
                pictureBox11.Load("D:\\диплом\\photo\\qwe.png");
                pictureBox11.Update();
                System.IO.File.Delete(@"D:\\диплом\\photo\\AmplitudePlot6.jpg");
            }
            if (pictureBox12.Image != null)
            {
                pictureBox12.Image.Dispose();
                pictureBox12.Image = null;
                pictureBox12.Load("D:\\диплом\\photo\\qwe.png");
                pictureBox12.Update();
                System.IO.File.Delete(@"D:\\диплом\\photo\\AmplitudePlot7.jpg");
            }
            if (pictureBox13.Image != null)
            {
                pictureBox13.Image.Dispose();
                pictureBox13.Image = null;
                pictureBox13.Load("D:\\диплом\\photo\\qwe.png");
                pictureBox13.Update();
                System.IO.File.Delete(@"D:\\диплом\\photo\\AmplitudePlot8.jpg");
            }
            if (pictureBox14.Image != null)
            {
                pictureBox14.Image.Dispose();
                pictureBox14.Image = null;
                pictureBox14.Load("D:\\диплом\\photo\\qwe.png");
                pictureBox14.Update();
                System.IO.File.Delete(@"D:\\диплом\\photo\\AmplitudePlot9.jpg");
            }
            if (pictureBox15.Image != null)
            {
                pictureBox15.Image.Dispose();
                pictureBox15.Image = null;
                pictureBox15.Load("D:\\диплом\\photo\\qwe.png");
                pictureBox15.Update();
                System.IO.File.Delete(@"D:\\диплом\\photo\\AmplitudePlot10.jpg");
            }
            swModel = SwApp.ActiveDoc;


            double[] props = swModel.GetMassProperties();
            

            CosmosWorks COSMOSWORKS;
            CwAddincallback CWObject;
            CWModelDoc ActDoc;
            CWStudyManager StudyMngr;
            CWStudy Study;
            CWResults CWResult;
            int errCode;

            CWObject = (CwAddincallback)SwApp.GetAddInObject("SldWorks.Simulation");
            COSMOSWORKS = CWObject.CosmosWorks;

            ActDoc = COSMOSWORKS.ActiveDoc;
            
            
            StudyMngr = ActDoc.StudyManager;
            
            StudyMngr.ActiveStudy =comboBox8.SelectedIndex;
            Study = StudyMngr.GetStudy(StudyMngr.ActiveStudy);
            int Atype = (int)Study.AnalysisType;
            
            CWResult = (CWResults)Study.Results;
            
            
            if (Atype == 1)
            {
                tabPage5.Parent = null;
                tabPage6.Parent = null;
                tabPage7.Parent = null;
                tabPage8.Parent = tabControl2;
                tabPage9.Parent = tabControl2;
                tabPage10.Parent = tabControl2;
                tabPage11.Parent = tabControl2;
                tabPage12.Parent = tabControl2;
                tabPage13.Parent = tabControl2;
                tabPage14.Parent = tabControl2;
                tabPage15.Parent = tabControl2;
                tabPage16.Parent = tabControl2;
                tabPage17.Parent = tabControl2;
                object[] frequency;
                frequency= (object[])CWResult.GetResonantFrequencies(out errCode);
                textBox37.Text = "________Список резонансных частот________";
                textBox37.Text += System.Environment.NewLine;
                textBox37.Text += System.Environment.NewLine + "|Режим No. |Частотный(Рад/сек) |Частотный(Герц) |Период(Секунды) |";
                for (int i = 0; i < frequency.Length; i+=4)
                {
                    string No = Convert.ToString(frequency[i]);
                    int j = No.Length;
                    while (j < 10) {
                        No += "_";
                        j = No.Length;
                    }
                    string freq1 = Convert.ToString(frequency[i + 1]);
                    j = freq1.Length;
                    while (j < 18)
                    {
                        freq1 += "_";
                        j = freq1.Length;
                    }
                    string freq2 = Convert.ToString(frequency[i + 2]);
                    j = freq2.Length;
                    while (j < 15)
                    {
                        freq2 += "_";
                        j = freq2.Length;
                    }
                    string period = Convert.ToString(frequency[i + 3]);
                    j = period.Length;
                    while (j < 16)
                    {
                        period += "_";
                        j = period.Length;
                    }
                    textBox37.Text += System.Environment.NewLine +"|"+ No + "|" + freq1 + "|" + freq2 + "|" + period + "|";
                }
                //AmplitudePlot
                
                string plot;
                CWResult.ActivatePlot("Амплитуда1");
                swModel.SaveAs("D:\\диплом\\photo\\AmplitudePlot1.jpg");
                plot = "D:\\диплом\\photo\\AmplitudePlot1.jpg";
                pictureBox6.Load(plot);

                CWResult.ActivatePlot("Амплитуда2");
                swModel.SaveAs("D:\\диплом\\photo\\AmplitudePlot2.jpg");
                plot = "D:\\диплом\\photo\\AmplitudePlot2.jpg";
                pictureBox7.Load(plot);

                CWResult.ActivatePlot("Амплитуда3");
                swModel.SaveAs("D:\\диплом\\photo\\AmplitudePlot3.jpg");
                plot = "D:\\диплом\\photo\\AmplitudePlot3.jpg";
                pictureBox8.Load(plot);

                CWResult.ActivatePlot("Амплитуда4");
                swModel.SaveAs("D:\\диплом\\photo\\AmplitudePlot4.jpg");
                plot = "D:\\диплом\\photo\\AmplitudePlot4.jpg";
                pictureBox9.Load(plot);

                CWResult.ActivatePlot("Амплитуда5");
                swModel.SaveAs("D:\\диплом\\photo\\AmplitudePlot5.jpg");
                plot = "D:\\диплом\\photo\\AmplitudePlot5.jpg";
                pictureBox10.Load(plot);

                CWResult.ActivatePlot("Амплитуда6");
                swModel.SaveAs("D:\\диплом\\photo\\AmplitudePlot6.jpg");
                plot = "D:\\диплом\\photo\\AmplitudePlot6.jpg";
                pictureBox11.Load(plot);

                CWResult.ActivatePlot("Амплитуда7");
                swModel.SaveAs("D:\\диплом\\photo\\AmplitudePlot7.jpg");
                plot = "D:\\диплом\\photo\\AmplitudePlot7.jpg";
                pictureBox12.Load(plot);

                CWResult.ActivatePlot("Амплитуда8");
                swModel.SaveAs("D:\\диплом\\photo\\AmplitudePlot8.jpg");
                plot = "D:\\диплом\\photo\\AmplitudePlot8.jpg";
                pictureBox13.Load(plot);

                CWResult.ActivatePlot("Амплитуда9");
                swModel.SaveAs("D:\\диплом\\photo\\AmplitudePlot9.jpg");
                plot = "D:\\диплом\\photo\\AmplitudePlot9.jpg";
                pictureBox14.Load(plot);

                CWResult.ActivatePlot("Амплитуда10");
                swModel.SaveAs("D:\\диплом\\photo\\AmplitudePlot10.jpg");
                plot = "D:\\диплом\\photo\\AmplitudePlot10.jpg";
                pictureBox15.Load(plot);

            }
            else
                {
                object[] stress;
                object[] strains;
                object[] disp;

                
                stress = (object[])CWResult.GetMinMaxStress(9, 0, 1, null, 0, out errCode);

                disp = (object[])CWResult.GetMinMaxDisplacement(3, 1, null, 0, out errCode);

                strains = (object[])CWResult.GetMinMaxStrain(6, 1, 1, null, out errCode);


                textBox37.Text = "Напряжение:";
                textBox37.Text += System.Environment.NewLine + "  Максимальное: " + stress[3] + "N/m^2";
                textBox37.Text += System.Environment.NewLine + "  Минимальное: " + stress[1] + "N/m^2";
                textBox37.Text += System.Environment.NewLine + "Перемещение:";
                textBox37.Text += System.Environment.NewLine + "  Максимальное: " + disp[3] + "mm";
                textBox37.Text += System.Environment.NewLine + "  Минимальное: " + disp[1] + "mm";
                textBox37.Text += System.Environment.NewLine + "Деформация:";
                textBox37.Text += System.Environment.NewLine + "  Максимальное: " + strains[3];
                textBox37.Text += System.Environment.NewLine + "  Минимальное: " + strains[1];
                //textBox37.Text += System.Environment.NewLine + CWResult.GetPlotNames()[0];
                string plot;
                tabPage5.Parent = tabControl2;
                tabPage6.Parent = tabControl2;
                tabPage7.Parent = tabControl2;
                tabPage8.Parent = null;
                tabPage9.Parent = null;
                tabPage10.Parent = null;
                tabPage11.Parent = null;
                tabPage12.Parent = null;
                tabPage13.Parent = null;
                tabPage14.Parent = null;
                tabPage15.Parent = null;
                tabPage16.Parent = null;
                tabPage17.Parent = null;
                //StressPlot
                //CWResult.CreatePlot((int)swsPlotResultTypes_e.swsResultStress, (int)swsStaticResultNodalStressComponentTypes_e.swsStaticNodalStress_VON, (int)swsUnit_e.swsUnitSI, false, out errCode);
                CWResult.ActivatePlot("Напряжение1");
                swModel.SaveAs("D:\\диплом\\photo\\StressPlot.jpg");
                plot = "D:\\диплом\\photo\\StressPlot.jpg";
                pictureBox3.Load(plot);
                
                //DispPlot
                //CWResult.CreatePlot((int)swsPlotResultTypes_e.swsResultDisplacementOrAmplitude, (int)swsDisplacementComponent_e.swsDisplacementComponentURES, 3, false, out errCode);
                CWResult.ActivatePlot("Перемещение1");
                swModel.SaveAs("D:\\диплом\\photo\\DispPlot.jpg");
                plot = "D:\\диплом\\photo\\DispPlot.jpg";
                pictureBox4.Load(plot);
                //StrainPlot
                //CWResult.CreatePlot((int)swsPlotResultTypes_e.swsResultStrain, (int)swsStaticResultElementalStrainComponentTypes_e.swsStaticElementalStrain_ESTRN, (int)swsUnit_e.swsUnitSI, false, out errCode);
                CWResult.ActivatePlot("Деформация1");
                swModel.SaveAs("D:\\диплом\\photo\\StrainPlot.jpg");
                plot = "D:\\диплом\\photo\\StrainPlot.jpg";
                pictureBox5.Load(plot);
                


            }



            /* IRayTraceRenderer swRayTraceRenderer = SwApp.GetRayTraceRenderer(1);
             if (swRayTraceRenderer == null)
             {

                 int fileerror = SwApp.LoadAddIn("PhotoView 360");
                 swRayTraceRenderer = SwApp.IGetRayTraceRenderer(1);
             }

             RayTraceRendererOptions swRayTraceRenderOptions = swRayTraceRenderer.RayTraceRendererOptions;





             string pict;
             pict = "./lter_" + k + ".jpg";
             bool status = swRayTraceRenderer.RenderToFile(Application.StartupPath + pict, 0, 0);
             status = swRayTraceRenderer.CloseRayTraceRender();

             pictureBox2.Load(Application.StartupPath + pict);
             */
            k++;
        }

        
    }
}
