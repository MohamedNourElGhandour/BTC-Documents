using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Btc;

namespace BTC_Documents
{
    public partial class frmDocuments : Form
    {
        public frmDocuments()
        {
            InitializeComponent();
        }
        Document d;
     

     

        private void ToolStripContactDatabase_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\contaFromDB.doc";
            oopen(PathOn);
        }

      
       

        private void ToolStripBookDelivry_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\BookRec.doc";
            oopen(PathOn);
        }


        private void oopen(string path)
        {
           
            try
            {
                d = new Document();
                d.OpenDocument(path);
            }
            catch (Exception y)
            {
                MessageBox.Show(y.Message);
            }
        }

        private void ToolStripRecDemand_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\EmpDem.doc";
            oopen(PathOn);
        }

        private void ToolStripWorkFollow_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\FollowWork.doc";
            oopen(PathOn);
        }

      

        private void ToolStripinstructor_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\TrainerAssei.doc";
            oopen(PathOn);
        }

        private void ToolStripEmployee_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\EmpAssei.doc";
            oopen(PathOn);
        }

        private void ToolStripnewEmp_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\empAssiAftr.doc";
            oopen(PathOn);
        }

        private void ToolStripEmpnewWork_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\EmpAftertra.doc";
            oopen(PathOn);
        }

        private void ToolStripVacation_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\VacForm.doc";
            oopen(PathOn);
        }

        private void ToolStripdemission_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\GoForm.doc";
            oopen(PathOn);
        }

        private void ToolStripAdmDici_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\No1.doc";
            oopen(PathOn);
        }

       

        private void ToolStripcompuIns_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\CompIns.doc";
            oopen(PathOn);
        }

        private void ToolStripEnglishIns_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\EngIns.doc";
            oopen(PathOn);
        }

        private void ToolStripExcuDir_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\ExcuDir.doc";
            oopen(PathOn);
        }

        private void ToolStripAdminDirec_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\AdminDir.doc";
            oopen(PathOn);
        }

        private void ToolStripBranchDirec_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\BranchDir.doc";
            oopen(PathOn);
        }

        private void ToolStripMarkEmp_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\MarketEmp.doc";
            oopen(PathOn);
        }

        private void ToolStripRegEmp_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\RegEmp.doc";
            oopen(PathOn);
        }

        private void ToolStripStorEmp_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\StorEmp.doc";
            oopen(PathOn);
        }

        private void ToolStripAccountant_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\Accountant.doc";
            oopen(PathOn);
        }

        private void ToolStripPriceList_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\PricePr.doc";
            oopen(PathOn);
        }

        private void ToolStripReciveobjects_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\AmanaRec.doc";
            oopen(PathOn);
        }

        






        
      

        private void ToolStripConnect_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\ContacForm.doc";
            oopen(PathOn);
        }

        

        private void المجموعاتToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\Groups.xls";
            oopen(PathOn);
        }

        private void ToolStripGroupForm_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\GroupForm.xls";
            oopen(PathOn);
        }
          
               

        private void ToolStripcontactCustomer_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\CallCus.doc";
                oopen(PathOn);
        }

        private void ToolStripTrainQues_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\AskCus.doc";
            oopen(PathOn);
        }

        private void ToolStripcerDem_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\demCer.doc";
            oopen(PathOn);
        }

        private void ToolStripDelay_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\Delay.doc";
            oopen(PathOn);

        }

        private void toolStripNoabsent_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\NoAbsent.doc";
            oopen(PathOn);
        }

        private void ToolStripGradeReport_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\TraiValue.doc";
            oopen(PathOn);
        }

        private void ToolStripEfada_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\Efada.doc";
            oopen(PathOn);
        }

        private void ToolStripReturn_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\retr.xls";
            oopen(PathOn);
        }

        private void ToolStripChange_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\Convertpay.doc";
            oopen(PathOn);

        }

        private void ToolStripMaterial_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\Mat.doc";
            oopen(PathOn);
        }

        private void ToolStripBuy_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\DemBuy.doc";
            oopen(PathOn);

        }

        private void ToolStripDailyReport_Click(object sender, EventArgs e)
        {
            string k = Application.StartupPath + "\\DailyReport.doc";
            oopen(k);
         
        }

        private void ToolStripContract_Click(object sender, EventArgs e)
        {

            string PathOn = Application.StartupPath + @"\contvalue.doc";
            oopen(PathOn);

        }

        private void ToolStripReRecode_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\Re.doc";
            oopen(PathOn);
        }

        private void toolStriNego_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\Nego.doc";
            oopen(PathOn);
        }

        private void toolStripComplainForm_Click(object sender, EventArgs e)
        {
            string PathOn = Application.StartupPath + @"\Comp.doc";
            oopen(PathOn);
        }

     
      
      

      
        

      
    }

}
