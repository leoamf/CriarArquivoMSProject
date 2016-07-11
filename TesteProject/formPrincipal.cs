 
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MSProject = NetOffice.MSProjectApi;
using NetOffice.MSProjectApi.Enums;
namespace TesteProject
{
    public partial class formPrincipal : Form
    {
        public formPrincipal()
        {
            InitializeComponent();
        }
    
        private void Form1_Load(object sender, EventArgs e)
        {
            MSProject.Application application = null;
            DateTime startTime = DateTime.Now;
            try
            {
                application = new MSProject.Application();
                MSProject.Project newProject = application.Projects.Add();
                for (int x = 1; x < 10; x++)
                { 
                    MSProject.Task task  = newProject.Tasks.Add("task" + x);
                    task.Start = startTime.AddDays(x);
                    task.Duration = x/2 +"d";
                    task.ResourceNames = "Resource " + x;
                    if(x>1)
                        task.Predecessors = (x-1).ToString();
  
                }
                if (null != application)
                {
                    application.FileSaveAs(@"C:\valendo2.mpp");
                    application.Quit(PjSaveType.pjSave);
                    application.Dispose();
                }
            }
            catch(Exception ex)
            {
                if (null != application)
                { 
                    application.Quit(PjSaveType.pjDoNotSave);
                    application.Dispose();
                }
          
            }   
       
        }
    }
}
