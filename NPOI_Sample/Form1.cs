using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.XSSF;

using System.Windows.Forms;

using System.IO;
namespace NPOI_Sample
{
    public partial class Form1 : Form
    {
        
    

        #region Definitions
        public const string Fn_IDIOM="idioms_2011_20190329.xls";


        #endregion

        #region Properties



        #endregion

        #region Methods - Initialization
        
        public Form1()
        {
            InitializeComponent();
            DataSet ds=new DataSet();
             string strFn= Application.StartupPath+"\\"+ Fn_IDIOM;
            GetData(strFn,ref ds);
            dataGridView1.DataSource= ds.Tables[0];
            
        }



        #endregion

        #region Methods - Other support function

        public Int32 GetData(string fn,ref DataSet dsDest)
        {
            Int32 retCode=0;
            try
            {
                string strFn=fn;
                bool blFileExist=false;
               
                blFileExist=File.Exists(strFn);
                if (!blFileExist)
                {
                    blFileExist=false;
                    MessageBox.Show("檔案["+strFn +"]不存在");
                
                }

                if(blFileExist)
                {
                    FileStream fs= new FileStream(strFn, FileMode.Open);

                    HSSFWorkbook ws= new HSSFWorkbook(fs);


                    var sheet= ws.GetSheet("工作表");

                    DataSet ds= new DataSet();
                    ds.Tables.Add("newTable");
                    for(int i=0;i<=sheet.GetRow(0).LastCellNum;i++)
                    {
                        string strColName="";
                        if(sheet.GetRow(0).GetCell(i)!=null)
                        {
                            strColName= sheet.GetRow(0).GetCell(i).ToString();
                        }
                    

                        if(ds.Tables[0].Columns.Contains(strColName))
                        {
                            strColName= strColName+"_1";
                        }
                        ds.Tables[0].Columns.Add(strColName);
                    }
                
                

                    for(int i=1;i<=sheet.LastRowNum;i++)
                    {
                    
                        DataRow dr= ds.Tables[0].NewRow();
                        for( int j=0;j<=sheet.GetRow(0).LastCellNum;j++)
                        {
                        
                            if(sheet.GetRow(i).GetCell(j)!=null)
                            {
                               dr[j]=sheet.GetRow(i).GetCell(j).ToString();
                            }else
                            {
                                dr[j]="";
                            }
                        
                        
                            //
                        }
                        ds.Tables[0].Rows.Add(dr);
                    }
                    //BindingSource bs= new BindingSource();
                    //bs.DataSource= ds.Tables[0];
                    
                    dsDest=ds;
                }
            }
            catch (Exception ex)
            {
                retCode=-1;
            }

            return retCode;
        }


        #endregion

        #region Events - Form



        #endregion

        #region Events - Button



        #endregion

        #region Events - ComboBox



        #endregion

        #region Events - TextBox



        #endregion

        #region Events - CheckBox



        #endregion

        #region Events - RadioButton



        #endregion

        #region Events - Timer



        #endregion

        #region Events - DatagGidView



        #endregion

        #region Events - GroupBox



        #endregion

        #region Events - ToolStrip



        #endregion

        #region Events - MenuStrip



        #endregion



    }
}
