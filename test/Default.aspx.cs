using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Net;
using System.Web.UI;
using System.Web.UI.WebControls;


namespace test
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                PopulateData();
                LblMessage.Text = "Current Database data!";
            }

        }

        private void PopulateData()
        {
            using (testDBEntities t = new testDBEntities())
            {
                GvData.DataSource = t.Grupo.ToList();
                GvData.DataBind();
            }
        }

        public void ImportExcel()
        {
            if (FileUpload1.PostedFile.ContentType == "application/vnd.ms-excel" || FileUpload1.PostedFile.ContentType == "application/vdn.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                try
                {
                    string fileName = Path.Combine(Server.MapPath("~/ImportDocument"), Guid.NewGuid().ToString() + Path.GetExtension(FileUpload1.PostedFile.FileName));
                    FileUpload1.PostedFile.SaveAs(fileName);


                    string ext = Path.GetExtension(FileUpload1.PostedFile.FileName);
                    string conString = "";

                    if (ext.ToLower() == ".xls")
                    {
                        conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties =\"Excel 8.0;HDR=Yes;IMEX=2\"";
                    }
                    else if (ext.ToLower() == ".xlsx")
                    {
                        conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties =\"Excel 12.0;HDR=Yes;IMEX=2\"";
                    }

                    string query = "Select [Alumno ID], [Nombres], [Apellido Materno], [Apellido Paterno], [Fecha Nacimiento], [Grado], [Grupo1], [Calificacion]";
                    OleDbConnection con = new OleDbConnection(conString);
                    if (con.State == System.Data.ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    OleDbCommand cmd = new OleDbCommand(query, con);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);

                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    da.Dispose();
                    con.Close();
                    con.Dispose();

                    //pasando a bd
                    using (testDBEntities dc = new testDBEntities())
                    {
                        foreach (DataRow dr in ds.Tables[0].Rows)
                        {
                            string aluId = dr["Alumno ID"].ToString();
                            var v = dc.Grupo.Where(a => a.AlumnoId.Equals(aluId)).FirstOrDefault();
                            if (v != null)
                            {
                                //update
                                v.Nombres = dr["Nombres"].ToString();
                                v.ApellidoMaterno = dr["Apellido Materno"].ToString();
                                v.ApellidoPaterno = dr["Apellido Paterno"].ToString();
                                v.FechaNacimiento = dr["Fecha Nacimiento"].ToString();
                                v.Grupo1 = dr["Grupo1"].ToString();
                                string grdo = dr["Grado"].ToString();
                                string grades = dr["Calificacion"].ToString();
                            }
                            else
                            {
                                //insert
                                dc.Grupo.Add(new Grupo
                                {

                                    AlumnoId = Convert.ToInt32(dr["Alumno ID"]),
                                    Nombres = dr["Nombres"].ToString(),
                                    ApellidoMaterno = dr["Apellido Materno"].ToString(),
                                    ApellidoPaterno = dr["Apellido Paterno"].ToString(),
                                    FechaNacimiento = dr["Fecha Nacimiento"].ToString(),
                                    Grado = Convert.ToInt32(dr["Grado"]),
                                    Grupo1 = dr["Grupo"].ToString(),
                                    Calificacion = Convert.ToInt32(dr["Calificacion"])

                                });

                            }
                        }

                        dc.SaveChanges();
                    }
                    PopulateData();
                    LblMessage.Text = "Se importo correctamente el documento!";
                }
                catch (Exception)
                {
                    throw;
                }



            }
        }
        protected void btnImport_Click(object sender, EventArgs e)
        {

            ImportExcel();
        }


        protected void BtnChart_Click(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                PopulateDataReport();
            }

        }

        private void PopulateDataReport()
        {
            using(testDBEntities dc = new testDBEntities())
            {
                var v = dc.Grupo.ToList();

                GvDataGrades.DataSource = v;
                GvDataGrades.DataBind();

                Chart1.DataSource = v;
                Chart1.DataBind();

            }
        }

        protected void BtnMax_Click(object sender, EventArgs e)
        {
            try
            {
                
                string conString = "";
                using (OleDbConnection conn = new OleDbConnection(conString))
                {
                    conn.Open();
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = conn;
                    OleDbDataAdapter objDA = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", conn);
                    DataSet ds = new DataSet();
                    objDA.Fill(ds);

                    using (testDBEntities dc = new testDBEntities())
                    {
                        foreach (DataRow dr in ds.Tables[0].Rows)
                        {
                            string grades = dr["Califiacion"].ToString();
                            string sheetName = dr["Grupo"].ToString();
                            if (!sheetName.EndsWith("$"))
                                continue;

                            cmd.CommandText = "SELECT Nombres, ApellidoMaterno, ApellidoPaterno FROM [" + sheetName + "] WHERE MAX(grades)";
                        }
                    }
                }
                

                
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}