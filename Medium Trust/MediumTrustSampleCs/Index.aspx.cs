using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web.UI;
using System.Web.UI.WebControls;
using GemBox.Document;

namespace MediumTrustSampleCs
{
    public partial class _Default : System.Web.UI.Page
    {
        private static int invoiceNumber = 0;

        protected void Page_Load(object sender, EventArgs e)
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            ComponentInfo.FreeLimitReached += (s1, e1) => e1.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

            // By specifying a location that is under ASP.NET application's control, 
            // GemBox.Document can use file system operations to retrieve font data even in Medium Trust environment.
            FontSettings.FontsBaseDirectory = Server.MapPath("Fonts/");

            if (!Page.IsPostBack)
            {
                // Set default values.
                this.LblInvoiceNumber.Text = (++invoiceNumber).ToString();
                this.TbxInvoiceDate.Text = DateTime.Now.ToShortDateString();

                // Fill grid with some dummy data.
                DataTable dt = new DataTable();
                dt.Columns.Add("Date", typeof(DateTime));
                dt.Columns.Add("Hours", typeof(int));
                dt.Columns.Add("Price", typeof(double));
                dt.Columns.Add("Total", typeof(double));

                dt.Rows.Add(DateTime.Now.AddDays(-2), 7, 35);
                dt.Rows.Add(DateTime.Now.AddDays(-1), 8, 35);

                Session["DataTable"] = dt;

                this.SetDataBinding();
            }
        }

        // On button click generate the document and stream it to the browser.
        protected void BtnGenerate_Click(object sender, EventArgs e)
        {
            DocumentModel document = GenerateDocument();

            string fileName = "Report." + this.RdbOutputFormat.SelectedValue;

            document.Save(this.Response, fileName);
        }
                
        private DocumentModel GenerateDocument()
        {
            DateTime invoiceDate;
            if (!DateTime.TryParse(this.TbxInvoiceDate.Text, out invoiceDate))
                this.TbxInvoiceDate.Text = (invoiceDate = DateTime.Now).ToShortDateString();

            return Process((DataTable)Session["DataTable"], int.Parse(this.LblInvoiceNumber.Text),
                        invoiceDate, this.TbxCustomerName.Text, this.TbxCustomerAddress.Text, this.DdlCustomerCountry.SelectedValue, this.TbxContactPerson.Text, this.TbxNotes.Text);
        }

        // Modify document with the data from web form using Mail Merge.
        // This method also shows how to modify document from different data sources.
        public DocumentModel Process(
            DataTable dt, int invoiceNumber, DateTime invoiceDate, string customerName, string companyAddress, string country, string contantactPerson, string notes)
        {
            string path = Path.Combine(Request.PhysicalApplicationPath, "Invoice.docx");

            // Load template document.
            DocumentModel document = DocumentModel.Load(path);

            // Subscribe to FieldMerging event (we want to format the output).
            document.MailMerge.FieldMerging += (sender, e) =>
            {
                if (e.IsValueFound)
                    switch (e.FieldName)
                    {
                        case "Date":
                            ((Run)e.Inline).Text = ((DateTime)e.Value).ToString("dddd, MMMM d, yyyy");
                            break;
                        case "Price":
                        case "Total":
                        case "TotalPrice":
                            ((Run)e.Inline).Text = ((double)e.Value).ToString("C");
                            break;
                    }
            };

            // Fill table.
            document.MailMerge.Execute(dt, "Item");

            // Fill invoice data.
            document.MailMerge.Execute(
                new
                {
                    Number = invoiceNumber,
                    InvoiceDate = invoiceDate.ToShortDateString()
                });


            // Fill customer data.
            document.MailMerge.Execute(
                new Dictionary<string, object>()
                { 
                    { "Name", customerName },
                    { "Address" , companyAddress },
                    { "Country" , country },
                    { "ContactPerson" , contantactPerson }
                });

            // Fill total.
            document.MailMerge.Execute(new { TotalPrice = dt.Rows.Cast<DataRow>().Sum(row => (double)row["Total"]) });

            // Fill notes.
            document.MailMerge.Execute(
                new KeyValuePair<string, object>[]
                {
                    new KeyValuePair<string, object>("Notes", notes)
                });

            return document;
        }
        
        private void SetDataBinding()
        {
            DataTable data = (DataTable)Session["DataTable"];
            double total = 0;
            foreach (DataRow row in data.Rows)
            {
                double totalRow = (int)row["Hours"] * (double)row["Price"];
                row["Total"] = totalRow;
                total += totalRow;
            }

            DataView dataView = data.DefaultView;

            this.GridView1.DataSource = dataView;
            dataView.AllowDelete = true;
            this.GridView1.DataBind();

            this.LblTotal.Text = "Total: " + total.ToString("C");
        }

        protected void GridView1_RowDeleting(object sender, GridViewDeleteEventArgs e)
        {
            DataTable people = (DataTable)Session["DataTable"];

            people.Rows[e.RowIndex].Delete();
            this.SetDataBinding();
        }

        protected void GridView1_RowEditing(object sender, GridViewEditEventArgs e)
        {
            this.GridView1.EditIndex = e.NewEditIndex;
            this.SetDataBinding();
        }

        protected void GridView1_RowUpdating(object sender, GridViewUpdateEventArgs e)
        {
            int i;
            int rowIndex = e.RowIndex;
            DataTable dt = (DataTable)Session["DataTable"];

            for (i = 1; i < dt.Columns.Count; i++)
            {
                var editTextBox = this.GridView1.Rows[rowIndex].Cells[i].Controls[0] as System.Web.UI.WebControls.TextBox;

                if (editTextBox != null)
                {
                    try
                    {
                        dt.Rows[rowIndex][i - 1] = editTextBox.Text;
                    }
                    catch (ArgumentException)
                    {                        
                    }
                }
            }

            this.GridView1.EditIndex = -1;
            this.SetDataBinding();
        }

        protected void GridView1_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
            this.GridView1.EditIndex = -1;
            this.SetDataBinding();
        }

        protected void BtnAddItem_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)Session["DataTable"];
            dt.Rows.Add(DateTime.Now, 8, 35);
            this.SetDataBinding();
        }        
    }
}