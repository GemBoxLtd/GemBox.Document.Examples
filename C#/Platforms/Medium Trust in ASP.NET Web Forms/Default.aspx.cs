using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web.UI;
using System.Web.UI.WebControls;
using GemBox.Document;

namespace MediumTrust
{
    public partial class Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            // If using the Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            // In order to create a PDF file in Medium Trust environment you'll need to specify
            // font files location that is under your ASP.NET application's control.
            // This is required because in partial-trust domain it's prohibited to access system's installed fonts.
            FontSettings.FontsBaseDirectory = Server.MapPath("Fonts/");

            // Set web form default values.
            if (!this.Page.IsPostBack)
            {
                this.txtNumber.Text = "10203";
                this.txtDate.Text = DateTime.Today.ToString("yyyy-MM-dd");
                this.txtCompany.Text = "ACME Corp";

                var items = new DataTable("Items");
                items.Columns.Add("Date", typeof(DateTime));
                items.Columns.Add("Hours", typeof(int));
                items.Columns.Add("Unit", typeof(double));
                items.Columns.Add("Price", typeof(double));

                items.Rows.Add(DateTime.Today.AddDays(-3), 6, 35);
                items.Rows.Add(DateTime.Today.AddDays(-2), 7, 35);
                items.Rows.Add(DateTime.Today.AddDays(-1), 8, 35);

                this.Session["ItemsTable"] = items;
                this.RefreshGridData();
            }
        }

        private DataTable GetGridData() => (DataTable)this.Session["ItemsTable"];

        private void RefreshGridData()
        {
            var items = this.GetGridData();

            double total = 0;
            foreach (DataRow row in items.Rows)
            {
                double price = (int)row["Hours"] * (double)row["Unit"];
                row["Price"] = price;
                total += price;
            }

            this.gridItems.DataSource = items;
            this.gridItems.DataBind();
            this.txtTotal.Text = total.ToString("0.00");
        }

        protected void btnCreateFile_Click(object sender, EventArgs e)
        {
            // Load template document.
            DocumentModel document = DocumentModel.Load(
                Path.Combine(this.Request.PhysicalApplicationPath, "InvoiceTemplate.docx"));

            if (!int.TryParse(this.txtNumber.Text, out int number))
                this.txtNumber.Text = default(int).ToString();

            if (!DateTime.TryParse(this.txtDate.Text, out DateTime date))
                this.txtDate.Text = default(DateTime).ToString();

            // Execute mail merge operations with data from web form.
            document.MailMerge.Execute(
                new Dictionary<string, object>()
                {
                    ["Number"] = number,
                    ["Date"] = date,
                    ["Company"] = this.txtCompany.Text,
                    ["TotalPrice"] = double.Parse(this.txtTotal.Text)
                });
            document.MailMerge.Execute(
                this.GetGridData());

            // Create document in specified format (PDF, DOCX, etc.) and
            // stream (download) it to client's browser.
            document.Save(this.Response, $"Invoice{this.ddlFileFormat.SelectedValue}");
        }

        protected void btnAddRow_Click(object sender, EventArgs e)
        {
            var items = this.GetGridData();
            items.Rows.Add(DateTime.Today, 8, 35);
            this.RefreshGridData();

        }

        protected void gridItems_RowUpdating(object sender, GridViewUpdateEventArgs e)
        {
            var items = this.GetGridData();

            int rowIndex = e.RowIndex;
            for (int columnIndex = 1; columnIndex < items.Columns.Count; columnIndex++)
            {
                var cell = this.gridItems.Rows[rowIndex].Cells[columnIndex];

                if (cell.Controls[0] is System.Web.UI.WebControls.TextBox textBox)
                {
                    try { items.Rows[rowIndex][columnIndex - 1] = textBox.Text; }
                    catch (ArgumentException) { }
                }
            }

            this.gridItems.EditIndex = -1;
            this.RefreshGridData();
        }

        protected void gridItems_RowEditing(object sender, GridViewEditEventArgs e)
        {
            this.gridItems.EditIndex = e.NewEditIndex;
            this.RefreshGridData();
        }

        protected void gridItems_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
            this.gridItems.EditIndex = -1;
            this.RefreshGridData();
        }

        protected void gridItems_RowDeleting(object sender, GridViewDeleteEventArgs e)
        {
            var items = this.GetGridData();
            items.Rows[e.RowIndex].Delete();
            this.RefreshGridData();
        }
    }
}