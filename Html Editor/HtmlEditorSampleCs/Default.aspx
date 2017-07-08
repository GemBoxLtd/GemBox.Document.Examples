<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="HtmlEditorSampleCs.Default" %>

<%@ Register Assembly="CKEditor.NET" Namespace="CKEditor.NET" TagPrefix="CKEditor" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>HTML Editor Sample</title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <CKEditor:CKEditorControl ID="htmlEditor" runat="server">
                Hello &lt;b&gt;World!&lt;/b&gt;
            </CKEditor:CKEditorControl>
        </div>
        <div>
            <asp:Button ID="exportButton" runat="server" Text="Export" OnClick="OnExportButtonClicked" />
            <asp:Literal ID="Literal1" Text=" to " runat="server" />
            <asp:DropDownList ID="outputFormatList" runat="server">
                <asp:ListItem Value="docx">DOCX</asp:ListItem>
                <asp:ListItem Value="html">HTML</asp:ListItem>
                <asp:ListItem Value="mht">MHTML</asp:ListItem>
                <asp:ListItem Value="rtf">RTF</asp:ListItem>
                <asp:ListItem Value="txt">TXT</asp:ListItem>
                <asp:ListItem Selected="True" Value="pdf">PDF</asp:ListItem>
                <asp:ListItem Value="xps">XPS</asp:ListItem>
                <asp:ListItem Value="png">PNG</asp:ListItem>
                <asp:ListItem Value="jpg">JPEG</asp:ListItem>
                <asp:ListItem Value="gif">GIF</asp:ListItem>
                <asp:ListItem Value="bmp">BMP</asp:ListItem>
                <asp:ListItem Value="tif">TIFF</asp:ListItem>
                <asp:ListItem Value="wdp">WMP</asp:ListItem>
            </asp:DropDownList>
        </div>
    </form>
</body>
</html>
