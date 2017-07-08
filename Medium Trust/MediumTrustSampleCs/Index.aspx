<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Index.aspx.cs" Inherits="MediumTrustSampleCs._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .style1
        {
            width: 103px;
        }
        .style2
        {
            width: 103px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <h2>Invoice:</h2>
        
        <table>
            <tr>
                <td class="style1">Number:</td>
                <td>
                    <asp:Label ID="LblInvoiceNumber" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="style1">Date:</td>
                <td>
                    <asp:TextBox ID="TbxInvoiceDate" runat="server" ></asp:TextBox>
                </td>
            </tr>
        </table>
        <br />
        <table>
            <tr>
                <td class="style2">Name:</td>
                <td>
                    <asp:TextBox ID="TbxCustomerName" runat="server">ACME Corp</asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style2">Address:</td>
                <td>
                    <asp:TextBox ID="TbxCustomerAddress" runat="server">240 Old Country Road, Springfield, IL</asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style2">Country:</td>
                <td>
                    <asp:DropDownList ID="DdlCustomerCountry" runat="server" Width="142px">
                        <asp:ListItem>AFGHANISTAN</asp:ListItem>
                        <asp:ListItem>ÅLAND ISLANDS</asp:ListItem>
                        <asp:ListItem>ALBANIA</asp:ListItem>
                        <asp:ListItem>ALGERIA</asp:ListItem>
                        <asp:ListItem>AMERICAN SAMOA</asp:ListItem>
                        <asp:ListItem>ANDORRA</asp:ListItem>
                        <asp:ListItem>ANGOLA</asp:ListItem>
                        <asp:ListItem>ANGUILLA</asp:ListItem>
                        <asp:ListItem>ANTARCTICA</asp:ListItem>
                        <asp:ListItem>ANTIGUA AND BARBUDA</asp:ListItem>
                        <asp:ListItem>ARGENTINA</asp:ListItem>
                        <asp:ListItem>ARMENIA</asp:ListItem>
                        <asp:ListItem>ARUBA</asp:ListItem>
                        <asp:ListItem>AUSTRALIA</asp:ListItem>
                        <asp:ListItem>AUSTRIA</asp:ListItem>
                        <asp:ListItem>AZERBAIJAN</asp:ListItem>
                        <asp:ListItem>BAHAMAS</asp:ListItem>
                        <asp:ListItem>BAHRAIN</asp:ListItem>
                        <asp:ListItem>BANGLADESH</asp:ListItem>
                        <asp:ListItem>BARBADOS</asp:ListItem>
                        <asp:ListItem>BELARUS</asp:ListItem>
                        <asp:ListItem>BELGIUM</asp:ListItem>
                        <asp:ListItem>BELIZE</asp:ListItem>
                        <asp:ListItem>BENIN</asp:ListItem>
                        <asp:ListItem>BERMUDA</asp:ListItem>
                        <asp:ListItem>BHUTAN</asp:ListItem>
                        <asp:ListItem>BOLIVIA, PLURINATIONAL STATE OF</asp:ListItem>
                        <asp:ListItem>BONAIRE, SINT EUSTATIUS AND SABA</asp:ListItem>
                        <asp:ListItem>BOSNIA AND HERZEGOVINA</asp:ListItem>
                        <asp:ListItem>BOTSWANA</asp:ListItem>
                        <asp:ListItem>BOUVET ISLAND</asp:ListItem>
                        <asp:ListItem>BRAZIL</asp:ListItem>
                        <asp:ListItem>BRITISH INDIAN OCEAN TERRITORY</asp:ListItem>
                        <asp:ListItem>BRUNEI DARUSSALAM</asp:ListItem>
                        <asp:ListItem>BULGARIA</asp:ListItem>
                        <asp:ListItem>BURKINA FASO</asp:ListItem>
                        <asp:ListItem>BURUNDI</asp:ListItem>
                        <asp:ListItem>CAMBODIA</asp:ListItem>
                        <asp:ListItem>CAMEROON</asp:ListItem>
                        <asp:ListItem>CANADA</asp:ListItem>
                        <asp:ListItem>CAPE VERDE</asp:ListItem>
                        <asp:ListItem>CAYMAN ISLANDS</asp:ListItem>
                        <asp:ListItem>CENTRAL AFRICAN REPUBLIC</asp:ListItem>
                        <asp:ListItem>CHAD</asp:ListItem>
                        <asp:ListItem>CHILE</asp:ListItem>
                        <asp:ListItem>CHINA</asp:ListItem>
                        <asp:ListItem>CHRISTMAS ISLAND</asp:ListItem>
                        <asp:ListItem>COCOS (KEELING) ISLANDS</asp:ListItem>
                        <asp:ListItem>COLOMBIA</asp:ListItem>
                        <asp:ListItem>COMOROS</asp:ListItem>
                        <asp:ListItem>CONGO</asp:ListItem>
                        <asp:ListItem>CONGO, THE DEMOCRATIC REPUBLIC OF THE</asp:ListItem>
                        <asp:ListItem>COOK ISLANDS</asp:ListItem>
                        <asp:ListItem>COSTA RICA</asp:ListItem>
                        <asp:ListItem>CÔTE D'IVOIRE</asp:ListItem>
                        <asp:ListItem>CROATIA</asp:ListItem>
                        <asp:ListItem>CUBA</asp:ListItem>
                        <asp:ListItem>CURAÇAO</asp:ListItem>
                        <asp:ListItem>CYPRUS</asp:ListItem>
                        <asp:ListItem>CZECH REPUBLIC</asp:ListItem>
                        <asp:ListItem>DENMARK</asp:ListItem>
                        <asp:ListItem>DJIBOUTI</asp:ListItem>
                        <asp:ListItem>DOMINICA</asp:ListItem>
                        <asp:ListItem>DOMINICAN REPUBLIC</asp:ListItem>
                        <asp:ListItem>ECUADOR</asp:ListItem>
                        <asp:ListItem>EGYPT</asp:ListItem>
                        <asp:ListItem>EL SALVADOR</asp:ListItem>
                        <asp:ListItem>EQUATORIAL GUINEA</asp:ListItem>
                        <asp:ListItem>ERITREA</asp:ListItem>
                        <asp:ListItem>ESTONIA</asp:ListItem>
                        <asp:ListItem>ETHIOPIA</asp:ListItem>
                        <asp:ListItem>FALKLAND ISLANDS (MALVINAS)</asp:ListItem>
                        <asp:ListItem>FAROE ISLANDS</asp:ListItem>
                        <asp:ListItem>FIJI</asp:ListItem>
                        <asp:ListItem>FINLAND</asp:ListItem>
                        <asp:ListItem>FRANCE</asp:ListItem>
                        <asp:ListItem>FRENCH GUIANA</asp:ListItem>
                        <asp:ListItem>FRENCH POLYNESIA</asp:ListItem>
                        <asp:ListItem>FRENCH SOUTHERN TERRITORIES</asp:ListItem>
                        <asp:ListItem>GABON</asp:ListItem>
                        <asp:ListItem>GAMBIA</asp:ListItem>
                        <asp:ListItem>GEORGIA</asp:ListItem>
                        <asp:ListItem>GERMANY</asp:ListItem>
                        <asp:ListItem>GHANA</asp:ListItem>
                        <asp:ListItem>GIBRALTAR</asp:ListItem>
                        <asp:ListItem>GREECE</asp:ListItem>
                        <asp:ListItem>GREENLAND</asp:ListItem>
                        <asp:ListItem>GRENADA</asp:ListItem>
                        <asp:ListItem>GUADELOUPE</asp:ListItem>
                        <asp:ListItem>GUAM</asp:ListItem>
                        <asp:ListItem>GUATEMALA</asp:ListItem>
                        <asp:ListItem>GUERNSEY</asp:ListItem>
                        <asp:ListItem>GUINEA</asp:ListItem>
                        <asp:ListItem>GUINEA-BISSAU</asp:ListItem>
                        <asp:ListItem>GUYANA</asp:ListItem>
                        <asp:ListItem>HAITI</asp:ListItem>
                        <asp:ListItem>HEARD ISLAND AND MCDONALD ISLANDS</asp:ListItem>
                        <asp:ListItem>HOLY SEE (VATICAN CITY STATE)</asp:ListItem>
                        <asp:ListItem>HONDURAS</asp:ListItem>
                        <asp:ListItem>HONG KONG</asp:ListItem>
                        <asp:ListItem>HUNGARY</asp:ListItem>
                        <asp:ListItem>ICELAND</asp:ListItem>
                        <asp:ListItem>INDIA</asp:ListItem>
                        <asp:ListItem>INDONESIA</asp:ListItem>
                        <asp:ListItem>IRAN, ISLAMIC REPUBLIC OF</asp:ListItem>
                        <asp:ListItem>IRAQ</asp:ListItem>
                        <asp:ListItem>IRELAND</asp:ListItem>
                        <asp:ListItem>ISLE OF MAN</asp:ListItem>
                        <asp:ListItem>ISRAEL</asp:ListItem>
                        <asp:ListItem>ITALY</asp:ListItem>
                        <asp:ListItem>JAMAICA</asp:ListItem>
                        <asp:ListItem>JAPAN</asp:ListItem>
                        <asp:ListItem>JERSEY</asp:ListItem>
                        <asp:ListItem>JORDAN</asp:ListItem>
                        <asp:ListItem>KAZAKHSTAN</asp:ListItem>
                        <asp:ListItem>KENYA</asp:ListItem>
                        <asp:ListItem>KIRIBATI</asp:ListItem>
                        <asp:ListItem>KOREA, DEMOCRATIC PEOPLE'S REPUBLIC OF</asp:ListItem>
                        <asp:ListItem>KOREA, REPUBLIC OF</asp:ListItem>
                        <asp:ListItem>KUWAIT</asp:ListItem>
                        <asp:ListItem>KYRGYZSTAN</asp:ListItem>
                        <asp:ListItem>LAO PEOPLE'S DEMOCRATIC REPUBLIC</asp:ListItem>
                        <asp:ListItem>LATVIA</asp:ListItem>
                        <asp:ListItem>LEBANON</asp:ListItem>
                        <asp:ListItem>LESOTHO</asp:ListItem>
                        <asp:ListItem>LIBERIA</asp:ListItem>
                        <asp:ListItem>LIBYA</asp:ListItem>
                        <asp:ListItem>LIECHTENSTEIN</asp:ListItem>
                        <asp:ListItem>LITHUANIA</asp:ListItem>
                        <asp:ListItem>LUXEMBOURG</asp:ListItem>
                        <asp:ListItem>MACAO</asp:ListItem>
                        <asp:ListItem>MACEDONIA, THE FORMER YUGOSLAV REPUBLIC OF</asp:ListItem>
                        <asp:ListItem>MADAGASCAR</asp:ListItem>
                        <asp:ListItem>MALAWI</asp:ListItem>
                        <asp:ListItem>MALAYSIA</asp:ListItem>
                        <asp:ListItem>MALDIVES</asp:ListItem>
                        <asp:ListItem>MALI</asp:ListItem>
                        <asp:ListItem>MALTA</asp:ListItem>
                        <asp:ListItem>MARSHALL ISLANDS</asp:ListItem>
                        <asp:ListItem>MARTINIQUE</asp:ListItem>
                        <asp:ListItem>MAURITANIA</asp:ListItem>
                        <asp:ListItem>MAURITIUS</asp:ListItem>
                        <asp:ListItem>MAYOTTE</asp:ListItem>
                        <asp:ListItem>MEXICO</asp:ListItem>
                        <asp:ListItem>MICRONESIA, FEDERATED STATES OF</asp:ListItem>
                        <asp:ListItem>MOLDOVA, REPUBLIC OF</asp:ListItem>
                        <asp:ListItem>MONACO</asp:ListItem>
                        <asp:ListItem>MONGOLIA</asp:ListItem>
                        <asp:ListItem>MONTENEGRO</asp:ListItem>
                        <asp:ListItem>MONTSERRAT</asp:ListItem>
                        <asp:ListItem>MOROCCO</asp:ListItem>
                        <asp:ListItem>MOZAMBIQUE</asp:ListItem>
                        <asp:ListItem>MYANMAR</asp:ListItem>
                        <asp:ListItem>NAMIBIA</asp:ListItem>
                        <asp:ListItem>NAURU</asp:ListItem>
                        <asp:ListItem>NEPAL</asp:ListItem>
                        <asp:ListItem>NETHERLANDS</asp:ListItem>
                        <asp:ListItem>NEW CALEDONIA</asp:ListItem>
                        <asp:ListItem>NEW ZEALAND</asp:ListItem>
                        <asp:ListItem>NICARAGUA</asp:ListItem>
                        <asp:ListItem>NIGER</asp:ListItem>
                        <asp:ListItem>NIGERIA</asp:ListItem>
                        <asp:ListItem>NIUE</asp:ListItem>
                        <asp:ListItem>NORFOLK ISLAND</asp:ListItem>
                        <asp:ListItem>NORTHERN MARIANA ISLANDS</asp:ListItem>
                        <asp:ListItem>NORWAY</asp:ListItem>
                        <asp:ListItem>OMAN</asp:ListItem>
                        <asp:ListItem>PAKISTAN</asp:ListItem>
                        <asp:ListItem>PALAU</asp:ListItem>
                        <asp:ListItem>PALESTINIAN TERRITORY, OCCUPIED</asp:ListItem>
                        <asp:ListItem>PANAMA</asp:ListItem>
                        <asp:ListItem>PAPUA NEW GUINEA</asp:ListItem>
                        <asp:ListItem>PARAGUAY</asp:ListItem>
                        <asp:ListItem>PERU</asp:ListItem>
                        <asp:ListItem>PHILIPPINES</asp:ListItem>
                        <asp:ListItem>PITCAIRN</asp:ListItem>
                        <asp:ListItem>POLAND</asp:ListItem>
                        <asp:ListItem>PORTUGAL</asp:ListItem>
                        <asp:ListItem>PUERTO RICO</asp:ListItem>
                        <asp:ListItem>QATAR</asp:ListItem>
                        <asp:ListItem>RÉUNION</asp:ListItem>
                        <asp:ListItem>ROMANIA</asp:ListItem>
                        <asp:ListItem>RUSSIAN FEDERATION</asp:ListItem>
                        <asp:ListItem>RWANDA</asp:ListItem>
                        <asp:ListItem>SAINT BARTHÉLEMY</asp:ListItem>
                        <asp:ListItem>SAINT HELENA, ASCENSION AND TRISTAN DA CUNHA</asp:ListItem>
                        <asp:ListItem>SAINT KITTS AND NEVIS</asp:ListItem>
                        <asp:ListItem>SAINT LUCIA</asp:ListItem>
                        <asp:ListItem>SAINT MARTIN (FRENCH PART)</asp:ListItem>
                        <asp:ListItem>SAINT PIERRE AND MIQUELON</asp:ListItem>
                        <asp:ListItem>SAINT VINCENT AND THE GRENADINES</asp:ListItem>
                        <asp:ListItem>SAMOA</asp:ListItem>
                        <asp:ListItem>SAN MARINO</asp:ListItem>
                        <asp:ListItem>SAO TOME AND PRINCIPE</asp:ListItem>
                        <asp:ListItem>SAUDI ARABIA</asp:ListItem>
                        <asp:ListItem>SENEGAL</asp:ListItem>
                        <asp:ListItem>SERBIA</asp:ListItem>
                        <asp:ListItem>SEYCHELLES</asp:ListItem>
                        <asp:ListItem>SIERRA LEONE</asp:ListItem>
                        <asp:ListItem>SINGAPORE</asp:ListItem>
                        <asp:ListItem>SINT MAARTEN (DUTCH PART)</asp:ListItem>
                        <asp:ListItem>SLOVAKIA</asp:ListItem>
                        <asp:ListItem>SLOVENIA</asp:ListItem>
                        <asp:ListItem>SOLOMON ISLANDS</asp:ListItem>
                        <asp:ListItem>SOMALIA</asp:ListItem>
                        <asp:ListItem>SOUTH AFRICA</asp:ListItem>
                        <asp:ListItem>SOUTH GEORGIA AND THE SOUTH SANDWICH ISLANDS</asp:ListItem>
                        <asp:ListItem>SOUTH SUDAN</asp:ListItem>
                        <asp:ListItem>SPAIN</asp:ListItem>
                        <asp:ListItem>SRI LANKA</asp:ListItem>
                        <asp:ListItem>SUDAN</asp:ListItem>
                        <asp:ListItem>SURINAME</asp:ListItem>
                        <asp:ListItem>SVALBARD AND JAN MAYEN</asp:ListItem>
                        <asp:ListItem>SWAZILAND</asp:ListItem>
                        <asp:ListItem>SWEDEN</asp:ListItem>
                        <asp:ListItem>SWITZERLAND</asp:ListItem>
                        <asp:ListItem>SYRIAN ARAB REPUBLIC</asp:ListItem>
                        <asp:ListItem>TAIWAN, PROVINCE OF CHINA</asp:ListItem>
                        <asp:ListItem>TAJIKISTAN</asp:ListItem>
                        <asp:ListItem>TANZANIA, UNITED REPUBLIC OF</asp:ListItem>
                        <asp:ListItem>THAILAND</asp:ListItem>
                        <asp:ListItem>TIMOR-LESTE</asp:ListItem>
                        <asp:ListItem>TOGO</asp:ListItem>
                        <asp:ListItem>TOKELAU</asp:ListItem>
                        <asp:ListItem>TONGA</asp:ListItem>
                        <asp:ListItem>TRINIDAD AND TOBAGO</asp:ListItem>
                        <asp:ListItem>TUNISIA</asp:ListItem>
                        <asp:ListItem>TURKEY</asp:ListItem>
                        <asp:ListItem>TURKMENISTAN</asp:ListItem>
                        <asp:ListItem>TURKS AND CAICOS ISLANDS</asp:ListItem>
                        <asp:ListItem>TUVALU</asp:ListItem>
                        <asp:ListItem>UGANDA</asp:ListItem>
                        <asp:ListItem>UKRAINE</asp:ListItem>
                        <asp:ListItem>UNITED ARAB EMIRATES</asp:ListItem>
                        <asp:ListItem>UNITED KINGDOM</asp:ListItem>
                        <asp:ListItem Selected="True">UNITED STATES</asp:ListItem>
                        <asp:ListItem>UNITED STATES MINOR OUTLYING ISLANDS</asp:ListItem>
                        <asp:ListItem>URUGUAY</asp:ListItem>
                        <asp:ListItem>UZBEKISTAN</asp:ListItem>
                        <asp:ListItem>VANUATU</asp:ListItem>
                        <asp:ListItem>VENEZUELA, BOLIVARIAN REPUBLIC OF</asp:ListItem>
                        <asp:ListItem>VIET NAM</asp:ListItem>
                        <asp:ListItem>VIRGIN ISLANDS, BRITISH</asp:ListItem>
                        <asp:ListItem>VIRGIN ISLANDS, U.S.</asp:ListItem>
                        <asp:ListItem>WALLIS AND FUTUNA</asp:ListItem>
                        <asp:ListItem>WESTERN SAHARA</asp:ListItem>
                        <asp:ListItem>YEMEN</asp:ListItem>
                        <asp:ListItem>ZAMBIA</asp:ListItem>
                        <asp:ListItem>ZIMBABWE</asp:ListItem>                       
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="style2">Contact person:</td>
                <td>
                    <asp:TextBox ID="TbxContactPerson" runat="server">Joe Smith</asp:TextBox>
                </td>
            </tr>
        </table>
        <br />
        <table>
        <tr>
            <td colspan="2">
                <asp:GridView ID="GridView1" runat="server" BackColor="#DEBA84" BorderColor="#DEBA84" 
                    BorderStyle="None" BorderWidth="1px" CellPadding="3" CellSpacing="2" AutoGenerateDeleteButton="True" AutoGenerateEditButton="True" 
                    OnRowCancelingEdit="GridView1_RowCancelingEdit" OnRowDeleting="GridView1_RowDeleting" OnRowEditing="GridView1_RowEditing" OnRowUpdating="GridView1_RowUpdating" AutoGenerateColumns="false" >
                    <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
                    <RowStyle BackColor="#FFF7E7" ForeColor="#8C4510" />
                    <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
                    <HeaderStyle BackColor="#D54A06" Font-Bold="True" ForeColor="White"/>
                    <Columns>
                        <asp:BoundField DataField="Date" HeaderText="Date"  DataFormatString="{0:D}" />
                        <asp:BoundField DataField="Hours" HeaderText="Work hours" DataFormatString="{0:D}" ItemStyle-HorizontalAlign="Right"/>
                        <asp:BoundField DataField="Price" HeaderText="Unit price (USD/hour)" DataFormatString="{0:C}" ItemStyle-HorizontalAlign="Right"/>
                        <asp:BoundField DataField="Total" HeaderText="Total" ReadOnly="true" DataFormatString="{0:C}" HeaderStyle-Width="100px" ItemStyle-HorizontalAlign="Right"/>
                    </Columns>
                </asp:GridView>
            </td>
        </tr>
        <tr>
        <td><asp:Button ID="BtnAddItem" runat="server" Text="Add row" onclick="BtnAddItem_Click" /></td>
        <td align="right">
            <asp:Label ID="LblTotal" runat="server"></asp:Label></td>
        </tr>        
        </table>
        <br />
        Notes: <br />
        <asp:TextBox ID="TbxNotes" runat="server" TextMode="MultiLine" Width="517px">Payment via check.</asp:TextBox>        
    </div>
    <br />
    <asp:RadioButtonList ID="RdbOutputFormat" runat="server">
        <asp:ListItem Value="pdf" Selected="True">PDF</asp:ListItem>
        <asp:ListItem Value="docx">DOCX</asp:ListItem>
        <asp:ListItem Value="html">HTML</asp:ListItem>
        <asp:ListItem Value="mht">MHTML</asp:ListItem>
        <asp:ListItem Value="rtf">RTF</asp:ListItem>
    </asp:RadioButtonList>
    <br />
    <asp:Button ID="BtnGenerate" runat="server" onclick="BtnGenerate_Click" Text="Generate Invoice"  />
    </form>
</body>

</html>
