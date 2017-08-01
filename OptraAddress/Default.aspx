<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="OptaAddress.Default" %>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
    
    <title>Opta Address Challenge</title>

    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <meta name="description" content="Optra Address Challenge" />
    <meta http-equiv='cache-control' content='no-cache' />
    <meta http-equiv='expires' content='0' />
    <meta http-equiv='pragma' content='no-cache' />
    <meta http-equiv="X-UA-Compatible" content="IE=9; IE=8; IE=EDGE" />
    
</head>
<body>
    <form id="form1" runat="server">
        <div id="mainBody" role="navigation" aria-labelledby="main body">
            <h4>Select a file to upload:</h4>

            <asp:FileUpload id="FileUpload1" runat="server" Width="500px"></asp:FileUpload>
            <br /><br />
            <asp:Button id="UploadButton" Text="Upload file" OnClick="UploadButton_Click" runat="server"></asp:Button>      
            <hr />
            <asp:Label id="UploadStatusLabel" runat="server"></asp:Label>   

            <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" EmptyDataText = "No files uploaded">
                <Columns>
                    <asp:BoundField DataField="Text" HeaderText="File Name" />
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:LinkButton ID="lnkDownload" Text = "View File Contents" CommandArgument = '<%# Eval("Value") %>' runat="server" OnClick = "DownloadFile"></asp:LinkButton>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:LinkButton ID = "lnkDelete" Text = "Delete" CommandArgument = '<%# Eval("Value") %>' runat = "server" OnClick = "DeleteFile" />
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
            </asp:GridView>
        </div>
    </form>
</body>
</html>
