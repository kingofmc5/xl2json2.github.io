<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="xl2json.aspx.cs" Inherits="jsonapp.xl2json" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>

    <script type="text/javascript">
    function copyToClipboard() {
        // Select the text inside the TextBox

        

        var textBox = document.getElementById('<%= txt_output.ClientID %>');
        if (textBox.value.trim() == "") {
            alert("json Value/Textbox is empty !");
            return;
        }
        textBox.select();

        try {
            // Copy the selected text to the clipboard
            document.execCommand('copy');
            alert('Text copied to clipboard!');
        } catch (err) {
            alert('Unable to copy text to clipboard. Please copy it manually.');
        }
        return false;
    }
    </script>


    <form id="form1" runat="server">
        <div style="min-height: 100%">


            <asp:FileUpload ID="FileUpload1" runat="server" />
            <asp:Button ID="Button1" runat="server" Text="get json" OnClick="upload_click" /><br />
            <div style="min-height: 100px" ></div>
            <asp:Button ID="btnCopy" runat="server" Text="Copy to Clipboard" OnClientClick="return copyToClipboard();" />
            <asp:TextBox ID ="txt_output" runat="server" TextMode="MultiLine" Width="95%" Height="600px" ></asp:TextBox>
            

        </div>
    </form>
</body>
</html>
