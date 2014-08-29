<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="CrossDomainCRUDWeb.Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script type="text/javascript" src="../Scripts/jquery-1.9.1.min.js"></script>
    <script type="text/javascript" src="../Scripts/knockout-3.2.0.js"></script>

    <script type="text/javascript" src="../Scripts/csom.lists.js"></script>
    <script type="text/javascript" src="../Scripts/csom.listitems.js"></script>
    <script type="text/javascript" src="../Scripts/csom.viewmodel.js"></script>
    <script type="text/javascript" src="../Scripts/App.js"></script>
</head>
<body>
    <form id="form1" runat="server">
        <div id="chrome_ctrl_placeholder"></div>
        <div>
            <div id="loadingDiv">Building lists. This shouldn't take long...</div>
            <div id="mainDiv" style="width: 100%">
                <div style="width: 100%">
                    <input id="inputText" type="text" autofocus placeholder="Text" />
                    <input type="button" id="newItemButton" value="Add Item" class="app-button" />
                    <p id="message"></p>
                </div>
                <div style="width: 100%">
                    <table id="announcementsTable">
                        <caption>Announcements</caption>
                        <thead>
                            <tr>
                                <th>Title</th>
                            </tr>
                        </thead>
                        <tbody data-bind="foreach: get_items()">
                            <tr>
                                <td data-bind="text: get_title()"></td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </form>
</body>
</html>
