VERSION 5.00
Begin {17016CEE-E118-11D0-94B8-00A0C91110ED} Directory 
   ClientHeight    =   2445
   ClientLeft      =   12510
   ClientTop       =   11670
   ClientWidth     =   5100
   _ExtentX        =   8996
   _ExtentY        =   4313
   MajorVersion    =   0
   MinorVersion    =   8
   StateManagementType=   1
   ASPFileName     =   "D:\National Western Life\VB_Source\Employee Directory\Directory.ASP"
   DIID_WebClass   =   "{12CBA1F6-9056-11D1-8544-00A024A55AB0}"
   DIID_WebClassEvents=   "{12CBA1F5-9056-11D1-8544-00A024A55AB0}"
   TypeInfoCookie  =   92
   BeginProperty WebItems {193556CD-4486-11D1-9C70-00C04FB987DF} 
      WebItemCount    =   2
      BeginProperty WebItem1 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "Search"
         DISPID          =   1282
         Template        =   ""
         Token           =   "WC@"
         DIID_WebItemEvents=   "{49694DDC-2868-11D5-B4D4-00C04F73D48A}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   0   'False
         OriginalTemplate=   ""
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   0
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
      BeginProperty WebItem2 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "SearchPage"
         DISPID          =   1283
         Template        =   ""
         Token           =   "WC@"
         DIID_WebItemEvents=   "{49694E84-2868-11D5-B4D4-00C04F73D48A}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   0   'False
         OriginalTemplate=   ""
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   0
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
   EndProperty
   NameInURL       =   "Directory"
End
Attribute VB_Name = "Directory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Application References

' Microsoft ActiveX Data Objects 2.0 library
' Microsoft Internet Controls

' This is an update to my prier version "IIS SQL Webpage Generator" located on PSC

' In order for this application to work you must meet these prerequisites:
' - Viual Basic 6.0, Professional or Enterprise Edition
' - Windows NT Server 4.0 with Internet Information Server 3.0 or higher, or
' - Windows NT Workstation 4.0 with Peer Web Services installed, or
' - Windows 95/98 with Personal Web Server installed

' Create a database named "EmpExt" and give it the following columns:
' FirstName, varchar
' LastName, varchar
' DeskNumber, varchar
' Extension, varchar
' email, varchar
' Department, varchar
' I included a screenshot of the "Design Table" with the zip package for a reference.
' Add a few complete records to the database.

' The program is connected to the database through ODBC:
' In your control panel open "Data Sources (ODBC)" and click the "System DSN" Tab.
' Click the "Add.." button then select your database driver, In my case I used SQL.
' Name the reference to the data source "DeskRef" and follow the wizard until you
' have a successful connection to the "EmpExt" database.

' Once you've gotten this far your ready to run the app
' a dialog will open in a tab called "Debugging" and the option "Start component: "
' selected. The will also be a check box "Use existing browser" checked. Leave
' everything as is and select "OK"
' In the addressbar there will be a path ending with "/Directory.asp"
' add the following to the end of that path "?wci=searchpage", So the results
' might look like this - "http://localhost/EmployeeDirectory/Directory.ASP?wci=searchpage"


Option Explicit
Option Compare Text

Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim db As ADODB.Connection

Dim First_Name As String
Dim Last_Name As String
Dim Desk_No As String
Dim Ext As String
Dim Emp_Email As String
Dim Emp_Dept As String

Private Sub Search_Respond()
    
    ' Variables
    Dim SearchType As String
    Dim SearchText As String
    
    ' Variable values
    SearchType = Request.QueryString("SearchType")
    SearchText = Request.QueryString("SearchText")
    
    ConnectToDB
    
    ' Search and order by selected field
    Select Case SearchType
        Case "FirstName"
        
            adoPrimaryRS.Open "select * from EmpExt where FirstName Like '" & SearchText & "%' Order By FirstName", db, adOpenStatic, adLockOptimistic
            mbDataChanged = False
            
        Case "LastName"
            
            adoPrimaryRS.Open "select * from EmpExt where LastName Like '" & SearchText & "%' Order By LastName", db, adOpenStatic, adLockOptimistic
            mbDataChanged = False
            
        Case "DeskNumber"
        
            adoPrimaryRS.Open "select * from EmpExt where DeskNumber Like '%" & SearchText & "%' Order By DeskNumber", db, adOpenStatic, adLockOptimistic
            mbDataChanged = False
            
        Case "Extension"
        
            adoPrimaryRS.Open "select * from EmpExt where Extension Like '" & SearchText & "%' Order By Extension", db, adOpenStatic, adLockOptimistic
            mbDataChanged = False
                        
        Case "Email"
        
            adoPrimaryRS.Open "select * from EmpExt where email Like '" & SearchText & "%' Order By email", db, adOpenStatic, adLockOptimistic
            mbDataChanged = False
                                    
        Case "Department"
        
            adoPrimaryRS.Open "select * from EmpExt where Department Like '" & SearchText & "%' Order By Department", db, adOpenStatic, adLockOptimistic
            mbDataChanged = False
            
        Case "Profile"
            
            Last_Name = Request.QueryString("LastName")
            adoPrimaryRS.Open "select * from EmpExt where FirstName = '" & SearchText & "' and LastName = '" & Last_Name & "'", db, adOpenStatic, adLockOptimistic
            mbDataChanged = False

    End Select

    StartHTML
    
    AddSearchForm
    
    ' Check if SearchType is Profile
    If Last_Name > "" Then
        Response.Write "Query = " & SearchText & " " & Last_Name
    Else
        Response.Write "Query = " & SearchText
    End If
    
    BuildTableHeadings
    
    ' Display all records found or exit
    If adoPrimaryRS.EOF = True Then
        Response.Write "<b>No Records Found</b>"
        db.Close
        With Response
            .Write "    </table>"
            .Write "</div>"
        End With
        
        EndHTML
        Exit Sub
    Else
        Do While Not adoPrimaryRS.EOF

            First_Name = adoPrimaryRS.Fields(0).Value
            Last_Name = adoPrimaryRS.Fields(1).Value
            Desk_No = adoPrimaryRS.Fields(2).Value
            Ext = adoPrimaryRS.Fields(3).Value
            Emp_Email = adoPrimaryRS.Fields(4).Value
            Emp_Dept = adoPrimaryRS.Fields(5).Value

            With Response
                .Write "    <tr>"
                ' Create link to Employee Profile database - later version
                ' This line opens the link in a new window
                .Write "        <td width=""200"" height=""1""><font size=""2""><a href=""directory.asp?wci=Search&SearchType=Profile&Searchtext=" & First_Name & "&LastName=" & Last_Name & """Target=""_Blank"">" & First_Name & " " & Last_Name & "</a></font></td>"
                ' This line opens the link in the same window
'                .Write "        <td width=""200"" height=""1""><font size=""2""><a href=""directory.asp?wci=Search&SearchType=Profile&Searchtext=" & First_Name & "&LastName=" & Last_Name & """>" & First_Name & " " & Last_Name & "</a></font></td>"
                .Write "        <td width=""100"" height=""1""><font size=""2""><p align=""center"">" & Desk_No & "</font></td>"
                .Write "        <td width=""50"" height=""1""><font size=""2""><p align=""center"">" & Ext & "</font></td>"
            End With
            
            ' Verify that E-Mail address has been stored in the database, If not then create E-Mail address with employee name
            If Emp_Email = "None" Then
                With Response
                    .Write "        <td width=""250"" height=""1""><font size=""2""><a href=""" & "mailto:" & Last_Name & ", " & First_Name & """>" & Last_Name & ", " & First_Name & "</font></td>"
                End With
            Else
                With Response
                    .Write "        <td width=""250"" height=""1""><font size=""2""><a href=""" & "mailto:" & Emp_Email & """>" & Emp_Email & "</font></td>"
                End With
            End If
            
            With Response
                .Write "        <td width=""200"" height=""1""><font size=""2"">" & Emp_Dept & "</font></td>"
                .Write "    </tr>"
            End With
            adoPrimaryRS.MoveNext
        Loop
    End If

    With Response
        .Write "    </table>"
        .Write "</div>"
    End With
    
    EndHTML
    
    db.Close
    
End Sub

Function BuildTableHeadings()

    ' Build table headings
    With Response
        .Write "<p>"
        .Write "<div align=""left"">"
        .Write "    <table border=""1"" width=""800"" cellpadding=""3"" cellspacing=""0"">"
        .Write "        <tr>"
        .Write "            <td width=""200"" height=""1"" bgcolor=""#6699CC"" ><font color=""#FFFFFF""><b>Employee</b></font></th>"
        .Write "            <td width=""100"" height=""1"" bgcolor=""#6699CC""><font color=""#FFFFFF""><b>Desk No.</b></font></th>"
        .Write "            <td width=""50"" height=""1"" bgcolor=""#6699CC""><font color=""#FFFFFF""><b>Extension</b></font></th>"
        .Write "            <td width=""250"" height=""1"" bgcolor=""#6699CC""><font color=""#FFFFFF""><b>E-Mail</b></font></th>"
        .Write "            <td width=""200"" height=""1"" bgcolor=""#6699CC""><font color=""#FFFFFF""><b>Department</b></font></th>"
        .Write "        </tr>"
    End With
    
End Function

Function ConnectToDB()

    Set db = New Connection
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=MSDASQL;dsn=DeskRef;uid=Sa;pwd=;database=DeskRef;"
    Set adoPrimaryRS = New Recordset
    
End Function

Function StartHTML()

    With Response
        .Write "<html>"
        .Write "<head>"
        .Write "<title>Spotlight Header</title>"
        .Write "</head>"
        .Write "<BODY BGCOLOR=""#FFFFFF"" LINK=""Blue"" VLINK=""Blue"" ALINK=""#BEF46C"" TEXT=""#000000"">"
    End With

End Function

Function EndHTML()

    With Response
        .Write "</body>"
        .Write "</html>"
    End With

End Function

Function AddSearchForm()

    With Response
        .Write "<form method=""GET"" action=""Directory.ASP"">"
        .Write "    <input type=""hidden"" name=""wci"" value=""Search""><div align=""left""><table border=""1"" width=""800"" height=""94"" cellspacing=""0"">"
        .Write "        <tr>"
        .Write "            <td width=""800"" height=""29"" bgcolor=""#6699CC""><div align=""center""><center><h2><em><u><big><font color=""#FFFFFF"">Employee Search</font></big></u></em></h2>"
        .Write "                </center></div></td>"
        .Write "        </tr>"
        .Write "        <tr>"
        .Write "            <td width=""800"" height=""20""><div align=""center""><center><p><u><strong><big>Search By:</big><br>"
        .Write "                </strong></u><input type=""radio"" value=""FirstName"" checked name=""SearchType"">First Name <input type=""radio"" name=""SearchType"" value=""LastName"">Last Name <input type=""radio"" name=""SearchType"" value=""DeskNumber"">Desk No. <input type=""radio"" name=""SearchType"" value=""Extension"">Extension <input type=""radio"" name=""SearchType"" value=""Email"">E-Mail <input type=""radio"" name=""SearchType"" value=""Department"">Department</td>"
        .Write "        </tr>"
        .Write "        <tr align=""center"">"
        .Write "            <td width=""800"" height=""43""><div align=""center""><center><p><input type=""text"" name=""SearchText"" size=""20""><input type=""submit"" value=""Search"" name=""B1""><input type=""reset"" value=""Reset"" name=""B2""></td>"
        .Write "        </tr>"
        .Write "</table></div>"
        .Write "</form>"
    End With
    
End Function

Private Sub SearchPage_Respond()

    StartHTML
    
    AddSearchForm
    
    EndHTML
    
End Sub

Private Sub WebClass_Start()
    
    'Write a reply to the user
    With Response
        .Write "<html>"
        .Write "<body>"
        .Write "<h1><font face=""Arial"">WebClass1's Starting Page</font></h1>"
        .Write "<p>This response was created in the Start event of WebClass1.</p>"
        .Write "</body>"
        .Write "</html>"
    End With

End Sub
