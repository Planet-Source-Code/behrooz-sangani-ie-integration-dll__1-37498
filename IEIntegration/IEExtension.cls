VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'=========================================================================================
'  IEExtension Class
'  Demo Usage of the Integration DLL
'=========================================================================================
'  Created By: Behrooz Sangani <bs20014@yahoo.com>
'  Published Date: 01/08/2002
'  WebSite: http://www.geocities.com/bs20014
'  Legal Copyright: Behrooz Sangani © 01/08/2002
'=========================================================================================
'This class is just a demo that collects the information the _
 script provides and shows them.
'You can use this method to integrate your application with _
 Internet Explorer to get the right clicked URL ...
'local variable(s) to hold property value(s)
Private mvarURL As Variant 'local copy
Private mvarPageURL As Variant 'local copy
Private mvarSelText As Variant 'local copy
Private mvarLinkType As Variant 'local copy

Public Property Let LinkType(ByVal vData As Variant)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.LinkType = 5
    mvarLinkType = vData
End Property

Public Property Set LinkType(ByVal vData As Variant)
    'used when assigning an Object to the property, on the left side of a Set statement.
    'Syntax: Set x.LinkType = Form1
    Set mvarLinkType = vData
End Property

Public Property Get LinkType() As Variant
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.LinkType
    If IsObject(mvarLinkType) Then
        Set LinkType = mvarLinkType
    Else
        LinkType = mvarLinkType
    End If
End Property

Public Property Let SelText(ByVal vData As Variant)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.SelText = 5
    mvarSelText = vData
End Property

Public Property Set SelText(ByVal vData As Variant)
    'used when assigning an Object to the property, on the left side of a Set statement.
    'Syntax: Set x.SelText = Form1
    Set mvarSelText = vData
End Property

Public Property Get SelText() As Variant
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.SelText
    If IsObject(mvarSelText) Then
        Set SelText = mvarSelText
    Else
        SelText = mvarSelText
    End If
End Property

Public Property Let PageURL(ByVal vData As Variant)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.PageURL = 5
    mvarPageURL = vData
End Property

Public Property Set PageURL(ByVal vData As Variant)
    'used when assigning an Object to the property, on the left side of a Set statement.
    'Syntax: Set x.PageURL = Form1
    Set mvarPageURL = vData
End Property

Public Property Get PageURL() As Variant
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.PageURL
    If IsObject(mvarPageURL) Then
        Set PageURL = mvarPageURL
    Else
        PageURL = mvarPageURL
    End If
End Property

Public Property Let URL(ByVal vData As Variant)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.URL = 5
    mvarURL = vData
End Property

Public Property Set URL(ByVal vData As Variant)
    'used when assigning an Object to the property, on the left side of a Set statement.
    'Syntax: Set x.URL = Form1
    Set mvarURL = vData
End Property

Public Property Get URL() As Variant
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.URL
    If IsObject(mvarURL) Then
        Set URL = mvarURL
    Else
        URL = mvarURL
    End If
End Property

Public Function ShowInfo()

    frmMain.txtPageURL = mvarPageURL
    frmMain.txtURL = mvarURL
    frmMain.txtSel = mvarSelText
    If mvarLinkType = "Link" Then
        frmMain.Opt(0).Value = True
    ElseIf mvarLinkType = "IMG" Then
        frmMain.Opt(1).Value = True
    End If
    'Remember that you can not show this form as non-modal
    'Remove the vbModal and see the error message you get
    frmMain.Show vbModal

End Function
