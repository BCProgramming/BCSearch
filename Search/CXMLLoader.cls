VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXMLLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'XMLLoader:



'requirements: MSXML3, CEventPropagator

'CEventPropagator is used to implement "basic" functionality of the commandbar, toolbars, etc, as provided in the XML.

'For example, it implements the playing of Sounds when clicking menus or opening menus (Before/AfterShowMenu).

'CEventPropagator Calls Back into This class when it detects an event in it's wrapped Control.


'used to load an XML document- that XML document will then remain loaded, and other routines can be used to populate different controls, given the NAME of the
'type of tag required.



'for example, a XML file can contain info on all sorts of stuff, but unlike my module-based solution, this one won't need to "know" about all the stuff at once.



'LoadXML loads the XML document.

'LoadImageList loads the specified Imagelist and returns it.

'searches through all "IMAGELIST" tags within "UICONFIG" finding one with the same name specified. Those <LISTIMAGE> tags are processed, and the resulting Imagelist is returned.








'Features:

'allows the loading of a VBAccelerator CommandBar Via XML.




'This code was SALVAGED from BASeEdit XP. God rest it's soul...
'I shouldn't have too hard of a time re-creating a new version, since the main hurdle
'was learning to use the commandbar controls. Besides, it isn't like I need to re-write EVERYTHING.
'just the code I need to re-write. Actually, now that  I really look at it, it was BASeEditor PRO that took a Crapload of time.
'and let's be honest here, BASeEditor Pro kind of blows.

'Anyway, I decided to toss this into a different module, since I can re-use it pretty easily.




'Usage information:


'the XML is fairly simple- a group of "<MENUSET>" tags, each with a group of "<MENU>" tags, which can, of course, be nested, to create heirarchal menus.



'NEW FEATURES TO ADD:
'(making this more of an all-around XML adapter for heirarchal control data)

'Imagelist loading (IE, "<IMAGELIST> tag) ***PARTIAL SUPPORT DONE***

'special toolbar-type command bar button handling. (or even the ability to load MSCOMCTL toolbars via XML... kind of weird though, since the dropdowns can only have one level)


'ideally, all this could be done in one function all at once- that is, simply different tags being used for different purposes. Problem being that the routine would have a vast
'number of arguments- on the other hand, they WILL be optional arguments...
Private Const SND_ASYNC As Long = &H1
Private Const SND_FILENAME As Long = &H20000

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long




Private Type mMenuData
    MenuOperation As String
End Type


Private Type mButtonData
    ButtonMenuset As String
End Type

Public Event ResolveImageKey(ByVal KeyString As String, ByRef PicIndex As Long)
Public Event GetImageList(ByVal Keyof As String, ByRef retList As cVBALImageList)


Private XMLImageLists As New Collection 'Collection of cvbaLImageList objects- these will be/are cached so future references to
'the same key won't need to be recreated.
'Since this class will generally only be present when commandbar controls are also referenced, the class is probably a good idea to include as well.



Private mHandledObjects() As CEventPropagator 'array of eventpropagators...
Private mHandledCount As Long 'count of EventPropagators...

Private mCurrentList As cVBALImageList
Private flLoaded As Boolean
Private XMLParser As MSXML2.DOMDocument30
Private IMLcache As Collection
Private PopupCache As Collection 'collection of cached popupmenus
Private mDataCache() As mMenuData 'used by xml loading of popup menu.
Private mButtonDataCache() As mButtonData
Private mButtonCacheLen As Long
Private DataCacheLen As Long

'HELPER ROUTINES FOR HANDLEDOBJECTS:
Public Function AddHandledControl(ControlAdd As Object) As Long
    AddHandledControl = mHandledCount
    If IsControlHandled(ControlAdd) Then Exit Function
    mHandledCount = mHandledCount + 1
    AddHandledControl = mHandledCount
    ReDim Preserve mHandledObjects(1 To mHandledCount)
    Set mHandledObjects(mHandledCount) = New CEventPropagator
    With mHandledObjects(mHandledCount)
        .Init Me, ControlAdd
    End With



End Function
Private Function IsControlHandled(ControlTest As Object) As Boolean
    Dim I As Long, castit As CEventPropagator
    For I = 1 To mHandledCount
        Set castit = mHandledObjects(I)
        If Not castit Is Nothing Then
            If castit.ControlHandled Is ControlTest Then
                IsControlHandled = True
                Exit Function
                
            End If
        End If
    
    Next


End Function



'-----------------------------


Public Function GetPopupMenuOperation(ByVal ItemData As Long) As String
    GetPopupMenuOperation = mDataCache(ItemData).MenuOperation
End Function
Public Sub LoadXMLFile(ByVal XMLFile As String)
    Dim XMLStr As String
    XMLStr = LoadFileText(XMLFile)
    LoadXML XMLStr


End Sub
Public Sub LoadXML(ByVal XMLSource As String)
    'must be called before other routines.
    Set XMLParser = New DOMDocument30
    If Not (XMLParser.LoadXML(XMLSource)) Then
        
    
        Debug.Print "failed to Load XML data:" & vbCrLf & vbCrLf & XMLParser.parseError.reason & " on line """ & XMLParser.parseError.Line & """" & vbCrLf & _
        "position """ & XMLParser.parseError.linepos & """."
        cDebug.Post "XML load failure in CXMLLoader::LoadXML"
        flLoaded = False
        Set XMLParser = Nothing
        Exit Sub
        
    Else
        flLoaded = True
        Debug.Print "XML load successful in CXMLLoader::LoadXML"
        
    End If


End Sub

Public Function LoadImageList(ByVal ListName As String) As cVBALImageList
    Dim toplevel As IXMLDOMNode
    Dim LoadList As cVBALImageList
    Dim looplevel As IXMLDOMNode
    
    On Error Resume Next
    Set LoadImageList = IMLcache(ListName)
    If Err = 0 Then Exit Function
    
    
    Set toplevel = GetTopLevel(XMLParser)
    For Each looplevel In toplevel.childNodes
        If looplevel.nodeName = "IMAGELIST" Then
            'new feature- loading of an imagelist.
            
            
            
        
        
            If StrComp(looplevel.Attributes.getNamedItem("NAME").nodeValue, ListName) = 0 Then
            Set LoadList = New cVBALImageList
            'XMLImageLists.Add LoadImageList, TopLevel.Attributes.getNamedItem("NAME").nodeValue
            
            'specify properties of the imagelist here
            'or more precisely, retrieve them via attributes.
            Dim AtWidth As Long, AtHeight As Long, ATbitDepth As Long
            Dim atSRC As String
            'also allow for SRC= for picturestrip.
            With looplevel.Attributes
            On Error Resume Next
            AtWidth = .getNamedItem("WIDTH").nodeValue
            AtHeight = .getNamedItem("HEIGHT").nodeValue
            ATbitDepth = .getNamedItem("DEPTH").nodeValue
            atSRC = .getNamedItem("SRC").nodeValue
            
            
            If AtHeight = 0 Then AtHeight = 32
            If AtWidth = 0 Then AtWidth = 32
            
            
            End With
            With LoadList
            .IconSizeX = AtWidth
            .IconSizeY = AtHeight
            .ColourDepth = IIf(ATbitDepth = 0, 32, ATbitDepth)
            LoadList.Create
            If atSRC <> "" Then
            'vbmagenta, for now.
            .AddFromFile atSRC, IMAGE_BITMAP, , True, vbMagenta
            
            End If
            
            End With
            
            
            
             LoadListImagesFromNode looplevel, LoadList
             Set LoadImageList = LoadList
             IMLcache.Add LoadList, ListName
            Exit Function
        End If
        On Error Resume Next
        
       
        End If
    
    Next
End Function
Private Function GetTopLevel(OfXML As DOMDocument30) As IXMLDOMNode
Dim toplevel As IXMLDOMNode
    For Each toplevel In XMLParser.childNodes
    
        If StrComp(toplevel.nodeName, "UICONFIG", vbTextCompare) = 0 Then
            
            Set GetTopLevel = toplevel
            Exit For
            
        End If
        
    Next




End Function
Public Sub LoadCommandBar(OnCtl As vbalCmdBar6.vbalCommandBar)
    'notice the lack of "keying" on here- this works with a control, and all commandbars are global anyway.
    
    'what should occur:
    'Load XML
    'load Imagelists
    'set commandbar to one of the imagelists
    'load commandbar
    'etc...
    Dim XMLParser As MSXML2.DOMDocument30
    On Error Resume Next
    AddHandledControl OnCtl
    On Error GoTo 0
    Dim toplevel As IXMLDOMNode, NewCommandBar As cCommandBar
    Dim TopLoop As IXMLDOMNode, ImageUse As Object, uiconfig As IXMLDOMNode
    'OK, this one is pretty simple.
    'The other XML generator, GenerateMenuFromXML, adds menus to a single Ccommandbar.
    'this one will create a ccommandbar from each "MENUSET" tag.
    'naturally, a required step in this procedure is parsing the XML.
    
    'ok, ANYWAY...
    'each top-level node should have the tag, "MENUSET"
    
    'hmm, XML can't have more then one top-level node. shitty.
    'So- they will all be inside a "UICONFIG" tag.
    
    Set uiconfig = GetTopLevel(XMLParser)
    If uiconfig Is Nothing Then
        Debug.Print "failed to locate uiconfig key."
    End If
    
    For Each toplevel In uiconfig.childNodes
        If toplevel.nodeName = "MENUSET" Then
        'each MenuSet is another cCommandBar that must be added
        'attributes: Name=key.
            
                
                
                'that is pretty much it.
                On Error Resume Next
                Set NewCommandBar = OnCtl.CommandBars.Add(toplevel.Attributes.getNamedItem("NAME").nodeValue)
                'ok, now the "tricky" part.
                'we want to recursively create the menu.
                If Err <> 0 Then
                    Set NewCommandBar = OnCtl.CommandBars.Item(toplevel.Attributes.getNamedItem("NAME").nodeValue)
                
                End If
                'Call GenerateMenuFromXML(toplevel.childNodes(0).xml, OnCtl, NewCommandBar)
                If Not NewCommandBar Is Nothing Then
                LoadBarFromXML toplevel, OnCtl, NewCommandBar
                'there we go.
                End If
           
    
    
            Else
    Debug.Print "unrecognized NodeName for commandbar load:" & toplevel.nodeName
    
            End If
    Next toplevel
    
    


End Sub

Private Sub LoadListImagesFromNode(ListNode As IXMLDOMNode, ToList As cVBALImageList)
'loads a SINGLE imagelist from the given node- that is, a "<IMAGELIST>" node.

'yes, the name is confusing.


Dim Ilkey As String, useList As cVBALImageList
Dim loopnode As IXMLDOMNode
Dim AttrWidth As Long, AttrHeight As Long, attrbits As Long
'assumes node is <IMAGELIST> node.
'alright, now this is progress.
'structure:
'<IMAGELIST NAME="<KEY>">
'<LISTIMAGE NAME="KEY" SRC="filename"></IMAGE>
'or
'<LISTIMAGE NAME="KEY" SRC="RES:(TYPE,ID)"></IMAGE>


   
    For Each loopnode In ListNode.childNodes
        loadListImageXML loopnode, ToList


    Next loopnode
   
   
End Sub



Private Sub loadListImageXML(XMLnode As IXMLDOMNode, ToList As cVBALImageList)
'loads the information for a single "LISTIMAGE" tag.

    '<LISTIMAGE> tags.
    '<LISTIMAGE NAME="KEY" SRC="filename"></IMAGE>
    'or
    '<LISTIMAGE NAME="KEY" SRC="RES:(TYPE,ID)"></IMAGE>

    Dim ImageSRC As String, ImageName As String
    Dim resType As String, ResID As String
    On Error Resume Next
    ImageName = XMLnode.Attributes.getNamedItem("NAME").nodeTypedValue
    ImageSRC = XMLnode.Attributes.getNamedItem("SRC").nodeTypedValue
    If ImageName = "" Or ImageSRC = "" Then
        Debug.Print "loadListImageXML: Required Attributes NAME and SRC not specified."
        Exit Sub

    End If
    If StrComp(Left$(ImageSRC, 4), "RES:", vbTextCompare) = 0 Then
        'resource.
        resType = Mid$(ImageSRC, 5, InStr(ImageSRC, ",") - 5)
        ResID = Mid$(ImageSRC, InStr(ImageSRC, ",") + 1)
        ResID = Replace$(ResID, ")", "")
        Call ToList.AddFromResourceID(0, App.hInstance, IMAGE_ICON, ImageName, False)
    Else
        'assume file reference
        ImageSRC = App.Path & "\" & ImageSRC
        Call ToList.AddFromFile(ImageSRC, IMAGE_ICON, ImageName, False)
    End If


End Sub


Public Sub LoadBarFromXML(ByVal XMLnode As IXMLDOMNode, inctl As vbalCmdBar6.vbalCommandBar, OnBar As vbalCmdBar6.cCommandBar)

    Dim XMLParser As MSXML2.DOMDocument30
    Set XMLParser = New MSXML2.DOMDocument30
    Dim Typesof As MSXML2.IXMLDOMNode
    Dim toplevel As MSXML2.IXMLDOMNode
    Dim newbar As cCommandBar
    Dim newbutton As cButton
    
'    If Not (XMLParser.LoadXML(XMLStr)) Then
'        Debug.Print "XML load failure"
'    End If
    
    '<MENU STYLE="style">
    '    <MENU property=value...>
        '(any sub-menus will be inside the <MENU> tag.
        '</MENU>
    '</MENU>
    'first, we need to clear out existing items.
    OnBar.Buttons.Clear
    'OK, If I know what I'm doing, we create a menu quite simply:
    'each top-level item needs to be added as a button.
    'For Each Typesof In XMLnode.childNodes
    
    
 '   If StrComp(Typesof.baseName, "MENU", vbTextCompare) = 0 Then
        'if is is a menu, we need to create one button on the command bar for
        'each top-level menu item.
'        Dim menuname As String
'        On Error Resume Next
'        menuname = Typesof.Attributes.getNamedItem("NAME").nodeTypedValue 'name of the menu.
'
'        On Error GoTo 0
    Set newbar = OnBar
    'loop through children, if existent.
    
    For Each toplevel In XMLnode.childNodes
        If StrComp(toplevel.baseName, "MENU", vbTextCompare) = 0 Then
            Set newbutton = MakeButton(inctl, inctl.Buttons, toplevel)
            newbar.Buttons.Add newbutton
            'if this node has children, add a submenu bar. otherwise, don't.
            If toplevel.childNodes.Length > 0 Then
                Set OnBar = inctl.CommandBars.Add(toplevel.Attributes.getNamedItem("NAME").nodeTypedValue & "::SUBMENU")
                
                Call CreateSubMenu(inctl, OnBar, toplevel)
                
                 Set newbutton.Bar = OnBar
                 
                 
             newbutton.ShowDropDownInToolbar = False
            Else
                'Stop
            End If
        End If
   Next
   'Next
'    For Each toplevel In XMLparser.childNodes
'        'add a new button, and add it to the commandbar...
'        Set newbutton = inctl.Buttons.Add("MENU::" & toplevel.baseName, , _
'            toplevel.Attributes("Caption").baseName)
'        newbutton.ShowCaptionInToolbar = True
'        OnBar.Buttons.Add newbutton
'        'create a new commandbar, populate it with this menus XML, and
'        'set it as that items bar.
'        Set newbar = inctl.CommandBars.Add(newbutton.Key & "::SUBMENU", newbutton.Caption & " Sub menu.")
'        Call CreateSubMenu(inctl, newbar, toplevel)
'        Set newbutton.Bar = newbar
'        newbutton.Enabled = True
'
'
'
'    Next toplevel
    'Else
    '    Debug.Print "huh. what the hell is a " & Typesof.baseName & "?"
    
    'End If
    
    inctl.MainMenu = True
    Set inctl.Toolbar = newbar
    
    'now it makes sense. Every commandbar is global, but each separate
    'commandbar control can be told to display a different commandbar in it's client area.
    'I get it!.
    
End Sub
Private Sub CreateSubMenu(inctl As vbalCmdBar6.vbalCommandBar, OnBar As cCommandBar, XMLnode As IXMLDOMNode)
Dim toplevel As MSXML2.IXMLDOMNode

Dim newbutton As cButton, newbar As cCommandBar
    For Each toplevel In XMLnode.childNodes
        'add a new button, and add it to the commandbar...
        
       ' newbutton.ShowCaptionInToolbar = True
        'OnBar.Buttons.Add newbutton
        Set newbutton = MakeButton(inctl, inctl.Buttons, toplevel)
        OnBar.Buttons.Add newbutton
        'create a new commandbar, populate it with this menus XML, and
        'set it as that items bar.
        'but only if it HAS items, otherwise, we'll
        'keep popping up empty menus.
        If toplevel.childNodes.Length > 0 Then
'            Set newbar = inctl.CommandBars.Add(newbutton.Key & "::SUBMENU", newbutton.Caption & " Sub menu.")
Set newbar = inctl.CommandBars.Add(newbutton.key & "::SUBMENU")
            
            Call CreateSubMenu(inctl, newbar, toplevel)
            Set newbutton.Bar = newbar
        End If
     
        
        
    
    Next toplevel






End Sub
Private Function MakeButton(InControl As vbalCommandBar, FromCollection As cButtons, nodeRep As IXMLDOMNode) As cButton
    Dim captionuse As String, KeyUse As String
    Dim styleuse As String, PicIndex As Variant
    Dim CheckedUse As String, ColorBoxUse As String, panelWidthUse As String
    Dim styleadd As vbalCmdBar6.EButtonStyle
    Dim childloop As IXMLDOMNode
    Dim tooltipuse As String, taguse As String
    Dim enab As Boolean
    Dim created As cButton, acceluse As String
    Dim ControlName As String
    Dim ShowCaptionInTbar As Boolean
    Dim ShowDropDownInTbar As Boolean
    PicIndex = -1
    ColorBoxUse = -1
    
    'name is the key. but the actual operation performed will be
    ' the tag.
    
    KeyUse = nodeRep.Attributes.getNamedItem("NAME").nodeTypedValue
    On Error Resume Next
    captionuse = nodeRep.Attributes.getNamedItem("CAPTION").nodeTypedValue
    'replace "{}" with an ampersand. XML F's up for some reason with the ampersand
    'in the XML.
    captionuse = Replace$(captionuse, "{}", "&")
    On Error Resume Next
    styleuse = nodeRep.Attributes.getNamedItem("STYLE").nodeTypedValue
    tooltipuse = nodeRep.Attributes.getNamedItem("TOOLTIP").nodeTypedValue
    acceluse = nodeRep.Attributes.getNamedItem("ACCEL").nodeTypedValue
    taguse = nodeRep.Attributes.getNamedItem("OPERATION").nodeTypedValue
    CheckedUse = nodeRep.Attributes.getNamedItem("CHECKED").nodeTypedValue
    ColorBoxUse = nodeRep.Attributes.getNamedItem("COLORBOX").nodeTypedValue
    panelWidthUse = nodeRep.Attributes.getNamedItem("PANELWIDTH").nodeTypedValue
    '"PanelControl"
    ControlName = nodeRep.Attributes.getNamedItem("PANELCONTROL").nodeTypedValue
    ShowCaptionInTbar = True
    ShowCaptionInTbar = nodeRep.Attributes.getNamedItem("SHOWCAPTIONINTB").nodeTypedValue
    ShowDropDownInTbar = nodeRep.Attributes.getNamedItem("SHOWDROPDOWNINTB").nodeTypedValue
    If ColorBoxUse = "0" Then ColorBoxUse = "-1"
    taguse = Trim$(taguse)
    If taguse = "" Then
        'if no OPERATION, set to name.
        taguse = KeyUse
    
    End If
    On Error Resume Next
    Dim Temp As String
    Temp = nodeRep.Attributes.getNamedItem("ENABLED").nodeTypedValue
    If Temp <> "" Then
        enab = CBool(Temp)
    Else
        enab = True
    End If
    
    'enab = IIf(nodeRep.Attributes.getNamedItem("ENABLED").nodeTypedValue <> "", CBool(nodeRep.Attributes.getNamedItem("ENABLED").nodeTypedValue), True)
    
    'If Err Then enab = True
    Err.Clear
    PicIndex = nodeRep.Attributes.getNamedItem("PIC").nodeTypedValue
    If Err Then PicIndex = -1
    On Error GoTo 0
    If styleuse <> "" Then
        Select Case StrConv(styleuse, vbProperCase)
            Case "Separator"
                styleadd = eSeparator
            Case "Check"
                styleadd = eCheck
            Case "Panel"
                styleadd = ePanel
            Case "Radio"
                styleadd = eRadio
            Case "Option", "Radionullable"
                styleadd = eRadioNullable
            Case "Split"
                styleadd = eSplit
            Case Else
                styleadd = eNormal
        End Select
                
    
    End If
    If Not IsNumeric(PicIndex) Then
    'add code to get a picture index from the string.
        'picindex = MDIMain.Get16x16Index(CStr(picindex))
        On Error Resume Next
        'picindex = AppObject.Get16x16Index(CStr(picindex))
        
        'we need to convert the given string to an actual index for the commandbar.
        'picindex = mCurrentList.ItemIndex(CStr(picindex))
        Dim getindex As Long
        RaiseEvent ResolveImageKey(PicIndex, getindex)
        PicIndex = getindex
        If Not IsNumeric(PicIndex) Then PicIndex = -1
        
        On Error GoTo 0
    End If
        Set created = FromCollection.Add(KeyUse, PicIndex, captionuse, styleadd, tooltipuse)
        created.ShowCaptionInToolbar = ShowCaptionInTbar
        created.ShowDropDownInToolbar = ShowDropDownInTbar
        created.enabled = enab
        created.Tag = taguse
        created.checked = CBool(Val(CheckedUse))
        created.ColorBox = CLng(Val(ColorBoxUse))
        created.PanelWidth = CLng(Val(panelWidthUse))
        
          If ControlName <> "" Then
                'set the control...
        
                Set created.PanelControl = CallByName(InControl.Parent, ControlName, VbGet)
                
    
        End If
        
        
        
       ' If acceluse <> "" Then Stop
        AccelStringToButton acceluse, created
        Set MakeButton = created
End Function
Private Sub AccelStringToButton(ByVal AccelString, ApplyTo As cButton)
    
    
        If InStr(AccelString, "^") <> 0 Then
            'ctrl
            ApplyTo.ShortcutModifiers = vbCtrlMask
        End If
        If InStr(AccelString, "+") <> 0 Then
            ApplyTo.ShortcutModifiers = ApplyTo.ShortcutModifiers Xor vbShiftMask
        
            
            'shift
        End If
        If InStr(AccelString, "%") <> 0 Then
            'alt
            ApplyTo.ShortcutModifiers = ApplyTo.ShortcutModifiers Xor vbAltMask
        End If
        AccelString = Replace$(AccelString, "^", "")
        AccelString = Replace$(AccelString, "+", "")
        AccelString = Replace$(AccelString, "%", "")
        With ApplyTo
        Select Case UCase$(AccelString)
            Case "F1"
                .ShortcutKey = vbKeyF1
            Case "F2"
            .ShortcutKey = vbKeyF2
            Case "F3"
            .ShortcutKey = vbKeyF3
            Case "F4"
            .ShortcutKey = vbKeyF4
            Case "F5"
            .ShortcutKey = vbKeyF5
            Case "F6"
            .ShortcutKey = vbKeyF6
            Case "F7"
            .ShortcutKey = vbKeyF7
            Case "F8"
            .ShortcutKey = vbKeyF8
            Case "F9"
            .ShortcutKey = vbKeyF9
            Case "F10"
            .ShortcutKey = vbKeyF10
            Case "F11"
            .ShortcutKey = vbKeyF11
            Case "F12"
            .ShortcutKey = vbKeyF12
            Case ""
            Case Else
            .ShortcutKey = Asc(UCase$(AccelString))
           
            End Select
        
        End With
        
        

End Sub
Public Function LoadFileText(ByVal FFile As String) As String
    Dim strLoad As String
    Dim fNum As Long
    fNum = FreeFile
    Open FFile For Input As fNum
        strLoad = Input$(LOF(fNum), fNum)
    Close #fNum
    LoadFileText = strLoad


End Function

Public Function GetXMLStr(VarFrom As Variant) As String
    Dim strCreate As String, I As Long
    'the start tag will have been made before.
    
    
    If Not IsArray(VarFrom) Then
        'if it isn't an array, create the item.
        strCreate = "<MENU NAME=""" & VarFrom & Int(Rnd * 1000) & """ CAPTION=""" & VarFrom & """></MENU>" & vbCrLf
        'Debug.Print strCreate
    Else
        'it is an array- make a new menu with that caption.
        strCreate = "<MENU NAME=""" & VarFrom(0) & Int(Rnd * 1000) & """ CAPTION=""" & VarFrom(0) & """>"
        'now, we recursively add the items for our subitems.
        'Debug.Print strCreate
        For I = 1 To UBound(VarFrom)
            strCreate = strCreate & GetXMLStr(VarFrom(I))
            'Debug.Print strCreate
        Next I
        'now close off the <MENU> tag.
        strCreate = strCreate & "</MENU>" & vbCrLf
    End If
    'return the built string.
    GetXMLStr = strCreate



End Function






'
'Public Sub PopulateColourMenu(BarPop As cCommandBar, OnCtl As vbalCommandBar)
''populates the bar.
'Dim CColour As ClsColour
'Dim ColourNames(), newbutton As cButton
'Dim I As Long, Formattedname As String
'Debug.Print "populated"
'Set CColour = New ClsColour
'ColourNames = Array("Black", "Navy Blue", "Green", "Teal", "Maroon", "Purple", "Olive", "Silver", "Gray", "Blue", "Lime", "Cyan", "Red", "Magenta", "Yellow", "White", "AliceBlue", "AntiqueWhite", "Aquamarine", "Azure", "Biege", "Bisque", "BlanchedAlmond", "BlueViolet", "Brown", "BurlyWood", "CadetBlue", "Chatreuse", "Chocolate", "Coral", "CornflowerBlue", "Cordsilk", "Crimson", "Lime", "Magenta", "Pink", "Salmon", "Tan")
''whew.
''Set newbutton = OnCtl.Buttons.Add(BarPop.Key & "::COLOUR::AUTOMATIC", , "Automatic", eNormal, "Automatic")
''newbutton.ColorBox = vbBlack
''newbutton.Caption = "Automatic"
''BarPop.Buttons.Add newbutton
'For I = 0 To UBound(ColourNames)
'    On Error Resume Next
'    Formattedname = ColourNames(I) & "(" & CColour.CreateFromName(ColourNames(I)).WebFormat & ")"
'    Set newbutton = OnCtl.Buttons.Add(BarPop.Key & "::COLOUR::" & ColourNames(I), , , eNormal, Formattedname)
'    If Err <> 0 Then
'        Err.Clear
'        Exit For
'    End If
'    newbutton.ColorBox = CColour.CreateFromName(ColourNames(I)).LongVal
'
'
'    newbutton.Visible = True
'    newbutton.Enabled = True
'    newbutton.Tag = "COLOUR::" & ColourNames(I)
'    BarPop.Buttons.Add newbutton
'    If (I + 1) Mod CSettings.ColoursPerRowInPopup = 0 And I > 0 Then
'       'create a new row.
'        Debug.Print "hi there, adding split." & I & " Mod 8 = " & I Mod 8
'        Set newbutton = OnCtl.Buttons.Add(BarPop.Key & "SEP#" & I, , , eSeparator)
'        newbutton.Visible = True
'        BarPop.Buttons.Add newbutton
'
'    End If
'
'   ' Set newbutton = OnCtl.Buttons.Add("SEPCUSTOM2", , , eSeparator)
'   ' BarPop.Buttons.Add newbutton
'    'The custom colors. These are stored in the registry.
'
'
'Next I
'
'
'
'
'
'End Sub
'

Public Function GetPopupMenu(ByVal PopupKey As String, formowner As Long, Optional usehiml As Long = 0, Optional ByRef CacheIndex As Long) As PopupMenu6.cPopupMenu
Static popup As cPopupMenu
Dim toplevel As IXMLDOMNode, looper As IXMLDOMNode
'how it appears to work.
'everything revolves around additem. additem is very important.

'The XML for MENUSET NAME="popupkey" will be loaded.


'I find it a bit strange, but apparently all the menus are loaded serially...
On Error Resume Next
'check for cached popup object
Set GetPopupMenu = PopupCache(PopupKey)
If Err = 0 Then
    'return index, too.
    
End If

'popup.Restore PopupKey
'If Err <> 0 Then
    Set popup = New cPopupMenu
    'find the key... in our xmldoc
    popup.ImageList = usehiml
    popup.hwndOwner = formowner
   ' popup.CreateSubClass formowner
    Set toplevel = GetTopLevel(XMLParser)
    For Each looper In toplevel.childNodes
        If StrComp(looper.nodeName, "MENUSET") = 0 Then
            If StrComp(looper.Attributes.getNamedItem("NAME").nodeTypedValue, PopupKey) = 0 Then
                LoadPopupXML looper, popup
                'popup.Store PopupKey
                Set GetPopupMenu = popup
                Exit Function
            End If
        
        
        End If
    
    Next looper


'End If

Set GetPopupMenu = popup


End Function
Private Sub AddPopupMenuItem(ToMenu As cPopupMenu, ItemNode As IXMLDOMNode, ParentIndex As Long)
    'adds the menu item described in the Node ItemNode, also recurses on each
    'child Node and calls LoadPopupXML for each one, to add them as children of this menu.
    Dim caption As String, Helptext As String, ItemData As Long, IconIndex As Long, checked As Boolean, enabled As Boolean, key As String
    Dim operation As String
    'itemdata will need to be used to somehow encode the OPERATION string of the MENU tag...
    'tomenu.AddItem scaption,shelptext,litemdata,lparentindex,liconindex,bchecked,benabled,skey
    Dim pickey As Variant, addedindex As Long
    Dim loopchild As IXMLDOMNode, styleuse As Variant, tooltipuse As String, acceluse As String
    With ItemNode.Attributes
        On Error Resume Next
        caption = .getNamedItem("CAPTION").nodeTypedValue
        operation = .getNamedItem("OPERATION").nodeTypedValue
        caption = Replace$(caption, "{}", "&")
        Helptext = .getNamedItem("HELPTEXT").nodeTypedValue
        Err.Clear
        pickey = .getNamedItem("PIC").nodeTypedValue

        If Err <> 0 Then pickey = -1
        
        
        checked = .getNamedItem("CHECKED").nodeTypedValue
        On Error Resume Next
        enabled = .getNamedItem("ENABLED").nodeTypedValue
        If Err <> 0 Then
            enabled = True
        End If
        key = .getNamedItem("NAME").nodeTypedValue
        
         styleuse = .getNamedItem("STYLE").nodeTypedValue
        tooltipuse = .getNamedItem("TOOLTIP").nodeTypedValue
        acceluse = .getNamedItem("ACCEL").nodeTypedValue
        'taguse = nodeRep.Attributes.getNamedItem("OPERATION").nodeTypedValue
        
        Select Case styleuse
            Case "SEPARATOR"
                caption = "-" & caption
        End Select
    End With
    If Not IsNumeric(pickey) Then
        RaiseEvent ResolveImageKey(pickey, IconIndex)
    End If
    
    
  
    'alrighty then, we have an "addedindex"
    'what do we do now? Iterate through sub-items, call addpopupmenuitem for each to add to the new menu item.
    
    'here's the catch- we need to store the "operation" value of each menu item, so that the event knows what to do.
    'so the question is, "HOW?"
    
    
    'well, after a little thought, I have decided on the solution; a string array, stored
    'at module level, that will store the operation string (and other data too) as an item in the array, the index of this item will be stored
    'in the itemdata prop for that menuitem.
    DataCacheLen = DataCacheLen + 1
    ReDim Preserve mDataCache(1 To DataCacheLen)
    mDataCache(DataCacheLen).MenuOperation = operation
    
      If pickey = -1 Then
        addedindex = ToMenu.AddItem(caption, Helptext, DataCacheLen, ParentIndex, , checked, enabled)
    Else
        addedindex = ToMenu.AddItem(caption, Helptext, DataCacheLen, ParentIndex, IconIndex, checked, enabled)
        
    End If
    
    
    For Each loopchild In ItemNode.childNodes
        AddPopupMenuItem ToMenu, loopchild, addedindex
    
    
    Next



End Sub
Private Sub LoadPopupXML(XMLnode As IXMLDOMNode, ToMenu As cPopupMenu)
'XMLnode is the MENUSET node.


Dim loopchild As IXMLDOMNode

For Each loopchild In XMLnode.childNodes
    AddPopupMenuItem ToMenu, loopchild, 0


Next loopchild








End Sub
Private Sub Class_Initialize()
    Set IMLcache = New Collection
    Set PopupCache = New Collection
End Sub

'VBA toolbar loader
''routine requires bulletproofing...
'Public Sub LoadVBAToolbar(VBAbar As vbalTBar6.cToolbar, ByVal BarName As String)
''XML:
'Dim ToolbarNode As IXMLDOMNode, found As Boolean
'Dim buttonnode As IXMLDOMNode
''<TOOLBAR NAME=<barname> IMAGELIST=<imagelistname>
'Dim tbarName As String, tbarList As String, listObj As cVBALImageList
'Dim ListHot As cVBALImageList, ListDisabled As cVBALImageList
'Dim strHot As String, StrDisabled As String
'Dim size As Long, wrappable As Boolean, picsize As Long, stylelist As Boolean
'Dim buttonmenu As String, newind As Long
'Dim toplevel As IXMLDOMNode
'If XMLParser Is Nothing Then
'    Err.Raise 9, "XCMLLoader::LoadVBAToolbar", "Cannot parse XML without XML file loaded."
'End If
'
'
'
'
'VBAbar.DestroyToolBar
'
'
'
'
'
'
'
''step One: retrieve the top level (UICONFIG) key.
'Set toplevel = GetTopLevel(XMLParser)
'
'
'
''loop through each node beneath UICONFIG...
'For Each ToolbarNode In toplevel.childNodes
'    On Error Resume Next
'    If StrComp(ToolbarNode.nodeName, "TOOLBAR", vbTextCompare) = 0 Then
'        'if we find a TOOLBAR tag...
'        If StrComp(ToolbarNode.Attributes.getNamedItem("NAME").nodeTypedValue, BarName, vbTextCompare) = 0 Then
'            'check that it has the name we're looking for.
'            If Err = 0 Then
'                found = True
'                'and if so, set the flag, and break out of the loop.
'                Exit For
'
'
'
'            End If
'        End If
'
'    End If
'    Err.Clear
'
'Next ToolbarNode
'
' 'ToolbarNode now points to proper toolbar- unless it's nothing.
''if it's nothing, the toolbar tag with the given name wasn't found. oh well.
'If ToolbarNode Is Nothing Then Exit Sub
'Dim showtext
'On Error Resume Next
'
''retrieve the attributes of the toolbar. we use LoadImageList for imagelist names, since that routine caches
''Imagelists, and can be used in a jiffy to retrieve the correct one.
'
'size = ToolbarNode.Attributes.getNamedItem("SIZE").nodeTypedValue
'tbarList = ToolbarNode.Attributes.getNamedItem("IMAGELIST").nodeTypedValue
'
'strHot = ToolbarNode.Attributes.getNamedItem("HOTIML").nodeTypedValue
'StrDisabled = ToolbarNode.Attributes.getNamedItem("DISABLEDIML").nodeTypedValue
'
'
''retrieve that imagelist object.
''showtext = ToolbarNode.Attributes.getNamedItem("SHOWTEXT").nodeTypedValue
'showtext = 1
'wrappable = ToolbarNode.Attributes.getNamedItem("WRAPPABLE").nodeTypedValue
'picsize = ToolbarNode.Attributes.getNamedItem("PICSIZE").nodeTypedValue
'
'
'
'Set listObj = LoadImageList(tbarList)
'Set ListHot = LoadImageList(strHot)
'Set ListDisabled = LoadImageList(StrDisabled)
'
'If size = 0 Then size = 32
'With VBAbar
'    '.CreateToolbar size, stylelist, showtext, wrappable, picsize
'     .CreateToolbar
'
'        'assign the imagelists, if necessary.
'    If Not listObj Is Nothing Then
'        .SetImageList listObj, CTBImageListNormal
'    End If
'    If Not ListHot Is Nothing Then
'        .SetImageList ListHot, CTBImageListHot
'    End If
'    If Not ListDisabled Is Nothing Then
'        .SetImageList ListDisabled, CTBImageListDisabled
'    End If
'
'
'End With
'
''ok, now we iterate on child nodes, <BUTTON> nodes, to be exact.
'
'For Each buttonnode In ToolbarNode.childNodes
'    If StrComp(buttonnode.nodeName, "BUTTON") = 0 Then
'        'hey- a button. woopee.
'        'add this button
'        'attributes:
'        'CAPTION
'        'PIC
'        'MENU
'        Dim caption As String, tooltip As String, buttonname As String
'        Dim PIC As String, pindex As Long, style As vbalTBar6.ECTBToolButtonSyle
'        'Menu is a MENUSET name.
'        mButtonCacheLen = mButtonCacheLen + 1
'        ReDim Preserve mButtonDataCache(1 To mButtonCacheLen)
'        mButtonDataCache(mButtonCacheLen).ButtonMenuset = buttonmenu
'
'        buttonname = buttonnode.Attributes.getNamedItem("NAME").nodeTypedValue
'
'        caption = buttonnode.Attributes.getNamedItem("CAPTION").nodeTypedValue
'
'
'        PIC = buttonnode.Attributes.getNamedItem("PIC").nodeTypedValue
'        If PIC <> "" And Not listObj Is Nothing Then
'            RaiseEvent ResolveImageKey(PIC, pindex)
'
'        End If
'
'        buttonmenu = buttonnode.Attributes.getNamedItem("MENU").nodeTypedValue
'
'        If buttonmenu <> "" Then
'            style = CTBDropDownArrow
'        Else
'            style = CTBAutoSize
'        End If
'        If pindex = 0 Then pindex = -1
'        Call VBAbar.AddButton(tooltip, pindex, , , caption, style, buttonname)
'        VBAbar.ButtonVisible(buttonname) = True
'        VBAbar.ButtonTextVisible(buttonname) = True
'    End If
'
'
'
'
'Next
'
'
'
'
'
'
'End Sub
'
'
'
'Public Sub LoadRebar(ToRebar As vbalTBar6.cReBar, ByVal RebarName As String)
'
'
''Rebar Loading abilities:
''the <REBAR> tag denotes the properties of a rebar- these properties are loaded in the LoadRebarXML() routine.
''beneath the tag, we find <BAND> tags. band tags can have the following attributes.
''NAME
'
''duh, the key of the band. moving on.
'
''Caption- caption text. also quite self explanatory.
'
''GRIPPER=True/false
'
''is there or is there not a gripper bar for this band.
'
''CHILD=
'
''now THIS is where things get messy.
'
''for now, it will simply set the tag of that band.
'Dim toplevel As IXMLDOMNode
'Dim looprebar As IXMLDOMNode, loopband As IXMLDOMNode
'Set toplevel = GetTopLevel(XMLParser)
'
'For Each looprebar In toplevel.childNodes
'    If StrComp(looprebar.nodeName, "REBAR") = 0 Then
'        'found a rebar tag.
'        If StrComp(looprebar.Attributes.getNamedItem("NAME").nodeTypedValue, RebarName, vbTextCompare) = 0 Then
'            'right name, too!
'            For Each loopband In looprebar.childNodes
'                If StrComp(loopband.nodeName, "BAND") = 0 Then
'                    'good, good. good.
'
'
'
'                End If
'
'
'            Next
'
'
'
'
'        End If
'
'
'
'    End If
'
'
'Next
'
'
'
'
'
'
'
'
'
'End Sub
'
Public Sub HandleControlEvent(OnCtl As Object, FromObject As CEventPropagator, ByVal EventName As String, Params() As Variant)
'
Dim extendit As VBControlExtender
'Set extendit = OnCtl
Debug.Print "HandleControlEvent, " & EventName


'step one: parse the name to understand the parameters...
If StrComp(EventName, "BEFORESHOWMENU", vbTextCompare) = 0 Then
'TODO// change to property.
'    PlaySound "C:\Windows\Media\start.wav", App.hInstance, SND_ASYNC + SND_FILENAME


ElseIf StrComp(EventName, "BUTTONCLICK", vbTextCompare) = 0 Then
    'only param is a cbutton:
    'TODO:// change to property of this class.
'PlaySound "C:\Windows\Media\start.wav", App.hInstance, SND_ASYNC + SND_FILENAME
   ' Debug.Print "Button tag=" & cbuttonget.Tag




End If




End Sub
