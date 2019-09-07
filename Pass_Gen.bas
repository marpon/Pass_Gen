'compile with "c:\Freebasic_Path\fbc.exe" "c:\Pass_Gen_path\Pass_Gen.bas" -x "Pass_Gen.exe" -s gui -v "Pass_Gen.rc" > "Pass_Gen.log" 2>&1
'
'   change Freebasic_Path\fbc.exe      to your own fbc path
'   change Pass_Gen_path\Pass_Gen.bas  to your own source path



#Include Once "windows.bi"
#Include Once "win/commctrl.bi"                  ' needed for WinXP Theme support

' ========================================================================================

#if __FB_VERSION__ < "1.02"
    #define GCLP_HBRBACKGROUND GCL_HBRBACKGROUND
#endif

Type FLY_RESCALEDATA
        FormInit                   As Long       ' Form initial width or height
        FormCur                    As Long       ' Form current width or height
        P_Init                     As Long       ' Point initial value
        P_New                      As Long       ' Point newly calculated value
End Type

' Structure that holds Form/Control information
Type FLY_DATA
        OldProc                    As WNDPROC    ' address to old window procedure
        hWndControl                As HWND       ' handle of the control
        hWndParent                 As HWND       ' handle of the parent
        IsForm                     As Long       ' flag indicating that this is a Form
        IsModal                    As Long       ' flag indicating modal form
        hFont                      As HFONT      ' handle to font
        hAccel                     As HACCEL     ' handle to accelerator table
        hBackBitmap                As HBITMAP    ' handle of background form bitmap
        hStatusBarTimer            as Long       ' handle of timer for statusbar
        ControlIndex               As Long       ' controlindex of control
        CtrlFocus                  As HWND       ' control with current focus
        SelText                    As Long       ' flag to highlight text in edit controls
        ProcessCtlColor            As Long       ' flag to process color messages
        IsForeSysColor             As Long       ' flag indicating to use system color
        IsBackSysColor             As Long       ' flag indicating to use system color
        hBackBrush                 As HBRUSH     ' brush for background color
        nForeColor                 As Long       ' foreground color
        nBackColor                 As Long       ' background color
        fResizeInit                As Long       ' flag signaling that resize has been initialized
        rcForm                     As Rect       ' Rectangle holding the Form's client dimensions (resizing)
        rcCtrl                     As Rect       ' Rectangle of the control in client relative dimensions (resizing)
        zConstraints               As zString * 30 ' resize constraints
End Type

' Common info that the programmer can access via the Shared APP variable.
Type APP_TYPE
        Comments                   As zString * MAX_PATH 	' Comments
        CompanyName                As zString * MAX_PATH 	' Company Name
        EXEName                    As zString * MAX_PATH 	' EXE name of program
        FileDescription            As zString * MAX_PATH 	' File Description
        hInstance                  As HINSTANCE  			' Instance handle of the program
        Path                       As zString * MAX_PATH 	' Current Path to the EXE
        ProductName                As zString * MAX_PATH 	' Product Name
        LegalCopyright             As zString * MAX_PATH 	' Legal Copyright
        LegalTrademarks            As zString * MAX_PATH 	' Legal Trademarks
        ProductMajor               As Long       			' Product Major number
        ProductMinor               As Long       			' Product Minor number
        ProductRevision            As Long       			' Product Revision number
        ProductBuild               As Long       			' Product Build number
        FileMajor                  As Long       			' File Major number
        FileMinor                  As Long       			' File Minor number
        FileRevision               As Long       			' File Revision number
        FileBuild                  As Long       			' File Build number
        ReturnValue                As Long       			' User value returned from FF_FormClose
End Type
Dim Shared App        As APP_TYPE

Dim Shared gFLY_AppID As zString *MAX_PATH
Dim Shared gFLY_FontHandles() As HFONT
Dim Shared gFLY_hDlgCurrent As HWND

Dim Shared HWND_FORM1 As HWND
Dim Shared IDC_FORM1_CHECK7 As Long
Dim Shared HWND_FORM1_CHECK7 As HWND
Dim Shared IDC_FORM1_CHECK6 As Long
Dim Shared HWND_FORM1_CHECK6 As HWND
Dim Shared IDC_FORM1_CHECK5 As Long
Dim Shared HWND_FORM1_CHECK5 As HWND
Dim Shared IDC_FORM1_TEXT3 As Long
Dim Shared HWND_FORM1_TEXT3 As HWND
Dim Shared IDC_FORM1_LABEL1 As Long
Dim Shared HWND_FORM1_LABEL1 As HWND
Dim Shared IDC_FORM1_TEXT2 As Long
Dim Shared HWND_FORM1_TEXT2 As HWND
Dim Shared IDC_FORM1_PICTURE1 As Long
Dim Shared HWND_FORM1_PICTURE1 As HWND
Dim Shared IDC_FORM1_CHECK3 As Long
Dim Shared HWND_FORM1_CHECK3 As HWND
Dim Shared IDC_FORM1_CHECK4 As Long
Dim Shared HWND_FORM1_CHECK4 As HWND
Dim Shared IDC_FORM1_COMMAND1 As Long
Dim Shared HWND_FORM1_COMMAND1 As HWND
Dim Shared IDC_FORM1_TEXT1 As Long
Dim Shared HWND_FORM1_TEXT1 As HWND
Dim Shared IDC_FORM1_CHECK1 As Long
Dim Shared HWND_FORM1_CHECK1 As HWND
Dim Shared IDC_FORM1_CHECK2 As Long
Dim Shared HWND_FORM1_CHECK2 As HWND
Dim Shared IDC_FORM1_CHECK8 As Long
Dim Shared HWND_FORM1_CHECK8 As HWND
Dim Shared IDC_FORM1_CHECK81 As Long
Dim Shared HWND_FORM1_CHECK81 As HWND
Dim Shared IDC_FORM1_CHECK9 As Long
Dim Shared HWND_FORM1_CHECK9 As HWND
Dim Shared IDC_FORM1_TEXT4 As Long
Dim Shared HWND_FORM1_TEXT4 As HWND
Dim Shared IDC_FORM1_LABEL2 As Long
Dim Shared HWND_FORM1_LABEL2 As HWND

Declare Function FORM1_Show(ByVal hWndParent As HWND , _
        ByVal ShowModalFlag As Long , _
        ByVal UserData As long = 0 _
        ) As HWND
Declare Sub FORM1_CreateControls(ByVal hWndForm As HWND)
Declare Function FORM1_FORMPROCEDURE(ByVal hWndForm As HWND , _
        ByVal wMsg As uInteger , _
        ByVal wParam As WPARAM , _
        ByVal lParam As LPARAM _
        ) As LRESULT
Declare Function FORM1_CODEPROCEDURE(ByVal hWndControl As HWND , _
        ByVal wMsg As uInteger , _
        ByVal wParam As WPARAM , _
        ByVal lParam As LPARAM _
        ) As LRESULT

Declare Function FORM1_COMMAND1_BN_CLICKED( _
        ByVal ControlIndex as Long , _
        ByVal hWndForm as HWnd , _
        ByVal hWndControl as HWnd , _
        ByVal idButtonControl as Long _
        ) as Long
Declare Function FORM1_WM_CREATE( _
        ByVal hWndForm as HWnd , _
        ByVal UserData as Long _
        ) as Long
Declare Function FF_WINMAIN(ByVal hInstance As HINSTANCE , _
        ByVal hPrevInstance as HINSTANCE , _
        ByRef lpCmdLine as String , _
        ByVal iCmdShow as Long) as Long
Declare Function FF_PUMPHOOK(ByRef uMsg As Msg) As Long
Declare Sub FLY_InitializeVariables()
Declare Sub FLY_SetAppVariables()
Declare Function FLY_AdjustWindowRect(ByVal hWndForm As HWND , _
        ByVal cxClient As Long , _
        ByVal cyClient As Long _
        ) As Long
Declare Function FLY_EnumSysColorChangeProc(ByVal hWnd As HWND , ByVal lParam As LPARAM) As Long
Declare Function FLY_SetControlData(ByVal hWndControl As HWND , _
        ByVal AllowSubclass As Long , _
        ByVal AllowSetFont As Long , _
        ByRef sFontString As String , _
        ByVal nControlIndex As Long , _
        ByVal nProcessColor As Long , _
        ByVal IsForeSysColor As Long , _
        ByVal IsBackSysColor As Long , _
        ByVal nForeColor As Long , _
        ByVal nBackColor As Long , _
        ByVal nTransparentBrush As HBRUSH , _
        ByVal CodeProcedure As WNDPROC , _
        ByRef sResizeRules As String , _
        ByVal nFontUpgrade As Long _
        ) As FLY_DATA ptr
Declare Function FLY_DoMessagePump(ByVal ShowModalFlag As Long , _
        ByVal hWndForm As HWND , _
        ByVal hWndParent As HWND , _
        ByVal nFormShowState As Long , _
        ByVal IsMDIForm As Long _
        ) As HWND
Declare Function FLY_GetActiveHandle(ByVal hWnd As HWND) as HWND
Declare Function FLY_ResizeRuleInitEnum(ByVal hWnd As HWND , ByVal lParam As LPARAM) As Long
Declare Function FLY_ResizeRuleEnum(ByVal hWnd As HWND , ByVal lParam As LPARAM) As Long
Declare Function WinMain(ByVal hInstance As HINSTANCE , _
        ByVal hPrevInstance As HINSTANCE , _
        ByRef lpCmdLine As String , _
        ByVal iCmdShow As Long _
        ) As Long
Declare Function FF_Main(ByVal hInstance As HINSTANCE , ByVal fwdReason As long) As long

Declare Function FF_Control_GetCheck(ByVal hWndControl As HWND) As Long

Declare Sub FF_Control_GetColor(ByVal hWndControl as HWnd , _
        ByRef ForeColor as Long , _
        ByRef BackColor as Long)

Declare Function FF_Control_GetTag(ByVal hWndControl As HWND) As String

Declare Function FF_Control_GetTag2(ByVal hWndControl As HWND) As String

Declare Function FF_Control_GetText(ByVal hWndControl As HWND) As String

Declare Sub FF_Control_SetCheck(ByVal hWndControl as HWnd , _
        ByVal nCheckState as Long)

Declare Sub FF_Control_SetColor(ByVal hWndControl as HWnd , _
        ByVal ForeColor as Long , _
        ByVal BackColor as Long)

Declare Sub FF_Control_SetTag(ByVal hWnd As HWND , _
        ByRef NewTag As String)

Declare Sub FF_Control_SetTag2(ByVal hWnd As HWND , _
        ByRef NewTag As String)

Declare Sub FF_CloseForm(ByVal hwndForm As HWND , _
        ByVal ReturnValue As Integer = 0)

Declare Function FF_AfxGetFontPointSize(ByVal nHeight As Long) As Long
Declare Function FF_GetFontInfo(ByRef sCurValue As String , _
        ByRef sFontName As String , _
        ByRef nPointSize As Long , _
        ByRef nWeight As Long , _
        ByRef nItalic As Long , _
        ByRef nUnderline As Long , _
        ByRef nStrikeOut As Long , _
        ByRef nFontStyle As Long _
        ) As Long
Declare Function FF_EnumCharSet( _
        ByRef elf As ENUMLOGFONT , _
        ByRef ntm As NEWTEXTMETRIC , _
        ByVal FontType As Long , _
        ByVal CharSet As Long _
        ) As Long
Declare Function FF_MakeFontEX(ByRef sFont As String , _
        ByVal PointSize As Long , _
        ByVal fBold As Long , _
        ByVal fItalic As Long , _
        ByVal fUnderline As Long , _
        ByVal StrikeThru As Long _
        ) As HFONT
Declare Function FF_MakeFontEx_Internal(ByRef sFont As String) As HFONT


Declare Function FF_Parse(ByRef sMainString as String , _
        ByRef sDelimiter as String , _
        ByVal nPosition as Integer _
        ) as String

Declare Function FF_Parse_Internal(ByRef sMainString as String , _
        ByRef sDelimiter as String , _
        ByRef nPosition as Integer , _
        ByRef nIsAny as Integer , _
        ByRef nLenDelimiter as Integer _
        ) as String

Type MyRand32
    Private:
        As Ulong intern, iseed, icount
		Declare Function permuteQPR(byval as ulong) As Ulong
    Public:
        Declare Constructor()
        Declare Constructor(byval as Ulong)
        Declare Destructor()
        Declare Function ranu() As Ulong
		Declare Function ranu(byval as Ulong) As Ulong
        Declare Function rand() As double
        Declare Function valu(byval as Ulong) As Ulong
        Declare Function vald(byval as double) As double
        Declare Sub seed(byval As Ulong)
		Declare Function GetSeed(byref ucount As Ulong = 0) as Ulong
		Declare Function urange(byval as Ulong, byval as Ulong) As Ulong
End Type

Private Constructor MyRand32()
	seed(100)
End Constructor

Private Constructor MyRand32(byval iv as Ulong)
	if iv = 0 THEN
		seed(100)
	else
		seed(iv)
	END IF
End Constructor

Private Destructor MyRand32()

End Destructor

Private Function MyRand32.ranu() As Ulong
	if intern < 4294967291 then
		intern += 3
	else
		intern = 1
	end if
    dim as ulong lint = permuteQPR(intern)
	icount += 1
	return permuteQPR(lint xor &H5bf03635)
End Function

Private Function MyRand32.ranu(byval z as Ulong) As Ulong
	if z = 0 THEN return ranu()
	intern = z
	iseed = z
    dim as ulong lint = permuteQPR(intern)
	icount = 0
	return permuteQPR(lint xor &H5bf03635)
End Function

Private Function MyRand32.rand() as double
    return ranu() / &HFFFFFFFFull        '4294967295
End Function

Private Function MyRand32.valu(byval imax As Ulong) As Ulong
    Return CULng(vald(imax ))
End Function

Private Function MyRand32.vald(byval dmax as double) as double
    return rand() * dmax           '4294967295
End Function

Private Sub MyRand32.seed(Byval z As Ulong)
    intern = z
	iseed = z
	icount = 0
End Sub

Private Function MyRand32.GetSeed(byref ucount As Ulong) as Ulong
	ucount = icount
	return iseed
End Function

Private Function MyRand32.uRange(byval mini as Ulong, byval maxi as Ulong) As Ulong
	dim as double dlint = rand()
	dim as ulong dif = maxi - mini
	dlint = (dlint * dif) + mini
	return  CULng(dlint)
End Function

Private Function MyRand32.permuteQPR(byval x as ulong) as ulong
    return rnd(iseed) * x
end function

private function MyRand32_RandomString(byval ival as ulong, byval iseed as Ulong , byref model0 as string, byval irep as long = 0) as string
	dim as string model = model0
	dim v as ulong = len(model)
	dim l as ulong
	dim k as ulong
	dim r as ulong
    dim as string res
	dim as string stemp
	dim as MyRand32 rd = iseed xor 2413687947

	while v < ival * 5
		model &= model
		v = len(model)
    WEND
	'? model
	'? v
	for l = 1 to ival
		k = rd.uRange(1, v)
		stemp = mid(model, k, 1)
		if irep = 1 and l > 1 THEN
			r= 0
			while  stemp = mid(res, l - 1, 1) and r < 8
				k = rd.uRange(1, v)
				stemp = mid(model, k, 1)
				r +=1
			wend
			if stemp = mid(res, l - 1, 1) THEN return res
        END IF
		res &= stemp
	NEXT
	'? res
    return res
end function

Private function get_Useed() as Ulong
	dim as double t = timer()
	randomize
	for i as integer = 0 to 5
		for j as integer = 0 to 2214 * rnd
			t += timer
		next
	NEXT
	t = t * 146789 * rnd
	while  t < 4503599627370495
		t *= 10.749
    wend
	while  t > 268435455
		t /= 10.4125
    wend
	return CULng(t)
End Function

Private Function FF_Remain_Internal(ByRef sMainString as String , _
            ByRef sMatchPattern as String _
            ) as String
    Dim nLenMain as Long = Len(sMainString)
    Dim i                 as Long
    i = InStr(sMainString , sMatchPattern)
    If i Then
        Function = Mid(sMainString , i + 1)
    Else
        Function = ""
    End If
End Function


' =====================================================================================
' Get the Windows version (based on code from Jose Roca)
' =====================================================================================
Private FUNCTION AfxGetWindowsVersion() AS Single
    Dim dwVersion         AS Long
    Dim nMajorVer         As Long
    Dim nMinorVer         As Long
    dwVersion = GetVersion
    nMajorVer = LOBYTE(LOWORD(dwVersion))
    nMinorVer = HIBYTE(LOWORD(dwVersion))
    FUNCTION = nMajorVer + (nMinorVer / 100)
END FUNCTION


' =====================================================================================
' Scales an horizontal coordinate according the DPI (dots per pixel) being used by the application.
' Based on code from Jsoe Roca
' =====================================================================================
Private Function AfxScaleX(BYVAL cx AS SINGLE) AS SINGLE
    #IfDef FF_NOSCALE
        Function = cx
    #Else
        Dim hDC               As HDC
        hDC = GetDC(Null)
        Function = cx * (GetDeviceCaps(hDC , LOGPIXELSX) / 96)
        ReleaseDC Null , hDC
    #EndIf
End Function


' =====================================================================================
' Scales a vertical coordinate according the DPI (dots per pixel) being used by the application.
' Based on code from Jose Roca
' =====================================================================================
Private Function AfxScaleY(BYVAL cy AS SINGLE) AS SINGLE
    #IfDef FF_NOSCALE
        Function = cy
    #Else
        Dim hDC               As HDC
        hDC = GetDC(Null)
        Function = cy * (GetDeviceCaps(hDC , LOGPIXELSY) / 96)
        ReleaseDC Null , hDC
    #EndIf
End Function
' =====================================================================================


' =====================================================================================
' Based on code from Jose Roca
' Creates a standard tooltip for a window's entire client area.
' Parameters:
' - hwnd = Handle to the window
' - strTooltipText = Tooltip text
' - bBalloon = Ballon tip (TRUE or FALSE)
' Return Value:
' The handle of the tooltip control
' =====================================================================================
Private Function FF_AddTooltip(BYVAL hwnd AS HWND , ByRef strTooltipText AS STRING , BYVAL bBalloon AS Long) As HWND

    IF hwnd = 0 Then Exit Function

    Dim hwndTT            AS HWND
    Dim dwStyle           As Long

    dwStyle = WS_POPUP OR TTS_NOPREFIX OR TTS_ALWAYSTIP
    IF bBalloon THEN dwStyle = dwStyle OR TTS_BALLOON
    hwndTT = CreateWindowEx(WS_EX_TOPMOST , "tooltips_class32" , "" , dwStyle , 0 , 0 , 0 , 0 , 0 , _
            Cast(HMENU , Null) , 0 , ByVal Cast(LPVOID , Null))
    IF hwndTT = 0 THEN Exit Function
    SetWindowPos(hwndTT , HWND_TOPMOST , 0 , 0 , 0 , 0 , SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOACTIVATE)

    ' // Register the window with the tooltip control
    Dim tti               AS TOOLINFO
    tti.cbSize = SIZEOF(tti)
    tti.uFlags = TTF_SUBCLASS
    tti.hwnd = hwnd
    tti.hinst = GetModuleHandle(BYVAL NULL)

    GetClientRect(hwnd , Varptr(tti.rect))

    ' // The length of the string must not exceed of 80 characters, including the terminating null
    IF LEN(strTooltipText) > 79 THEN strTooltipText = LEFT(strTooltipText , 79)
    tti.lpszText = STRPTR(strTooltipText)
    tti.uId = 0
    SendMessage hwndTT , TTM_ADDTOOL , 0 , Cast(LPARAM , Varptr(tti))
    Function = hwndTT

End Function

Private Function FF_Control_GetCheck(ByVal hWndControl As HWND) As Long

    ' Do a check to ensure that this is actually a window handle
    If IsWindow(hWndControl) Then

        ' Get the check state
        If SendMessage(hWndControl , BM_GETCHECK , 0 , 0) = BST_CHECKED Then
            Function = True
        End If

    End If

End Function

Private Sub FF_Control_GetColor(ByVal hWndControl as HWnd , _
            ByRef ForeColor as Long , _
            ByRef BackColor as Long)

    Dim ff                as FLY_DATA Ptr
    Dim hBrush            as HBRUSH
    Dim LOG_BRUSH         as LOGBRUSH

    ' Do a check to ensure that this is actually a window handle
    If IsWindow(hWndControl) Then

        ff = GetProp(hWndControl , "FLY_PTR")

        If ff Then
            ' Determine is we are dealing with a Form or Control
            If ff -> IsForm Then
                hBrush = Cast(HBRUSH , GetClassLongPtr(hWndControl , GCLP_HBRBACKGROUND))
                'Get the Color From the brush:
                GetObject hBrush , SizeOf(LOG_BRUSH) , ByVal VarPtr(LOG_BRUSH)
                ForeColor = 0
                BackColor = LOG_BRUSH.lbColor
                Exit Sub
            End If

            If ff -> ProcessCtlColor Then
                If ff -> IsForeSysColor Then
                    ForeColor = GetSysColor(ff -> nForeColor)
                Else
                    ForeColor = ff -> nForeColor
                End If

                If ff -> IsBackSysColor Then
                    BackColor = GetSysColor(ff -> nBackColor)
                Else
                    BackColor = ff -> nBackColor
                End If

            Else
                ForeColor = GetSysColor(ff -> nForeColor)
                BackColor = GetSysColor(ff -> nBackColor)
            End If

        End If

    End If
End Sub

Private Function FF_Control_GetTag(ByVal hWndControl As HWND) As String

    Dim pzString          As ZString Ptr

    'Get the pointer to the Tag string from the hWnd PROP.
    pzString = GetProp(hWndControl , "FLY_TAGPROPERTY")

    If pzString = 0 Then
        'Null pointer
        Function = ""
    Else
        Function = *pzString
    End If

End Function

Private Function FF_Control_GetTag2(ByVal hWndControl As HWND) As String

    Dim pzString          As ZString Ptr

    'Get the pointer to the Tag2 string from the hWnd PROP.
    pzString = GetProp(hWndControl , "FLY_TAGPROPERTY2")

    If pzString = 0 Then
        'Null pointer
        Function = ""
    Else
        Function = *pzString
    End If

End Function



Private Function FF_Control_GetText(ByVal hWndControl As HWND) As String

    Dim nBufferSize       As Long
    Dim nBuffer           As String

    ' Do a check to ensure that this is actually a window handle
    If IsWindow(hWndControl) Then

        ' Get the length of the text
        nBufferSize = GetWindowTextLength(hWndControl)
        If nBufferSize = 0 Then
            Function = "" : Exit Function
        End If

        ' Add an extra character for the Nul terminator
        nBufferSize = nBufferSize + 1

        ' Create the temporary buffer
        nBuffer = Space(nBufferSize)

        ' Retrieve the text
        GetWindowText hWndControl , ByVal StrPtr(nBuffer) , nBufferSize

        ' Remove the Null
        nBuffer = RTrim(nBuffer , chr(0))

        Function = nBuffer

    End If


End Function



Private Sub FF_Control_SetCheck(ByVal hWndControl as HWnd , _
            ByVal nCheckState as Long)

    ' Do a check to ensure that this is actually a window handle
    If IsWindow(hWndControl) Then

        ' Set the check state
        If nCheckState Then
            SendMessage hWndControl , BM_SETCHECK , BST_CHECKED , 0
        Else
            SendMessage hWndControl , BM_SETCHECK , BST_UNCHECKED , 0
        End If

    End If

End Sub



Private Sub FF_Control_SetColor(ByVal hWndControl as HWnd , _
            ByVal ForeColor as Long , _
            ByVal BackColor as Long)

    Dim ff_control        as FLY_DATA Ptr
    Dim Redraw_Flag       as Long

    ' Do a check to ensure that this is actually a window handle
    If IsWindow(hWndControl) Then

        ff_control = GetProp(hWndControl , "FLY_PTR")

        If ff_control Then

            ' Determine is we are dealing with a Form or Control
            If ff_control -> IsForm Then

                ' ForeColor has no effect on a Form

                ' Set the Background color
                If BackColor <> - 1 Then
                    If ff_control -> hBackBrush Then
                        DeleteObject ff_control -> hBackBrush
                    End If
                    ff_control -> hBackBrush = CreateSolidBrush(BackColor)
                    ff_control -> nBackColor = BackColor
                    ff_control -> IsBackSysColor = False
                    SetClassLongPtr hWndControl , GCLP_HBRBACKGROUND , Cast(Integer , CreateSolidBrush(BackColor))
                    Redraw_Flag = True
                End If

            Else

                ' Set the Foreground color
                If ForeColor <> - 1 Then
                    ff_control -> nForeColor = ForeColor
                    ff_control -> IsForeSysColor = False
                    Redraw_Flag = True
                End If

                ' Set the Background color
                If BackColor <> - 1 Then
                    If ff_control -> hBackBrush Then
                        DeleteObject ff_control -> hBackBrush
                    End If
                    ff_control -> hBackBrush = CreateSolidBrush(BackColor)
                    ff_control -> nBackColor = BackColor
                    ff_control -> IsBackSysColor = False
                    Redraw_Flag = True
                End If

            End If

            If Redraw_Flag Then
                InvalidateRect hWndControl , ByVal 0 , TRUE
                UpdateWindow hWndControl
            End If

        End If

    End If
End Sub



Private Sub FF_Control_SetTag(ByVal hWnd As HWND , _
            ByRef NewTag As String)

    If IsWindow(hWnd) = 0 Then Exit Sub

    Dim pzString          As ZString Ptr

    ' Get the pointer to the Tag string from the hWnd PROP.
    pzString = GetProp(hWnd , "FLY_TAGPROPERTY")

    ' Free any currently allocated memory
    If pzString Then HeapFree GetProcessHeap() , 0 , ByVal pzString

    If Right(NewTag , 1) <> Chr(0) Then NewTag = NewTag & Chr(0)

    pzString = HeapAlloc(GetProcessHeap() , HEAP_ZERO_MEMORY , Len(NewTag) + 1)
    If pzString Then * pzString = NewTag

    SetProp hWnd , "FLY_TAGPROPERTY" , pzString

End Sub


Private Sub FF_Control_SetTag2(ByVal hWnd As HWND , _
            ByRef NewTag As String)

    If IsWindow(hWnd) = 0 Then Exit Sub

    Dim pzString          As ZString Ptr

    ' Get the pointer to the Tag string from the hWnd PROP.
    pzString = GetProp(hWnd , "FLY_TAGPROPERTY2")

    ' Free any currently allocated memory
    If pzString Then HeapFree GetProcessHeap() , 0 , ByVal pzString

    If Right(NewTag , 1) <> Chr(0) Then NewTag = NewTag & Chr(0)

    pzString = HeapAlloc(GetProcessHeap() , HEAP_ZERO_MEMORY , Len(NewTag) + 1)
    If pzString Then * pzString = NewTag

    SetProp hWnd , "FLY_TAGPROPERTY2" , pzString

End Sub


Private Sub FF_CloseForm(ByVal hwndForm As HWND , _
            ByVal ReturnValue As Integer = 0)

    If IsWindow(hWndForm) = False Then Exit Sub

    Dim hWndParent        As HWND

    Dim ff                As FLY_DATA Ptr
    ff = GetProp(hwndForm , "FLY_PTR")

    If ff Then
        If ff -> IsModal Then
            ' Reset the focus back to the parent of the modal form.
            ' Enable Mouse and keyboard input for the Parent Form
            hWndParent = Cast(HWND , GetWindowLongPtr(hWndForm , GWLP_HWNDPARENT))
            If IsWindow(hWndParent) Then
                EnableWindow hWndParent , True
                SetActiveWindow hWndParent
            End If
        End If
    End If

    'Destroy the window. This will delete any font objects and release
    'the memory held by the type structure.
    DestroyWindow hWndForm

    'Set the value that will be returned to the calling program.
    App.ReturnValue = ReturnValue
End Sub


' ========================================================================================
' Returns the point size of a font given its logical height
' Based on code from Jose Roca
' ========================================================================================
Private Function FF_AfxGetFontPointSize(ByVal nHeight As Long) As Long
    Dim hdc               As HDC
    hdc = CreateDC( "DISPLAY" , ByVal Null , ByVal Null , ByVal Null)
    If hdc = Null Then Exit Function
    Dim cyPixelsPerInch   As Long
    cyPixelsPerInch = GetDeviceCaps(hdc , LOGPIXELSY)
    DeleteDC hdc
    Dim nPointSize        As Long
    nPointSize = MulDiv(nHeight , 72 , cyPixelsPerInch)
    If nPointSize < 0 Then nPointSize = - nPointSize
    Function = nPointSize
End Function


Private Function FF_GetFontInfo(ByRef sCurValue As String , _
            ByRef sFontName As String , _
            ByRef nPointSize As Long , _
            ByRef nWeight As Long , _
            ByRef nItalic As Long , _
            ByRef nUnderline As Long , _
            ByRef nStrikeOut As Long , _
            ByRef nFontStyle As Long _
            ) As Long

    ' FireFly handles two types of font strings. The first is the traditional
    ' logfont string and the second is the simplified PB version. FireFly is moving
    ' towards the simplified version but needs to handle the logfont version
    ' for backwards compatibility.

    ' Count the number of commas in the incoming string
    Dim nCount            As Integer
    Dim i                 As Integer

    For i = 0 To Len(sCurValue) - 1
        If sCurValue[i] = 44 Then nCount = nCount + 1
    Next

    If nCount > 10 Then
        ' This must be the old style logfont structure
        ' "Tahoma,-11,0,0,0,400,0,0,0,0,3,2,1,34" ' 400 = normal, 700 = bold
        sFontName = FF_Parse(sCurValue , "," , 1)
        nPointSize = FF_AfxGetFontPointSize(ValInt(FF_Parse(sCurValue , "," , 2)))
        nWeight = ValInt(FF_Parse(sCurValue , "," , 6))
        nItalic = ValInt(FF_Parse(sCurValue , "," , 7))
        nUnderline = ValInt(FF_Parse(sCurValue , "," , 8))
        nStrikeOut = ValInt(FF_Parse(sCurValue , "," , 9))
        If nWeight = FW_BOLD Then nFontStyle = nFontStyle + 1 ' normal/bold
        If nItalic Then nFontStyle = nFontStyle + 2 ' italic
        If nUnderline Then nFontStyle = nFontStyle + 4 ' underline
        If nStrikeOut Then nFontStyle = nFontStyle + 8 ' strikeout
    Else
        ' This is the new style (3 parts)
        sFontName = FF_Parse(sCurValue , "," , 1)
        nPointSize = ValInt(FF_Parse(sCurValue , "," , 2))
        nFontStyle = ValInt(FF_Parse(sCurValue , "," , 3))
        nWeight = FW_NORMAL
        If (nFontStyle And 1) Then nWeight = FW_BOLD
        If (nFontStyle And 2) Then nItalic = TRUE
        If (nFontStyle And 4) Then nUnderline = TRUE
        If (nFontStyle And 8) Then nStrikeOut = TRUE
    End If

    Function = 0

End Function


Private Function FF_EnumCharSet( _
            ByRef elf As ENUMLOGFONT , _
            ByRef ntm As NEWTEXTMETRIC , _
            ByVal FontType As Long , _
            ByVal CharSet As Long _
            ) As Long

    CharSet = elf.elfLogFont.lfCharSet
    Function = TRUE
End Function


Private Function FF_MakeFontEX(ByRef sFont As String , _
            ByVal PointSize As Long , _
            ByVal fBold As Long , _
            ByVal fItalic As Long , _
            ByVal fUnderline As Long , _
            ByVal StrikeThru As Long _
            ) As HFONT

    Dim tlf               As LOGFONT
    Dim hDC               As HDC
    Dim CharSet           As Integer
    Dim m_DPI             As Integer

    ' Create a high dpi aware font
    If Len(sFont) = 0 Then Exit Function

    hDC = GetDC(0)

    EnumFontFamilies hDC , ByVal StrPtr(sFont) , Cast(FONTENUMPROC , @FF_EnumCharSet) , ByVal Cast(LPARAM , VarPtr(CharSet))

    m_DPI = GetDeviceCaps(hDC , LOGPIXELSX)
    PointSize = (PointSize * m_DPI) \ GetDeviceCaps(hDC , LOGPIXELSY)

    tlf.lfHeight = - MulDiv(PointSize , GetDeviceCaps(hDC , LOGPIXELSY) , 72) ' logical font height
    tlf.lfWidth = 0                              ' average character width
    tlf.lfEscapement = 0                         ' escapement
    tlf.lfOrientation = 0                        ' orientation angles
    tlf.lfWeight = fBold                         ' font weight
    tlf.lfItalic = fItalic                       ' italic(TRUE/FALSE)
    tlf.lfUnderline = fUnderline                 ' underline(TRUE/FALSE)
    tlf.lfStrikeOut = StrikeThru                 ' strikeout(TRUE/FALSE)
    tlf.lfCharSet = Charset                      ' character set
    tlf.lfOutPrecision = OUT_TT_PRECIS           ' output precision
    tlf.lfClipPrecision = CLIP_DEFAULT_PRECIS    ' clipping precision
    tlf.lfQuality = DEFAULT_QUALITY              ' output quality
    tlf.lfPitchAndFamily = FF_DONTCARE           ' pitch and family
    tlf.lfFaceName = sFont                       ' typeface name

    ReleaseDC 0 , hDC

    Function = CreateFontIndirect(@tlf)

End Function



Private Function FF_MakeFontEx_Internal(ByRef sFont As String) As HFONT

    Dim hFont             As HFONT
    Dim sFontName         As String
    Dim nPointSize        As Long
    Dim nWeight           As Long
    Dim nItalic           As Long
    Dim nUnderline        As Long
    Dim nStrikeOut        As Long
    Dim nFontStyle        As Long

    FF_GetFontInfo sFont , sFontName , nPointSize , nWeight , nItalic , nUnderline , nStrikeOut , nFontStyle
    hFont = FF_MakeFontEx(sFontName , nPointSize , nWeight , nItalic , nUnderline , nStrikeOut)

    ' Add this font to the global shared font array
    Dim nSlot As Long = - 1
    For x As Long = LBound(gFLY_FontHandles) To Ubound(gFLY_FontHandles)
        If gFLY_FontHandles(x) = 0 Then
            nSlot = x : Exit For
        End If
    Next
    If nSlot = - 1 Then                          ' could not find an empty slot
        nSlot = Ubound(gFLY_FontHandles) + 1
        ReDim Preserve gFLY_FontHandles(nSlot) As HFONT
    End If
    gFLY_FontHandles(nSlot) = hFont

    Function = hFont

End Function



''
'' FF_PARSE
'' Return a delimited field from a string expression.
''
'' Delimiter contains a string of one or more characters that must be fully matched to be successful.
'' If nPosition evaluates to zero or is outside of the actual field count, an empty string is returned.
'' If nPosition is negative then fields are searched from the right to left of the MainString.
'' Delimiters are case-sensitive,
''
Private Function FF_Parse(ByRef sMainString as String , _
            ByRef sDelimiter as String , _
            ByVal nPosition as Integer _
            ) as String
    ' The parse must match the entire deliminter string
    Function = FF_Parse_Internal(sMainString , sDelimiter , nPosition , False , Len(sDelimiter))
End Function


''
'' FF_PARSE_INTERNAL
'' Used by both FF_Parse and FF_ParseAny (internal function)
''
Private Function FF_Parse_Internal(ByRef sMainString as String , _
            ByRef sDelimiter as String , _
            ByRef nPosition as Integer , _
            ByRef nIsAny as Integer , _
            ByRef nLenDelimiter as Integer _
            ) as String

    ' Returns the nPosition-th substring in a string sMainString with separations sDelimiter (one or more charcters),
    ' beginning with nPosition = 1

    Dim         as Integer i
    Dim         as Integer j
    Dim         as Integer count
    Dim s                 as String
    Dim fReverse as Integer = IIf(nPosition < 0 , True , False)

    nPosition = Abs(nPosition)
    count = 0
    i = 1
    s = sMainString

    If fReverse Then
        ' Reverse search
        ' Get the start of the token (j) by searching in reverse
        If nIsAny Then
            i = InStrRev(sMainString , Any sDelimiter)
        Else
            i = InStrRev(sMainString , sDelimiter)
        End If
        Do While i > 0                           ' if not found loop will be skipped
            j = i + nLenDelimiter
            count += 1
            i = i - nLenDelimiter
            If count = nPosition Then Exit Do
            If nIsAny Then
                i = InStrRev(sMainString , Any sDelimiter , i)
            Else
                i = InStrRev(sMainString , sDelimiter , i)
            End If
        Loop
        If i = 0 Then j = 1

        ' Now continue forward to get the end of the token
        If nIsAny Then
            i = InStr(j , sMainString , Any sDelimiter)
        Else
            i = InStr(j , sMainString , sDelimiter)
        End If
        If (i > 0) Or (count = nPosition) Then
            If i = 0 Then
                s = Mid(sMainString , j)
            Else
                s = Mid(sMainString , j , i - j)
            End If
        End If

    Else
        ' Forward search
        Do
            j = i
            If nIsAny Then
                i = InStr(i , sMainString , Any sDelimiter)
            Else
                i = InStr(i , sMainString , sDelimiter)
            End If
            If i > 0 Then count += 1 : i += nLenDelimiter
        Loop Until(i = 0) Or (count = nPosition)

        If (i > 0) Or (count = nPosition - 1) Then
            If i = 0 Then
                s = Mid(sMainString , j)
            Else
                s = Mid(sMainString , j , i - nLenDelimiter - j)
            End If
        End If
    End If

    Return s

End Function


'------------------------------------------------------------------------------
' FORM1 Auto-Generated FireFly Functions
'------------------------------------------------------------------------------

Private Function FORM1_Show(ByVal hWndParent As HWND , _
            ByVal ShowModalFlag As Long , _
            ByVal UserData As long = 0 _
            ) As HWND



    Dim szClassName       As zString *MAX_PATH
    Dim wce               As WndClassEx

    Dim FLY_nLeft         As Long
    Dim FLY_nTop          As Long
    Dim FLY_nWidth        As Long
    Dim FLY_nHeight       As Long

    Dim hWndForm          As HWND
    Dim IsMDIForm         as Long

    Static IsInitialized  As Long

    'Determine if the Class already exists. If it does not, then create the class.
    szClassName = "FORM_PASS_GENERATOR_FORM1_CLASS"
    wce.cbSize = SizeOf(wce)
    IsInitialized = GetClassInfoEx(App.hInstance , szClassName , varptr(wce))

    If IsInitialized = 0 Then
        wce.cbSize = SizeOf(wce)
        wce.STYLE = CS_VREDRAW Or CS_HREDRAW Or CS_DBLCLKS
        wce.lpfnWndProc = Cast(WNDPROC , @FORM1_FORMPROCEDURE)
        wce.cbClsExtra = 0
        wce.cbWndExtra = 0
        wce.hInstance = App.hInstance
        wce.hIcon = LoadImage(App.hInstance , "IMAGE_KEY_SOLID" , IMAGE_ICON , 32 , 32 , LR_SHARED)
        wce.hCursor = LoadCursor(0 , ByVal IDC_ARROW)
        wce.hbrBackground = Cast(HBRUSH , COLOR_BTNFACE + 1)
        wce.lpszMenuName = 0
        wce.lpszClassName = VarPtr(szClassName)
        wce.hIconSm = LoadImage(App.hInstance , "IMAGE_KEY_SOLID" , IMAGE_ICON , 16 , 16 , LR_SHARED)

        If RegisterClassEx(varptr(wce)) Then IsInitialized = TRUE
    End If


    ' // Set the flag if this is an MDI form we are creating
    IsMDIForm = FALSE

    ' // Save the optional UserData to be checked in the Form Procedure CREATE message
    App.ReturnValue = UserData

    ' Set the size/positioning of the window about to be created
    FLY_nLeft = afxScaleX(200) : FLY_nTop = afxScaleY(150)
    FLY_nWidth = afxScaleX(806) : FLY_nHeight = afxScaleY(425)



    ' Create a window using the registered class
    hWndForm = CreateWindowEx(WS_EX_WINDOWEDGE Or WS_EX_CONTROLPARENT _
            Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR _
             , _
            szClassName , _                              ' window class name
            "Password Generator    (Marpon -  v 1.0)" , _            ' window caption
            WS_POPUP Or WS_THICKFRAME Or WS_CAPTION Or WS_SYSMENU _
            Or WS_MINIMIZEBOX Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN _
             , _
            FLY_nLeft , FLY_nTop , FLY_nWidth , FLY_nHeight , _
            hwndParent , _                               ' parent window handle
            Cast(HMENU , Cast(LONG_PTR , Null)) , _      ' window menu handle/child identifier
            App.hInstance , _                            ' program instance handle
            ByVal Cast(LPVOID , Cast(LONG_PTR , UserData)))

    If IsWindow(hWndForm) = 0 Then
        Function = 0 : Exit Function
    End If

    Function = FLY_DoMessagePump(ShowModalFlag , hWndForm , hWndParent , SW_SHOWNORMAL , IsMDIForm)
End Function



'------------------------------------------------------------------------------
' Create all child controls for FORM1
'------------------------------------------------------------------------------
Private Sub FORM1_CreateControls(ByVal hWndForm As HWND)

    ' All of the child controls for the FORM1 form are
    ' created in this subroutine. This subroutine is called
    ' from the WM_CREATE message of the form.

    Dim FLY_zTempString   As zString *MAX_PATH

    Dim FLY_hPicture      As HANDLE
    Dim hWndControl       As HWND
    Dim FLY_ScrollInfo    As ScrollInfo
    Dim FLY_TBstring      As String
    Dim FLY_Notify        As NMHDR
    Dim FLY_ImageListNormal As HANDLE
    Dim FLY_ImageListDisabled As HANDLE
    Dim FLY_ImageListHot  As HANDLE

    Dim FLY_hStatusBar    As HWND
    Dim FLY_hRebar        As HWND
    Dim FLY_hToolBar      As HWND
    Dim FLY_tempLong      As Long
    Dim FLY_hFont         As HFONT
    Dim FLY_hIcon         As HICON
    Dim FLY_hMemDC        As HDC
    Dim FLY_ImageIndex    As Long
    Dim FLY_ClientOffset  As Long

    Dim FLY_RECT          As Rect
    Dim FLY_TabRect       As Rect

    Dim FLY_TCITEM        As TC_ITEM
    Dim uVersion          As OSVERSIONINFO
    Dim DateRange(1)      As SYSTEMTIME


    Dim ff                As FLY_DATA Ptr
    Dim FLY_child         As FLY_DATA Ptr
    Dim FLY_parent        As FLY_DATA Ptr

    FLY_parent = GetProp(hWndForm , "FLY_PTR")

	'------------------------------------------------------------------------------
    ' Create TEXT4 [TextBox] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING _
            Or WS_EX_RIGHTSCROLLBAR , _
            "Edit" , _
            "" , _
            WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT _
            Or ES_AUTOHSCROLL , _
            afxScaleX(136) , afxScaleY(FLY_ClientOffset + 208) , _
            afxScaleX(497) , afxScaleY(24) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_TEXT4)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Tahoma,10,0" , 0 , TRUE , _
            TRUE , TRUE , COLOR_WINDOWTEXT , _
            COLOR_WINDOW , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_TEXT4 = hWndControl

    SetWindowText hWndControl , ""
    SendMessage hWndControl , EM_SETLIMITTEXT , 0 , 0
    SendMessage hWndControl , EM_SETMARGINS , EC_LEFTMARGIN Or EC_RIGHTMARGIN , MAKELONG(0 , 0)


    ' Set flag for the TextBox control to select all text when gaining focus.
    ff -> SelText = FALSE

	'------------------------------------------------------------------------------
    ' Create CHECK9 [CheckBox] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_LEFT Or WS_EX_LTRREADING , _
            "Button" , _
            "Exclusions" , _
            WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT _
            Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER _
             , _
            afxScaleX(32) , afxScaleY(FLY_ClientOffset + 210) , _
            afxScaleX(101) , afxScaleY(21) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_CHECK9)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Tahoma,10,0" , 0 , TRUE , _
            TRUE , TRUE , COLOR_WINDOWTEXT , _
            COLOR_BTNFACE , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_CHECK9 = hWndControl

    SendMessage hWndControl , BM_SETCHECK , BST_UNCHECKED , 0

	'------------------------------------------------------------------------------
    ' Create CHECK81 [CheckBox] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_LEFT Or WS_EX_LTRREADING , _
            "Button" , _
            "Without repetitions" , _
            WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT _
            Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER _
             , _
            afxScaleX(480) , afxScaleY(FLY_ClientOffset + 254) , _
            afxScaleX(181) , afxScaleY(21) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_CHECK81)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Tahoma,10,0" , 0 , TRUE , _
            TRUE , TRUE , COLOR_WINDOWTEXT , _
            COLOR_BTNFACE , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_CHECK81 = hWndControl

    SendMessage hWndControl , BM_SETCHECK , BST_UNCHECKED , 0
	'------------------------------------------------------------------------------
    ' Create CHECK8 [CheckBox] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_LEFT Or WS_EX_LTRREADING , _
            "Button" , _
            "Without duplicates" , _
            WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT _
            Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER _
             , _
            afxScaleX(283) , afxScaleY(FLY_ClientOffset + 254) , _
            afxScaleX(130) , afxScaleY(21) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_CHECK8)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Tahoma,10,0" , 0 , TRUE , _
            TRUE , TRUE , COLOR_WINDOWTEXT , _
            COLOR_BTNFACE , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_CHECK8 = hWndControl

    SendMessage hWndControl , BM_SETCHECK , BST_UNCHECKED , 0

    '------------------------------------------------------------------------------
    ' Create CHECK7 [CheckBox] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_LEFT Or WS_EX_LTRREADING , _
            "Button" , _
            "Your Customs" , _
            WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT _
            Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER _
             , _
            afxScaleX(32) , afxScaleY(FLY_ClientOffset + 176) , _
            afxScaleX(101) , afxScaleY(21) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_CHECK7)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Tahoma,10,0" , 0 , TRUE , _
            TRUE , TRUE , COLOR_WINDOWTEXT , _
            COLOR_BTNFACE , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_CHECK7 = hWndControl

    SendMessage hWndControl , BM_SETCHECK , BST_UNCHECKED , 0


    '------------------------------------------------------------------------------
    ' Create CHECK6 [CheckBox] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_LEFT Or WS_EX_LTRREADING , _
            "Button" , _
            "  with lower case letters    a b c ..." , _
            WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT _
            Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER _
             , _
            afxScaleX(32) , afxScaleY(FLY_ClientOffset + 40) , _
            afxScaleX(237) , afxScaleY(33) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_CHECK6)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Tahoma,10,0" , 0 , TRUE , _
            TRUE , TRUE , COLOR_WINDOWTEXT , _
            COLOR_BTNFACE , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_CHECK6 = hWndControl

    SendMessage hWndControl , BM_SETCHECK , BST_CHECKED , 0


    '------------------------------------------------------------------------------
    ' Create CHECK5 [CheckBox] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_LEFT Or WS_EX_LTRREADING , _
            "Button" , _
            "  with upper case letters    A B C ..." , _
            WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT _
            Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER _
             , _
            afxScaleX(32) , afxScaleY(FLY_ClientOffset + 16) , _
            afxScaleX(221) , afxScaleY(33) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_CHECK5)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Tahoma,10,0" , 0 , TRUE , _
            TRUE , TRUE , COLOR_WINDOWTEXT , _
            COLOR_BTNFACE , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_CHECK5 = hWndControl

    SendMessage hWndControl , BM_SETCHECK , BST_CHECKED , 0


    '------------------------------------------------------------------------------
    ' Create TEXT3 [TextBox] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING _
            Or WS_EX_RIGHTSCROLLBAR , _
            "Edit" , _
            "" , _
            WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT _
            Or ES_AUTOHSCROLL , _
            afxScaleX(136) , afxScaleY(FLY_ClientOffset + 176) , _
            afxScaleX(497) , afxScaleY(24) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_TEXT3)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Tahoma,10,0" , 0 , TRUE , _
            TRUE , TRUE , COLOR_WINDOWTEXT , _
            COLOR_WINDOW , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_TEXT3 = hWndControl

    SetWindowText hWndControl , ""
    SendMessage hWndControl , EM_SETLIMITTEXT , 0 , 0
    SendMessage hWndControl , EM_SETMARGINS , EC_LEFTMARGIN Or EC_RIGHTMARGIN , MAKELONG(0 , 0)


    ' Set flag for the TextBox control to select all text when gaining focus.
    ff -> SelText = FALSE


    '------------------------------------------------------------------------------
    ' Create LABEL1 [Label] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_LEFT Or WS_EX_LTRREADING , _
            "Static" , _
            "Number of characters" , _
            WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN _
            Or SS_LEFT Or SS_NOTIFY , _
            afxScaleX(92) , afxScaleY(FLY_ClientOffset + 256) , _
            afxScaleX(185) , afxScaleY(19) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_LABEL1)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Tahoma,10,0" , 0 , TRUE , _
            TRUE , TRUE , COLOR_WINDOWTEXT , _
            COLOR_BTNFACE , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_LABEL1 = hWndControl

    '------------------------------------------------------------------------------
    ' Create LABEL2 [Label] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_LEFT Or WS_EX_LTRREADING , _
            "Static" , _
            "Select your choices and click 'Generate' " , _
            WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN _
            Or SS_CENTER Or SS_NOTIFY , _
            afxScaleX(52) , afxScaleY(FLY_ClientOffset + 298) , _
            afxScaleX(565) , afxScaleY(19) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_LABEL2)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Tahoma,12,1" , 0 , TRUE , _
            TRUE , TRUE , COLOR_WINDOWTEXT , _
            COLOR_BTNFACE , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_LABEL2 = hWndControl
    '------------------------------------------------------------------------------
    ' Create TEXT2 [TextBox] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING _
            Or WS_EX_RIGHTSCROLLBAR , _
            "Edit" , _
            "" , _
            WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_CENTER _
            Or ES_AUTOHSCROLL , _
            afxScaleX(32) , afxScaleY(FLY_ClientOffset + 252) , _
            afxScaleX(54) , afxScaleY(26) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_TEXT2)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Tahoma,11,1" , 0 , TRUE , _
            FALSE , TRUE , 16711680 , _
            COLOR_WINDOW , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_TEXT2 = hWndControl

    SetWindowText hWndControl , "12"
    SendMessage hWndControl , EM_SETLIMITTEXT , 0 , 0
    SendMessage hWndControl , EM_SETMARGINS , EC_LEFTMARGIN Or EC_RIGHTMARGIN , MAKELONG(0 , 0)


    ' Set flag for the TextBox control to select all text when gaining focus.
    ff -> SelText = FALSE


    '------------------------------------------------------------------------------
    ' Create PICTURE1 [Picture] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(0 , _
            "Static" , _
            "" , _
            WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS _
            Or WS_CLIPCHILDREN Or SS_ICON Or SS_NOTIFY Or SS_REALSIZEIMAGE _
             , _
            afxScaleX(658) , afxScaleY(FLY_ClientOffset + 25) , _
            afxScaleX(100) , afxScaleY(100) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_PICTURE1)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "" , 0 , FALSE , _
            FALSE , FALSE , 0 , _
            0 , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_PICTURE1 = hWndControl

    GetClientRect hWndControl , VarPtr(FLY_RECT)
    SendMessage hWndControl , STM_SETIMAGE , IMAGE_ICON , Cast(LPARAM , LoadImage(App.hInstance , "IMAGE_KEY_SOLID" , IMAGE_ICON , FLY_RECT.Right , FLY_RECT.Bottom , 32768))


    '------------------------------------------------------------------------------
    ' Create CHECK3 [CheckBox] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_LEFT Or WS_EX_LTRREADING , _
            "Button" , _
            "  with separators               + - _ * / < = > ( { [ ] " & _
            "} ) " , _
            WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT _
            Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER _
             , _
            afxScaleX(32) , afxScaleY(FLY_ClientOffset + 112) , _
            afxScaleX(478) , afxScaleY(33) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_CHECK3)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Tahoma,10,0" , 0 , TRUE , _
            TRUE , TRUE , COLOR_WINDOWTEXT , _
            COLOR_BTNFACE , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_CHECK3 = hWndControl

    SendMessage hWndControl , BM_SETCHECK , BST_UNCHECKED , 0


    '------------------------------------------------------------------------------
    ' Create CHECK4 [CheckBox] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_LEFT Or WS_EX_LTRREADING , _
            "Button" , _
            "  avec specials                  # @ | % && $ \ ^" , _
            WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT _
            Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER _
             , _
            afxScaleX(32) , afxScaleY(FLY_ClientOffset + 136) , _
            afxScaleX(324) , afxScaleY(33) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_CHECK4)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Tahoma,10,0" , 0 , TRUE , _
            TRUE , TRUE , COLOR_WINDOWTEXT , _
            COLOR_BTNFACE , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_CHECK4 = hWndControl

    SendMessage hWndControl , BM_SETCHECK , BST_UNCHECKED , 0


    '------------------------------------------------------------------------------
    ' Create COMMAND1 [CommandButton] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_LEFT Or WS_EX_LTRREADING , _
            "Button" , _
            "Generate" , _
            WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT _
            Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER _
             , _
            afxScaleX(656) , afxScaleY(FLY_ClientOffset + 326) , _
            afxScaleX(100) , afxScaleY(25) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_COMMAND1)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Tahoma,10,0" , 0 , FALSE , _
            FALSE , FALSE , 0 , _
            0 , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_COMMAND1 = hWndControl


    '------------------------------------------------------------------------------
    ' Create TEXT1 [TextBox] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING _
            Or WS_EX_RIGHTSCROLLBAR , _
            "Edit" , _
            "" , _
            WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_CENTER _
            Or ES_AUTOHSCROLL Or ES_READONLY , _
            afxScaleX(32) , afxScaleY(FLY_ClientOffset + 326) , _
            afxScaleX(602) , afxScaleY(28) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_TEXT1)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Times New Roman,14,1" , 0 , TRUE , _
            FALSE , FALSE , 196 , _
            13434879 , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_TEXT1 = hWndControl

    SendMessage hWndControl , EM_SETLIMITTEXT , 0 , 0
    SendMessage hWndControl , EM_SETMARGINS , EC_LEFTMARGIN Or EC_RIGHTMARGIN , MAKELONG(0 , 0)


    ' Set flag for the TextBox control to select all text when gaining focus.
    ff -> SelText = FALSE


    '------------------------------------------------------------------------------
    ' Create CHECK1 [CheckBox] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_LEFT Or WS_EX_LTRREADING , _
            "Button" , _
            "  with digits                      0 1 2 ..." , _
            WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT _
            Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER _
             , _
            afxScaleX(32) , afxScaleY(FLY_ClientOffset + 64) , _
            afxScaleX(252) , afxScaleY(33) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_CHECK1)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Tahoma,10,0" , 0 , TRUE , _
            TRUE , TRUE , COLOR_WINDOWTEXT , _
            COLOR_BTNFACE , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_CHECK1 = hWndControl

    SendMessage hWndControl , BM_SETCHECK , BST_CHECKED , 0


    '------------------------------------------------------------------------------
    ' Create CHECK2 [CheckBox] control.
    '------------------------------------------------------------------------------
    hWndControl = CreateWindowEx(WS_EX_LEFT Or WS_EX_LTRREADING , _
            "Button" , _
            "  with ponctuations            : ; . , ? !" , _
            WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT _
            Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER _
             , _
            afxScaleX(32) , afxScaleY(FLY_ClientOffset + 88) , _
            afxScaleX(260) , afxScaleY(33) , _
            hWndForm , Cast(HMENU , Cast(LONG_PTR , IDC_FORM1_CHECK2)) , _
            App.hInstance , ByVal 0)

    ff = FLY_SetControlData(hWndControl , TRUE , TRUE , _
            "Tahoma,10,0" , 0 , TRUE , _
            TRUE , TRUE , COLOR_WINDOWTEXT , _
            COLOR_BTNFACE , Cast(HBRUSH , - 1) , Cast(WNDPROC , @FORM1_CODEPROCEDURE) , _
            "" , FALSE)

    ' If this control has a ToolTip specified then set it now
    FLY_zTempString = ""
    If Len(FLY_zTempString) then
        FF_AddToolTip hWndControl , FLY_zTempString , FALSE
    End If

    ' Set the Tag properties for the Control
    FF_Control_SetTag hWndControl , ""
    FF_Control_SetTag2 hWndControl , ""

    HWND_FORM1_CHECK2 = hWndControl

    SendMessage hWndControl , BM_SETCHECK , BST_UNCHECKED , 0


    ' Set focus to the lowest control in the Tab Order
    SetFocus HWND_FORM1_COMMAND1

End Sub


'------------------------------------------------------------------------------
' FORM1_FORMPROCEDURE
'------------------------------------------------------------------------------
Private Function FORM1_FORMPROCEDURE(ByVal hWndForm As HWND , _
            ByVal wMsg As uInteger , _
            ByVal wParam As WPARAM , _
            ByVal lParam As LPARAM _
            ) As LRESULT

    Dim FLY_UserData      As Long Ptr
    Dim FLY_pNotify       As NMHDR Ptr
    Dim FLY_pTBN          As TBNOTIFY Ptr
    Dim FLY_lpToolTip     As TOOLTIPTEXT Ptr
    Dim FLY_zTempString   As zString *MAX_PATH
    Dim FLY_TCITEM        As TC_ITEM
    Dim FLY_bm            As Bitmap
    Dim FLY_Rect          As Rect
    Dim FLY_TabRect       As Rect
    Dim FLY_StatusRect    As Rect
    Dim FLY_ToolbarRect   As Rect
    Dim FLY_ProcAddress   As Long
    Dim FLY_hFont         As HFONT
    Dim FLY_hDC           As HDC
    Dim FLY_hForeColor    As Long
    Dim FLY_hBackBrush    As Long
    Dim FLY_hwndParent    As HWND
    Dim FLY_nResult       As HANDLE
    Dim FLY_tempLong      As Long
    Dim FLY_ControlIndex  As Long
    Dim FLY_StatusBarHeight As Long
    Dim FLY_ToolBarHeight As Long

    Dim ff                As FLY_DATA Ptr
    Dim FLY_control       As FLY_DATA Ptr
    If IsWindow(hWndForm) And (wMsg <> WM_CREATE) Then ff = GetProp(hWndForm , "FLY_PTR")


    ' Handle the sizing of any toolbars/statusbars
    If wMsg = WM_SIZE Then



        ' Initialize the resize rules data (if needed)
        EnumChildWindows hWndForm , Cast(WNDENUMPROC , @FLY_ResizeRuleInitEnum) , 0
        If wParam <> SIZE_MINIMIZED Then
            EnumChildWindows hWndForm , Cast(WNDENUMPROC , @FLY_ResizeRuleEnum) , 0
        End If
    End If


    If wMsg = WM_ACTIVATE then
        If wParam = FALSE then
            gFLY_hDlgCurrent = 0
        Else
            gFLY_hDlgCurrent = hWndForm
        End If
    End If


    ' Get the ControlIndex (if any)
    Select Case wMsg
        Case WM_COMMAND , WM_HSCROLL , WM_VSCROLL
            If lParam <> 0 Then FLY_nResult = Cast(HANDLE , lParam)
        Case WM_NOTIFY
            FLY_pNotify = Cast(NMHDR Ptr , lParam)
            If FLY_pNotify -> hWndFrom <> 0 Then FLY_nResult = FLY_pNotify -> hWndFrom
    End Select
    If FLY_nResult Then FLY_control = GetProp(Cast(HWND , FLY_nResult) , "FLY_PTR")
    If FLY_control Then FLY_ControlIndex = FLY_control -> ControlIndex Else FLY_ControlIndex = - 1



    ' The following CASE calls notification events from Controls on the Form and
    ' for any of these messages that the user is handling themselves.
    Select Case wMsg

        Case WM_COMMAND

            If (LoWord(wParam) = IDC_FORM1_COMMAND1) And (HiWord(wParam) = BN_CLICKED) Then
                FLY_tempLong = FORM1_COMMAND1_BN_CLICKED(FLY_ControlIndex , hWndForm , Cast(HWND , lParam) , LoWord(wParam))
                If FLY_tempLong Then Function = FLY_tempLong : Exit Function
            End If

        Case WM_HSCROLL

        Case WM_VSCROLL

        Case WM_NOTIFY

    End Select

    ' Handle any custom messages if necessary.

    ' The following CASE processes the internal FireFly requirements
    Select Case wMsg

        Case WM_GETMINMAXINFO
            ' Do not process this message for MDI Child forms because it will interfere
            ' with the maximizing of the child form.
            If (GetWindowLongPtr(hWndForm , GWL_EXSTYLE) And WS_EX_MDICHILD) <> WS_EX_MDICHILD Then
                DefWindowProc hWndForm , wMsg , wParam , lParam
                Dim FLY_pMinMaxInfo   As MINMAXINFO Ptr
                FLY_pMinMaxInfo = Cast(MINMAXINFO Ptr , lParam)
                FLY_pMinMaxInfo -> ptMinTrackSize.x = FLY_pMinMaxInfo -> ptMinTrackSize.x
                FLY_pMinMaxInfo -> ptMinTrackSize.y = FLY_pMinMaxInfo -> ptMinTrackSize.y
                FLY_pMinMaxInfo -> ptMaxTrackSize.x = FLY_pMinMaxInfo -> ptMaxTrackSize.x
                FLY_pMinMaxInfo -> ptMaxTrackSize.y = FLY_pMinMaxInfo -> ptMaxTrackSize.y
                Function = 0 : Exit Function
            End If

        Case WM_SYSCOMMAND
            If (wParam And &HFFF0) = SC_CLOSE Then
                SendMessage hWndForm , WM_CLOSE , wParam , lParam
                Exit Function
            End If

        Case WM_SETFOCUS
            ' Set the focus back to the correct child control.
            If ff Then SetFocus ff -> CtrlFocus

        Case WM_SYSCOLORCHANGE , WM_THEMECHANGED
            ' Re-create any background brushes for the Controls on the Form.
            EnumChildWindows hWndForm , Cast(WNDENUMPROC , @FLY_EnumSysColorChangeProc) , 0

            ' Forward this message to any Common Controls because they do not automatically
            ' receive this message.

        Case WM_NOTIFY

        Case WM_CLOSE
            ' If we are dealing with a modal Form then it is important to re-enable the
            ' parent form and make it active prior to this form being destroyed.
            If ff Then
                If ff -> IsModal Then
                    ' Reset the focus back to the parent of the modal form.
                    ' Enable Mouse and keyboard input for the Parent Form
                    EnableWindow ff -> hWndParent , TRUE
                    SetActiveWindow ff -> hWndParent
                    ShowWindow hWndForm , SW_HIDE
                End If
            End If

        Case WM_DESTROY
            ' Delete any user defined Timer controls
            HWND_FORM1 = 0
            ' Delete the Fonts/Brushes used for this form/control and its associated property
            If ff Then
                If ff -> hBackBrush Then DeleteObject(Cast(HGDIOBJ , ff -> hBackBrush))
                ' Destroy the Accelerator Table if it was created
                If ff -> hAccel Then DestroyAcceleratorTable ff -> hAccel
                If ff -> hStatusBarTimer Then KillTimer ByVal 0 , ff -> hStatusBarTimer

                ' Re-claim the memory used by the type structure
                HeapFree GetProcessHeap() , 0 , ByVal Cast(PVOID , ff)

                ' Remove font from global array
                For x As Long = LBound(gFLY_FontHandles) To Ubound(gFLY_FontHandles)
                    If gFLY_FontHandles(x) = ff -> hFont Then
                        gFLY_FontHandles(x) = 0 : Exit For
                    End If
                Next

                ' Remove the property holding the Tag property string.
                FLY_nResult = GetProp(hWndForm , "FLY_TAGPROPERTY")
                If FLY_nResult Then
                    HeapFree GetProcessHeap() , 0 , ByVal Cast(PVOID , FLY_nResult)
                    RemoveProp hWndForm , "FLY_TAGPROPERTY"
                End If
                FLY_nResult = GetProp(hWndForm , "FLY_TAGPROPERTY2")
                If FLY_nResult Then
                    HeapFree GetProcessHeap() , 0 , ByVal Cast(PVOID , FLY_nResult)
                    RemoveProp hWndForm , "FLY_TAGPROPERTY2"
                End If

            End If
            RemoveProp hWndForm , "FLY_PTR"
            PostQuitMessage 0
            Function = 0 : Exit Function

        Case WM_CREATE
            ff = HeapAlloc(GetProcessHeap() , HEAP_ZERO_MEMORY , SizeOf(*ff))
            If ff Then
                SetProp hWndForm , "FLY_PTR" , ff' Store the pointer for later use
                ff -> IsForm = TRUE              ' This is a Form, not a Control.
            Else
                Function = - 1 : Exit Function   ' Return -1 to break the action
            End If

            ' Retrieve any user defined value that was passed via the Show function.
            FLY_UserData = Cast(Long Ptr , lParam)


            ' Set the Shared variable for this Form.
            HWND_FORM1 = hWndForm


            ' Set the Tag properties for the Form
            FF_Control_SetTag hWndForm , ""
            FF_Control_SetTag2 hWndForm , ""

            FORM1_CreateControls hWndForm

            ' Create the Keyboard accelerator if necessary


            ' The MDIClient (if applicatable) is being created, but not yet shown. Allow the user
            ' to process their commands prior to showing it.


            ' The form is being created, but not yet shown. Allow the user to process their
            ' commands prior to showing the Form's controls.
            FLY_tempLong = FORM1_WM_CREATE(hWndForm , *FLY_UserData)
            If FLY_tempLong Then Function = FLY_tempLong : Exit Function

            ' Create any user defined Timer controls


        Case WM_ERASEBKGND

        Case WM_TIMER

        Case WM_CTLCOLOREDIT , _
                    WM_CTLCOLORLISTBOX , _
                    WM_CTLCOLORSTATIC , _
                    WM_CTLCOLORBTN

            ' Message received from child control about to display.
            ' Set the Color properties here.
            ' lParam is the handle of the Control. wParam is the hDC.
            FLY_hDC = Cast(HDC , wParam)

            FLY_control = GetProp(Cast(HWND , lParam) , "FLY_PTR")
            If FLY_control = 0 Then
                FLY_control = GetProp(GetFocus() , "FLY_PTR")
                If FLY_control = 0 Then Exit Select
            End If

            If FLY_control -> ProcessCtlColor Then

                SetTextColor FLY_hDC , IIF(FLY_control -> IsForeSysColor , GetSysColor(FLY_control -> nForeColor) , FLY_control -> nForeColor)
                SetBkColor FLY_hDC , IIF(FLY_control -> IsBackSysColor , GetSysColor(FLY_control -> nBackColor) , FLY_control -> nBackColor)

                ' If this is a TextBox then we must use the OPAQUE STYLE, otherwise
                ' we will get a lot of screen garbage in the control when scrolling.
                GetClassName FLY_control -> hWndControl , FLY_zTempString , SizeOf(FLY_zTempString)

                ' Return the handle of the brush used to paint the background
                Function = Cast(LRESULT , FLY_control -> hBackBrush)

                If UCase(FLY_zTempString) = "EDIT" Then
                    SetBkMode FLY_hDC , OPAQUE

                    ' If the control is disabled then attempt to use the disabled color
                    If IsWindowEnabled(Cast(HWND , lParam)) = FALSE Then
                        SetBkColor FLY_hDC , GetSysColor(COLOR_BTNFACE)
                        Function = Cast(LRESULT , GetSysColorBrush(COLOR_BTNFACE))
                    End If

                Else
                    SetBkMode FLY_hDC , TRANSPARENT
                End If

                Exit Function

            End If

    End Select


    Function = DefWindowProc(hWndForm , wMsg , wParam , lParam)

End Function



'------------------------------------------------------------------------------
' FORM1_CODEPROCEDURE
'------------------------------------------------------------------------------
Private Function FORM1_CODEPROCEDURE(ByVal hWndControl As HWND , _
            ByVal wMsg As uInteger , _
            ByVal wParam As WPARAM , _
            ByVal lParam As LPARAM _
            ) As LRESULT

    ' All messages for every control on the the FORM1 form are processed
    ' in this function.

    Dim FLY_ProcAddress   As WNDPROC
    Dim FLY_hFont         As HFONT
    Dim FLY_hBackBrush    As Long
    Dim FLY_nResult       As HANDLE
    Dim FLY_tempLong      As Long
    Dim FLY_TopForm       As Long
    Dim FLY_ControlIndex  As Long
    Dim FLY_TCITEM        As TC_ITEM
    Dim uVersion          As OSVERSIONINFO
    Dim FLY_zTempString   As zString *MAX_PATH
    Dim FLY_Rect          As Rect

    Dim ff                As FLY_DATA Ptr
    Dim FLY_parent        As FLY_DATA Ptr
    If hWndControl Then ff = GetProp(hWndControl , "FLY_PTR")
    If ff Then FLY_ControlIndex = ff -> ControlIndex Else FLY_ControlIndex = - 1


    ' The following CASE processes the internal FireFly requirements prior
    ' to processing the user-defined events (These are handled in a
    ' separate CASE following this one).
    Select Case wMsg
        Case WM_DESTROY
            ' Unsublass the Control
            If ff Then

                ' Delete the brushes used for this Control
                If ff -> hBackBrush Then DeleteObject(Cast(HGDIOBJ , ff -> hBackBrush))

                FLY_ProcAddress = ff -> OldProc

                ' Re-claim the memory used by the type structure
                If ff Then HeapFree(GetProcessHeap() , 0 , ByVal Cast(PVOID , ff))
                RemoveProp hWndControl , "FLY_PTR"

                ' Remove the properties holding the Tag property strings.
                FLY_nResult = GetProp(hWndControl , "FLY_TAGPROPERTY")
                If FLY_nResult Then
                    HeapFree GetProcessHeap() , 0 , ByVal Cast(PVOID , FLY_nResult)
                    RemoveProp hWndControl , "FLY_TAGPROPERTY"
                End If
                FLY_nResult = GetProp(hWndControl , "FLY_TAGPROPERTY2")
                If FLY_nResult Then
                    HeapFree GetProcessHeap() , 0 , ByVal Cast(PVOID , FLY_nResult)
                    RemoveProp hWndControl , "FLY_TAGPROPERTY2"
                End If

                DeleteObject ff -> hFont

                ' Allow the WM_DESTROY message to be processed by the control.
                SetWindowLongPtr hWndControl , GWLP_WNDPROC , Cast(LONG_PTR , FLY_ProcAddress)
                CallWindowProc Cast(WNDPROC , FLY_ProcAddress) , hWndControl , wMsg , wParam , lParam

            End If
            Function = 0 : Exit Function

        Case WM_SETFOCUS
            ' If this is a TextBox, we check to see if we need to highlight the text.
            If ff Then
                If ff -> SelText Then
                    SendMessage hWndControl , EM_SETSEL , 0 , - 1
                Else
                    SendMessage hWndControl , EM_SETSEL , - 1 , 0
                End If
            End If

            ' Store the focus control in the parent form
            ' If this Form is a TabControl child Form then we need to store the CtrlFocus
            ' in the parent of this child Form.
            FLY_tempLong = GetWindowLongPtr(GetParent(hWndControl) , GWL_STYLE)
            If (FLY_tempLong And WS_CHILD) = WS_CHILD Then
                ' must be a TabControl child dialog
                FLY_parent = GetProp(GetParent(GetParent(hWndControl)) , "FLY_PTR")
            Else
                FLY_parent = GetProp(GetParent(hWndControl) , "FLY_PTR")
            End If
            If FLY_parent Then FLY_parent -> CtrlFocus = hWndControl


        Case WM_SIZE , WM_MOVE
            ' Handle any TabControl child pages that are autosize. If the TabControl changes
            ' size then it is caught here. We only need to deal with TabControlChildAutoSize
            ' forms because other forms stay at the fixed size.

    End Select

    ' The following case calls each user defined function.
    Select Case wMsg
        Case 0
    End Select

    ' Handle any custom messages if necessary.


    ' This control is subclassed, therefore we must send all unprocessed messages
    ' to the original window procedure.
    Function = CallWindowProc(ff -> OldProc , hWndControl , wMsg , wParam , lParam)

End Function





'--------------------------------------------------------------------------------
Private Function FORM1_WM_CREATE( _
            ByVal hWndForm as HWnd , _                   ' handle of Form
            ByVal UserData as Long _                     ' optional user defined value
            ) as Long

    Function = 0                                 ' change according to your needs
End Function


' FINISH INCLUDE ¤ 003 _ D:\freebasic\FireFly_FB_378\Projects\Pass_Generator\Release\CODEGEN_PASS_GENERATOR_FORM1_FORM.inc
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤




'#PARSESTART# 'do not delete this line


Private Function FF_WINMAIN(ByVal hInstance As HINSTANCE , _
            ByVal hPrevInstance as HINSTANCE , _
            ByRef lpCmdLine as String , _
            ByVal iCmdShow as Long) as Long


    ' If this function returns TRUE (non-zero) then the actual WinMain will exit
    ' thus ending the program. You can do program initialization in this function.

    Function = False                             'return TRUE if you want the program to end.

End Function

'[END_WINMAIN]

Private Function FF_PUMPHOOK(ByRef uMsg As Msg) As Long


    ' If this function returns FALSE (zero) then the processing of the regular
    ' FireFly message pump will continue. Returning TRUE (non-zero) will bypass
    ' the regular message pump.

    Function = FALSE                             ' return TRUE if you need to bypass the FireFly message pump

End Function

'[END_PUMPHOOK]


Private Sub FLY_InitializeVariables()
    ' All FireFly variables relating to Forms and Controls are initialized here. This
    ' includes any Control Arrays that have been defined and control identifiers
    ' per the listing in the Declares include file.

    IDC_FORM1_CHECK7 = 1000
    IDC_FORM1_CHECK6 = 1001
    IDC_FORM1_CHECK5 = 1002
    IDC_FORM1_TEXT3 = 1003
    IDC_FORM1_LABEL1 = 1004
    IDC_FORM1_TEXT2 = 1005
    IDC_FORM1_PICTURE1 = 1006
    IDC_FORM1_CHECK3 = 1007
    IDC_FORM1_CHECK4 = 1008
    IDC_FORM1_COMMAND1 = 1009
    IDC_FORM1_TEXT1 = 1010
    IDC_FORM1_CHECK1 = 1011
    IDC_FORM1_CHECK2 = 1012
	IDC_FORM1_CHECK8 = 1013
	IDC_FORM1_CHECK9 = 1014
	IDC_FORM1_TEXT4 = 1015
	IDC_FORM1_LABEL1 = 1016
End Sub


Private Sub FLY_SetAppVariables()
    ' All FireFly App variables are initialized here. This Type variable provides
    ' easy access to many commonly used FireFly settings.

    Dim zTemp             As zString *MAX_PATH
    Dim x                 As Long

    App.CompanyName = ""
    App.FileDescription = ""
    App.ProductName = ""
    App.LegalCopyright = ""
    App.LegalTrademarks = ""
    App.Comments = ""

    App.ProductMajor = 1
    App.ProductMinor = 0
    App.ProductRevision = 0
    App.ProductBuild = 0

    App.FileMajor = 1
    App.FileMinor = 0
    App.FileRevision = 0
    App.FileBuild = 0

    ' App.hInstance is set in WinMain/LibMain

    ' Retrieve program full path and EXE/DLL name
    GetModuleFileName App.hInstance , zTemp , MAX_PATH

    x = InStr(- 1 , zTemp , Any ":/\")

    If x Then
        App.Path = Left(zTemp , x)
        App.EXEname = Mid(zTemp , x + 1)
    Else
        App.Path = ""
        App.EXEname = zTemp
    End If


    ' The following two arrays are used to allow FireFly to reuse font
    ' handles that are common to one or more Controls. This saves us
    ' from consuming a GDI resource for every created control font.
    ' ReDim gFLY_FontNames(0) As String
    ReDim gFLY_FontHandles(0) As HFONT

End Sub


Private Function FLY_AdjustWindowRect(ByVal hWndForm As HWND , _
            ByVal cxClient As Long , _
            ByVal cyClient As Long _
            ) As Long

    Dim dwStyle           As Long
    Dim hMenu             As HMENU
    Dim rc                As Rect

    If (cxClient <> 0) And (cyClient <> 0) Then
        dwStyle = GetWindowLongPtr(hWndForm , GWL_STYLE)
        rc.Left = 0 : rc.Top = 0 : rc.Right = cxClient : rc.Bottom = cyClient
        hMenu = GetMenu(hWndForm)
        AdjustWindowRectEx VarPtr(rc) , dwStyle ,(hMenu <> 0) , GetWindowLongPtr(hWndForm , GWL_EXSTYLE)
        If (dwStyle and WS_HSCROLL) = WS_HSCROLL Then rc.Bottom = rc.Bottom + GetSystemMetrics(SM_CYHSCROLL)
        If (dwStyle and WS_VSCROLL) = WS_VSCROLL Then rc.Right = rc.Right + GetSystemMetrics(SM_CXVSCROLL)
        SetWindowPos hWndForm , 0 , 0 , 0 , _
                rc.Right - rc.Left , rc.Bottom - rc.Top , _
                SWP_NOZORDER Or SWP_NOMOVE Or SWP_NOACTIVATE
    End If

    Function = 0

End Function



' Enum all child windows of the Form (basically, all Controls) in order to handle
' the WM_SYSCOLORCHANGE message. This will delete existing brushes and recreate
' them using the new system colors.
Private Function FLY_EnumSysColorChangeProc(ByVal hWnd As HWND , ByVal lParam As LPARAM) As Long
    Dim ff                As FLY_DATA Ptr

    ff = GetProp(hWnd , "FLY_PTR")
    If ff Then
        If ff -> ProcessCtlColor Then
            ' Create the new colors/brushes if we are using System colors
            If ff -> IsForeSysColor Then ff -> nForeColor = GetSysColor(ff -> nForeColor)
            If ff -> IsBackSysColor Then
                If ff -> hBackBrush Then DeleteObject(Cast(HGDIOBJ , ff -> hBackBrush))
                ff -> hBackBrush = GetSysColorBrush(ff -> nBackColor)
            End If
        End If
    End If

    Function = TRUE
End Function

Private Function FLY_SetControlData(ByVal hWndControl As HWND , _
            ByVal AllowSubclass As Long , _
            ByVal AllowSetFont As Long , _
            ByRef sFontString As String , _
            ByVal nControlIndex As Long , _
            ByVal nProcessColor As Long , _
            ByVal IsForeSysColor As Long , _
            ByVal IsBackSysColor As Long , _
            ByVal nForeColor As Long , _
            ByVal nBackColor As Long , _
            ByVal nTransparentBrush As HBRUSH , _
            ByVal CodeProcedure As WNDPROC , _
            Byref sResizeRules As String , _
            ByVal nFontUpgrade As Long _
            ) As FLY_DATA ptr

    Dim zClassName        As zString *50
    Dim sFont             As String
    Dim ff                As FLY_DATA Ptr

    ff = HeapAlloc(GetProcessHeap() , HEAP_ZERO_MEMORY , SizeOf(*ff))
    If ff Then
        ' Store the pointer for later use
        SetProp hWndControl , "FLY_PTR" , ff

        ' Subclass the control
        If AllowSubclass = TRUE Then
            ff -> OldProc = Cast(WNDPROC , SetWindowLongPtr(hWndControl , GWLP_WNDPROC , Cast(LONG_PTR , CodeProcedure)))
        End If

        ' Set the Font for this Form/Control
        If AllowSetFont = TRUE Then

            ' FireFly handles two types of font strings. The first is the traditional
            ' logfont string and the second is the simplified PB version. FireFly is moving
            ' towards the simplified version but needs to handle the logfont version
            ' for backwards compatibility.

            ' If the FontUpgrade property is active then we need to replace the FontName
            ' based on the current operating system.
            If nFontUpgrade <> 0 Then
                IF AfxGetWindowsVersion >= 6 THEN' Vista/Windows 7+
                    sFontString = "Segoe UI," & FF_Remain_Internal(sFontString , ",")
                Else
                    sFontString = "Tahoma," & FF_Remain_Internal(sFontString , ",")
                End If
            End If

            ff -> hFont = FF_MakeFontEx_Internal(sFontString)

            If ff -> hFont Then
                SendMessage hWndControl , WM_SETFONT , Cast(WPARAM , ff -> hFont) , True
                If UCase(zClassName) = "SYSDATETIMEPICK32" Then
                    SendMessage hWndControl , DTM_SETMCFONT , Cast(WPARAM , ff -> hFont) , True
                End If
            End If

        End If

        ff -> ControlIndex = nControlIndex
        ff -> IsForm = FALSE
        ff -> hwndControl = hWndControl
        ff -> zConstraints = sResizeRules

        ' Flag to process the WM_CTLCOLOR??? messages.
        ff -> ProcessCtlColor = nProcessColor
        If ff -> ProcessCtlColor Then
            ff -> IsForeSysColor = IsForeSysColor
            ff -> IsBackSysColor = IsBackSysColor
            ff -> nForeColor = nForeColor
            ff -> nBackColor = nBackColor
            ' Create a Brush for painting the background of this Control.
            If nTransparentBrush <> Cast(HBRUSH , - 1) Then
                ff -> hBackBrush = nTransparentBrush
            Else
                ff -> hBackBrush = IIF(ff -> IsBackSysColor , GetSysColorBrush(ff -> nBackColor) , CreateSolidBrush(ff -> nBackColor))
            End If
        End If

    End If

    Function = ff                                ' return the pointer to the data structure

End Function



Private Function FLY_DoMessagePump(ByVal ShowModalFlag As Long , _
            ByVal hWndForm As HWND , _
            ByVal hWndParent As HWND , _
            ByVal nFormShowState As Long , _
            ByVal IsMDIForm As Long _
            ) As HWND

    Dim zTempString       As zString *MAX_PATH
    Dim hWndActive        As HWND
    Dim msg               As MSG
    Dim ff                As FLY_DATA Ptr

    ff = GetProp(hWndForm , "FLY_PTR")
    If ff = 0 Then Exit Function


    ' If this is an MDI child form, then it can not be displayed as modal.
    If (GetWindowLongPtr(hWndForm , GWL_EXSTYLE) And WS_EX_MDICHILD) = WS_EX_MDICHILD Then ShowModalFlag = FALSE


    If ShowModalFlag = TRUE Then

        ' Determine the top level window of the active control
        WHILE(GetWindowLongPtr(hWndParent , GWL_STYLE) AND WS_CHILD) <> 0
            hWndParent = GetParent(hWndParent)
            IF (GetWindowLongPtr(hWndParent , GWL_EXSTYLE) AND WS_EX_MDICHILD) <> 0 THEN EXIT WHILE
        WEND

        ' Set property to show that this is a Modal Form. The hWndParent is
        ' stored in a property because that window needs to be re-enabled
        ' prior to modal form being destroyed and exiting the modal message loop.
        ff -> IsModal = TRUE
        ff -> hWndParent = hWndParent

        ' Disable Mouse and keyboard input for the Parent Form
        If hWndParent <> 0 Then EnableWindow(hWndParent , False)

        ShowWindow hWndForm , nFormShowState
        UpdateWindow hWndForm

        ' Main message loop:
        Do While GetMessage(@Msg , 0 , 0 , 0)
            ' Exit the modal message loop if the Form was destroyed (important).
            If IsWindow(hWndForm) = FALSE Then Exit Do
            If msg.message = WM_QUIT Then Exit Do

            If FF_PUMPHOOK(Msg) = 0 Then
                If IsMDIForm = TRUE Then
                    If TranslateMDISysAccel(hWndForm , @Msg) <> 0 Then Continue Do
                End If

                If TranslateAccelerator(hWndForm , ff -> hAccel , @Msg) = 0 Then

                    hWndActive = FLY_GetActiveHandle(gFLY_hDlgCurrent)

                    ' Handle the strange situation where pressing ESCAPE in a multiline
                    ' textbox causes the application to receive a WM_CLOSE that could
                    ' cause an application to terminate. also allow the TAB key to move
                    ' in and out of a multiline textbox.
                    GetClassName GetFocus , zTempString , SizeOf(zTempString)
                    Select Case Ucase(zTempString)
                        Case "EDIT"
                            IF (GetWindowLongPtr(GetFocus , GWL_STYLE) AND ES_MULTILINE) = ES_MULTILINE THEN
                                If (msg.message = WM_KEYDOWN) And (Msg.wParam = VK_ESCAPE) Then
                                    msg.message = WM_COMMAND
                                    msg.wParam = MakeLong(IDCANCEL , 0)
                                    msg.lParam = 0
                                ElseIf (Msg.message = WM_CHAR) And (Msg.wParam = 9) Then
                                    ' allow the Tab key to tab out of a multiline textbox
                                    If (GetAsyncKeyState(VK_SHIFT) And &H8000) = 0 Then
                                        SetFocus GetNextDlgTabItem(GetParent(Msg.hWnd) , Msg.hWnd , FALSE)
                                    Else
                                        SetFocus GetNextDlgTabItem(GetParent(Msg.hWnd) , Msg.hWnd , TRUE)
                                    End If
                                    msg.message = WM_NULL
                                End If
                            End If
                    End Select

                    If (hWndActive = 0) Or (IsDialogMessage(hWndActive , @Msg) = 0) Then
                        TranslateMessage @Msg
                        DispatchMessage @Msg
                    End If

                End If
            End If
        Loop

        Function = Cast(HWND , Cast(LONG_PTR , App.ReturnValue))
        App.ReturnValue = 0

    Else
        ShowWindow hWndForm , nFormShowState
        Function = hWndForm
    End If

End Function


Private Function FLY_GetActiveHandle(ByVal hWnd As HWND) as HWND

    Static szClassName    As zString *100

    ' Determine the top level window of the active control
    Do While(GetWindowLongPtr(hWnd , GWL_STYLE) AND WS_CHILD) <> 0
        IF (GetWindowLongPtr(hWnd , GWL_EXSTYLE) AND WS_EX_MDICHILD) <> 0 THEN EXIT DO
        GetClassName hWnd , szClassName , SizeOf(szClassName)
        If Left(szClassName , 5) = "FORM_" Then EXIT DO
        hWnd = GetParent(hWnd)
    loop

    FUNCTION = hWnd

End Function


Private Function FLY_ResizeRuleInitEnum(ByVal hWnd As HWND , ByVal lParam As LPARAM) As Long

    Dim ff                As FLY_DATA Ptr

    ff = GetProp(hWnd , "FLY_PTR")
    If ff = 0 Then Function = TRUE : Exit Function

    If ff -> fResizeInit = TRUE Then Function = TRUE : Exit Function

    ' Store form width and height for future reference
    GetClientRect GetParent(hWnd) , Varptr(ff -> rcForm)

    ' Get the control's rectangle and convert it to client coordinates
    GetWindowRect hWnd , Varptr(ff -> rcCtrl)
    MapWindowPoints 0 , GetParent(hWnd) , Cast(LPPOINT , Varptr(ff -> rcCtrl)) , 2

    ff -> fResizeInit = TRUE

    Function = TRUE

End Function



Private Function FLY_ResizeRuleEnum(ByVal hWnd As HWND , ByVal lParam As LPARAM) As Long

    Dim buf(1 To 4)       As FLY_RESCALEDATA
    Dim i                 As Long
    Dim Constr            As String
    Dim Constraints       As String
    Dim ConstrRef         As String
    Dim RefPntPercentage  As Single
    Dim ControlName       As String
    Dim rc                As Rect
    Dim ff                As FLY_DATA Ptr

    ff = GetProp(hWnd , "FLY_PTR")

    If ff = 0 Then Function = TRUE : Exit Function
    If ff -> fResizeInit = FALSE Then Function = TRUE : Exit Function

    Constraints = UCase(ff -> zConstraints)
    if Len(Constraints) < 7 Then Function = TRUE : Exit Function

    ' Control was already initialized. Get its initial coordinates to be used for reference.
    ' Store the controls initial coordinates in the calculation buffer.
    ' One entry for every side of the control.
    Buf(1).P_Init = ff -> rcCtrl.Left            ' Left x-coordinate of the control
    Buf(2).P_Init = ff -> rcCtrl.Top             ' Top y-coordinate of the control
    Buf(3).P_Init = ff -> rcCtrl.Right           ' Right x-coordinate of the control
    Buf(4).P_Init = ff -> rcCtrl.Bottom          ' Bottom y-coordinate of the control

    ' Store the form's width and height in the calculation buffer.
    ' One entry for every side of the form
    Buf(1).FormInit = ff -> rcForm.Right         ' For Left x-coordinate of the control
    Buf(2).FormInit = ff -> rcForm.Bottom        ' For Top y-coordinate of the control
    Buf(3).FormInit = ff -> rcForm.Right         ' For Right x-coordinate of the control
    Buf(4).FormInit = ff -> rcForm.Bottom        ' For Bottom y-coordinate of the control


    ' -------------------------------------------------------------------------------
    ' (2) Get current dimensions of form and store the coordinates in a
    ' buffer to be calculated later on
    ' -------------------------------------------------------------------------------
    GetClientRect GetParent(hWnd) , Varptr(rc)
    Buf(1).FormCur = rc.Right                    ' Left x-coordinate
    Buf(2).FormCur = rc.Bottom                   ' Top y-coordinate
    Buf(3).FormCur = rc.Right                    ' Right x-coordinate
    Buf(4).FormCur = rc.Bottom                   ' Bottom y-coordinate


    ' -------------------------------------------------------------------------------
    ' (3) Calculate new control coordinates (for every one of the 4 control sides)
    ' -------------------------------------------------------------------------------
    For i = 1 To 4
        Constr = FF_Parse(Constraints , "," , i)

        Select Case Left(Constr , 1)
            Case "S"                             ' Scaled coordinate
                Buf(i).P_New = Buf(i).P_Init * (Buf(i).FormCur / Buf(i).FormInit)

            Case "F"                             ' Fixed coordinate
                ConstrRef = Mid(Constr , 2)      ' Remove the "F"
                Select Case ConstrRef
                    Case "L" : RefPntPercentage = 0 ' Left side
                    Case "R" : RefPntPercentage = 100 ' Right side
                    Case "T" : RefPntPercentage = 0 ' Top
                    Case "B" : RefPntPercentage = 100 ' Bottom
                    Case Else                    ' Must have been a number
                        RefPntPercentage = Val(ConstrRef)
                End Select

                ' Calc the new position
                Buf(i).P_New = Buf(i).P_Init - _
                        (Buf(i).FormInit * (RefPntPercentage / 100)) + _
                        (Buf(i).FormCur * (RefPntPercentage / 100))

            Case Else
                'MsgBox "Unknown constraint parameter:" & Left$(Constr, 1) & " in " & Constraints
        End Select

    Next


    ' -------------------------------------------------------------------------------
    ' (4) Redraw the control
    ' -------------------------------------------------------------------------------
    SetWindowPos hWnd , 0 , _
            Buf(1).P_New , Buf(2).P_New , _
            Buf(3).P_New - Buf(1).P_New , Buf(4).P_New - Buf(2).P_New , _
            SWP_NOZORDER

    Function = TRUE                              ' Continue enumeration of children

End Function



Private Function WinMain(ByVal hInstance As HINSTANCE , _
            ByVal hPrevInstance As HINSTANCE , _
            ByRef lpCmdLine As String , _
            ByVal iCmdShow As Long _
            ) As Long

    ' Initialize the Common Controls. This is necessary on WinXP platforms in order
    ' to ensure that the XP Theme styles work.
    Dim uCC               As INITCOMMONCONTROLSEX
    Dim hLib              As HINSTANCE

    uCC.dwSize = SizeOf(uCC)
    uCC.dwICC = ICC_NATIVEFNTCTL_CLASS Or ICC_COOL_CLASSES Or ICC_BAR_CLASSES Or _
            ICC_TAB_CLASSES Or ICC_USEREX_CLASSES Or ICC_WIN95_CLASSES Or _
            ICC_STANDARD_CLASSES Or ICC_ANIMATE_CLASS Or ICC_DATE_CLASSES Or _
            ICC_HOTKEY_CLASS Or ICC_INTERNET_CLASSES Or ICC_LISTVIEW_CLASSES Or _
            ICC_PAGESCROLLER_CLASS Or ICC_PROGRESS_CLASS Or ICC_TREEVIEW_CLASSES Or _
            ICC_UPDOWN_CLASS

    InitCommonControlsEx VarPtr(uCC)

    ' Define variable used to uniquely identify this application. This is used
    ' by the FireFly function FF_PrevInstance.
    gFLY_AppID = "FORM_Pass_Generator_Form1_CLASS"


    ' Define Control Arrays if necessary
    FLY_InitializeVariables


    ' Set the values for the Shared App variable
    App.hInstance = hInstance
    FLY_SetAppVariables

    ' Call the Function FLY_WinMain(). If that function returns True
    ' then cease to continue execution of this program.
    If FF_WINMAIN(hInstance , hPrevInstance , lpCmdLine , iCmdShow) Then
        Function = App.ReturnValue
        Exit Function
    End If

    ' Create the Startup Form.
    Form1_Show 0 , TRUE

    Function = App.ReturnValue

End Function

private function nodup(byref ss0 as string, byref ncount as long ) as string

	dim as ubyte ptr uptr = strptr(ss0)
	dim as string sret = chr(uptr[0])
	dim as long x
	dim as long y = 2
	dim as long z
	dim as long w = 1
	do while w < ncount
		dim char as ubyte = uptr[y - 1]
		if  char = 0  then
			ncount = w
			return sret
		end if
		z = 0
		for x = 1 to y - 1
			if char = asc(sret, x) then
				z = 1
				exit for
			end if
		NEXT
		if z = 0 THEN
			sret &= chr(char)
			w += 1
        END IF
		y += 1
	LOOP
	return sret
end function

Private Function FF_TextBox_SetText (ByVal hWndControl as HWnd, _
                             ByRef TheText as String) as Long

    ' Do a check to ensure that this is actually a window handle
    If IsWindow(hWndControl) Then
       ' Set the control's text
         SetWindowText hWndControl, TheText
    End If
    Return 0
End Function


Private Function RemoveAnyChar(byref instring as string , byref remchars as string) as string

    Dim As String outstring = ""
	dim as long iflag
	For i As long = 0 To Len(instring) - 1
		iflag = 0
		for j as long = 0 to Len(remchars) - 1
            If instring[i] = Remchars[j] then
				iflag = 1
				exit for
			end if
        Next
		if iflag = 0 THEN
			outstring &= chr(instring[i])
        END IF
    NEXT
    Return outstring

End Function

'--------------------------------------------------------------------------------
Private Function FORM1_COMMAND1_BN_CLICKED( _
            ByVal ControlIndex as Long , _               ' index in Control Array
            ByVal hWndForm as HWnd , _                   ' handle of Form
            ByVal hWndControl as HWnd , _                ' handle of Control
            ByVal idButtonControl as Long _              ' identifier of button
            ) as Long
    Function = 0                                 ' change according to your needs
	dim ForeColor as Long
	dim BackColor as Long
	dim as long ncount = val(FF_Control_GetText(HWND_FORM1_TEXT2))
	FF_TextBox_SetText(HWND_FORM1_TEXT1, "")
	FF_Control_GetColor(HWND_FORM1_LABEL1 , ForeColor, BackColor)

	FF_Control_SetColor(HWND_FORM1_LABEL2 , ForeColor, BackColor)
	FF_TextBox_SetText(HWND_FORM1_LABEL2, "Select your choices et click 'Generate' ")

    dim as long CHECK1 = FF_Control_GetCheck(HWND_FORM1_CHECK1)
    dim as long CHECK2 = FF_Control_GetCheck(HWND_FORM1_CHECK2)
    dim as long CHECK3 = FF_Control_GetCheck(HWND_FORM1_CHECK3)
    dim as long CHECK4 = FF_Control_GetCheck(HWND_FORM1_CHECK4)
    dim as long CHECK5 = FF_Control_GetCheck(HWND_FORM1_CHECK5)
    dim as long CHECK6 = FF_Control_GetCheck(HWND_FORM1_CHECK6)
    dim as long CHECK7 = FF_Control_GetCheck(HWND_FORM1_CHECK7)
	dim as long CHECK8 = FF_Control_GetCheck(HWND_FORM1_CHECK8)
	dim as long CHECK9 = FF_Control_GetCheck(HWND_FORM1_CHECK9)
	dim as long CHECK81 = FF_Control_GetCheck(HWND_FORM1_CHECK81)

    dim as string s1 = "012345678901234567890123456789"
    dim as string s2 = ":;.,?!:;.,?!:;.,?!:;.,?!"
    dim as string s3 = "+-_*/<=>({[]})+-_*/<=>({[]})+-_*/<=>({[]})+-_*/<=>({[]})"
    dim as string s4 = "#@|%&$\^#@|%&$\^#@|%&$\^#@|%&$\^"
    dim as string s5 = "ABCDEFGHIJKLMNOPQRSTUVWXYZABCDEFGHIJKLMNOPQRSTUVWXYZABCDEFGHIJKLMNOPQRSTUVWXYZ"
    dim as string s6 = "abcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyz"
    dim as string s7 = FF_Control_GetText(HWND_FORM1_TEXT3)
	dim as string s9 = FF_Control_GetText(HWND_FORM1_TEXT4)
    dim as string s0 = ""

    if CHECK1 = true THEN s0 &= s1
    if CHECK2 = true THEN s0 &= s2
    if CHECK3 = true THEN s0 &= s3
    if CHECK4 = true THEN s0 &= s4
    if CHECK5 = true THEN s0 &= s5
    if CHECK6 = true THEN s0 &= s6
    if CHECK7 = true THEN s0 &= s7 & s7 & s7
	if CHECK9 = true and s9 <> "" THEN s0 = RemoveAnyChar(s0 , s9)

	if ncount < 1 or  s0 = "" THEN
		messagebox(hWndForm, !"\n\n Select different choices and length > 0 \n\n\t Please retry." , "Error : invalid choices", MB_SYSTEMMODAL + MB_ICONERROR )
		exit function
    END IF

    'print s0
	s0 = s0 & s0 & s0
	dim as Ulong useed = get_Useed()
	dim as string ss0 = MyRand32_RandomString(len(s0), useed , s0)
	ss0 = MyRand32_RandomString(len(ss0), useed + 11, ss0)
	if CHECK81 = true THEN
		ss0 = MyRand32_RandomString(ncount * 2, useed + 13, ss0, 1)
	else
		ss0 = MyRand32_RandomString(ncount * 2, useed + 13, ss0)
	end if
	'print ss0
	dim as long ucount = ncount
	dim as string ss3
	if CHECK8 = true THEN
		ss3 = nodup(ss0, ucount)
	else
		ss3 = left(ss0, ncount)
		ucount = len(ss3)
    END IF

	FF_TextBox_SetText(HWND_FORM1_TEXT1, ss3)
    if ucount < ncount THEN
		FF_Control_SetColor(HWND_FORM1_LABEL2 , 255, BackColor )
		FF_TextBox_SetText(HWND_FORM1_LABEL2, "Too short.  Length = " & ucount & "  only!")
	else
		FF_Control_SetColor(HWND_FORM1_LABEL2 ,bgr(0,200,0), BackColor )
		FF_TextBox_SetText(HWND_FORM1_LABEL2, "Correct Password.")
    END IF
End Function

End WinMain(GetModuleHandle(null) , null , Command() , SW_NORMAL)




