#cs ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Name..................: IBM_ACS_API_Core_UDF
    Description...........: UDF for 'IBM i Access Client Solutions' API
    Dependencies..........: IBM i ACS 32-bit application, corresponding DLLs
    Documentation.........: https://www.ibm.com/support/knowledgecenter/en/ssw_ibm_i_71/rzaik/rzaikemulator.htm
                            https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/emulator_programming08.htm

    Author................: exorcistas@github.com
    Modified..............: 2021-03-05
    Version...............: v0.8.0.1
#ce ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#AutoIt3Wrapper_UseX64=n    ;-- required to run in 32bit environment
#include-once
#include <Array.au3>
#include <String.au3>

Global Const $IBMACS_PCSHLL = "PCSHLL32.dll"
Global Const $IBMACS_EHLAPI = "EHLAPI32.dll"
Global $IBMACS_DLL_FILE = $IBMACS_PCSHLL
Global $IBMACS_DEBUG_ENABLE = False

#Region FUNCTIONS_LIST
#cs	===================================================================================================================================
%% SUPPORT %%
    __ReplaceChars2ASCII

%% CORE %%
    _IBMACS_OpenDll()
    _IBMACS_CloseDll()
    _IBMACS_DLLCall($_iFunction, $_sDataString = "", $_iDataLength = 0, $_iReturnCode = "")
    _IBMACS_ConnectPresentationSpace($_sSessionName = "A", $_bConnect = True)
    _IBMACS_SendKey($_sKey)
    _IBMACS_CopyFieldToString($_iRow, $_iCol, $iLength)
    _IBMACS_CopyStringToField($_sString, $_iRow, $_iCol)
    _IBMACS_SetCursor($_iRow, $_iCol)
    _IBMACS_Wait()
    _IBMACS_CopyPresentationScreen()
    _IBMACS_ResetSystem()
    _IBMACS_SetSessionParameters($_sParams)
#ce	===================================================================================================================================
#EndRegion FUNCTIONS_LIST

#Region SUPPORT_FUNCTIONS
    Func __ReplaceChars2ASCII($_sText)
        #region swedish
            $_sText = StringReplace($_sText, "ä", Chr(132), 0, 1)
            $_sText = StringReplace($_sText, "Ä", Chr(142), 0, 1)
            $_sText = StringReplace($_sText, "å", Chr(134), 0, 1)
            $_sText = StringReplace($_sText, "Å", Chr(143), 0, 1)
            $_sText = StringReplace($_sText, "ö", Chr(148), 0, 1)
            $_sText = StringReplace($_sText, "Ö", Chr(153), 0, 1)
        #endregion swedish

        #region other
            $_sText = StringReplace($_sText, "ç", Chr(135), 0, 1)
            $_sText = StringReplace($_sText, "Ç", Chr(128), 0, 1)
            $_sText = StringReplace($_sText, "ü", Chr(129), 0, 1)
            $_sText = StringReplace($_sText, "Ü", Chr(154), 0, 1)
            $_sText = StringReplace($_sText, "é", Chr(130), 0, 1)
            $_sText = StringReplace($_sText, "É", Chr(144), 0, 1)
            $_sText = StringReplace($_sText, "â", Chr(131), 0, 1)
            $_sText = StringReplace($_sText, "Â", Chr(182), 0, 1)
            $_sText = StringReplace($_sText, "à", Chr(133), 0, 1)
            $_sText = StringReplace($_sText, "À", Chr(183), 0, 1)
            $_sText = StringReplace($_sText, "ê", Chr(136), 0, 1)
            $_sText = StringReplace($_sText, "Ê", Chr(210), 0, 1)
            $_sText = StringReplace($_sText, "ë", Chr(137), 0, 1)
            $_sText = StringReplace($_sText, "Ë", Chr(211), 0, 1)
            $_sText = StringReplace($_sText, "è", Chr(138), 0, 1)
            $_sText = StringReplace($_sText, "È", Chr(212), 0, 1)
            $_sText = StringReplace($_sText, "ï", Chr(139), 0, 1)
            $_sText = StringReplace($_sText, "Ï", Chr(216), 0, 1)
            $_sText = StringReplace($_sText, "î", Chr(140), 0, 1)
            $_sText = StringReplace($_sText, "Î", Chr(215), 0, 1)
            $_sText = StringReplace($_sText, "ì", Chr(141), 0, 1)
            $_sText = StringReplace($_sText, "Ì", Chr(222), 0, 1)
            $_sText = StringReplace($_sText, "ì", Chr(141), 0, 1)
            $_sText = StringReplace($_sText, "Ì", Chr(222), 0, 1)
            $_sText = StringReplace($_sText, "æ", Chr(145), 0, 1)
            $_sText = StringReplace($_sText, "Æ", Chr(146), 0, 1)
            $_sText = StringReplace($_sText, "ô", Chr(147), 0, 1)
            $_sText = StringReplace($_sText, "Ô", Chr(226), 0, 1)
            $_sText = StringReplace($_sText, "ò", Chr(149), 0, 1)
            $_sText = StringReplace($_sText, "Ò", Chr(227), 0, 1)
            $_sText = StringReplace($_sText, "û", Chr(150), 0, 1)
            $_sText = StringReplace($_sText, "Û", Chr(234), 0, 1)
            $_sText = StringReplace($_sText, "ù", Chr(151), 0, 1)
            $_sText = StringReplace($_sText, "Ù", Chr(235), 0, 1)
            $_sText = StringReplace($_sText, "ÿ", Chr(152), 0, 1)
            $_sText = StringReplace($_sText, "ý", Chr(236), 0, 1)
            $_sText = StringReplace($_sText, "Ý", Chr(237), 0, 1)
            $_sText = StringReplace($_sText, "á", Chr(160), 0, 1)
            $_sText = StringReplace($_sText, "Á", Chr(181), 0, 1)
            $_sText = StringReplace($_sText, "í", Chr(161), 0, 1)
            $_sText = StringReplace($_sText, "Í", Chr(214), 0, 1)
            $_sText = StringReplace($_sText, "ó", Chr(162), 0, 1)
            $_sText = StringReplace($_sText, "Ó", Chr(224), 0, 1)
            $_sText = StringReplace($_sText, "ú", Chr(163), 0, 1)
            $_sText = StringReplace($_sText, "Ú", Chr(233), 0, 1)
            $_sText = StringReplace($_sText, "ñ", Chr(164), 0, 1)
            $_sText = StringReplace($_sText, "Ñ", Chr(165), 0, 1)
            $_sText = StringReplace($_sText, "õ", Chr(228), 0, 1)
            $_sText = StringReplace($_sText, "Õ", Chr(229), 0, 1)
        #endregion other

        Return $_sText
    EndFunc
#EndRegion SUPPORT_FUNCTIONS

#Region API_CORE_FUNCTIONS

    #cs #FUNCTION# ==========================================================================================================================
        Name...............: _IBMACS_OpenDll()
        Description .......: Open DLL connection
        Syntax.............: see DLLOpen function documentation
        Parameters ........: -
        Return values .....: Success - dll handle;
                             Failure - False + @error + @extended

        Author ............: exorcistas@github.com
        Modified...........: 2020-01-07
    #ce =====================================================================================================================================
    Func _IBMACS_OpenDll()
        $_dll = DLLOpen($IBMACS_DLL_FILE)
            If $_dll = -1 Then Return SetError(1, 1, False)

        Return $_dll
    EndFunc

    #cs #FUNCTION# ==========================================================================================================================
        Name...............: _IBMACS_CloseDll()
        Description .......: Close DLL connection
        Syntax.............: see DllClose function documentation
        Parameters ........: -
        Return values .....: -
        Author ............: exorcistas@github.com
        Modified...........: 2020-01-07
    #ce =====================================================================================================================================
    Func _IBMACS_CloseDll()
        DllClose($IBMACS_DLL_FILE)
    EndFunc

    #cs #FUNCTION# ==========================================================================================================================
        Name...............: _IBMACS_DLLCall($_iFunction, $_sDataString = "", $_iDataLength = 0, $_iReturnCode = "")
        Description .......: Generic function to call other functions from the dll
        Syntax.............: see DllCall function documentation
        Parameters ........: $_iFunction - function number; $_sDataString - string to send to function; $_iDataLength - expected data length; $_iReturnCode - return code
        Return values .....: Success - Array of response; 
                             Failure - False + @error + @extended

        Author ............: exorcistas@github.com
        Modified...........: 2020-01-07
    #ce =====================================================================================================================================
    Func _IBMACS_DLLCall($_iFunction, $_sDataString = "", $_iDataLength = 0, $_iReturnCode = "")
        $_Result = DllCall($IBMACS_DLL_FILE, "none", "hllapi", "word*", $_iFunction, "str", $_sDataString, "word*", $_iDataLength, "word*", $_iReturnCode)
            If @error Then Return SetError(@error, @extended, False)
            If $IBMACS_DEBUG_ENABLE Then
                If IsArray($_Result) Then 
                    _ArrayDisplay($_Result, $IBMACS_DLL_FILE & " function (" & $_iFunction & ")")
                Else
                    ConsoleWrite($IBMACS_DLL_FILE & " call response of function (" & $_iFunction & "): " & $_Result & @CRLF)
                EndIf
            EndIf

        Return $_Result
    EndFunc

    #cs #FUNCTION# ==========================================================================================================================
        Name...............: _IBMACS_ConnectPresentationSpace($_sSessionName = "A", $_bConnect = True)
        Description .......: The Connect Presentation Space function establishes a connection between your EHLLAPI application program and the host presentation space.
        Syntax.............: https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/emulator_programming08.htm#func1
                             https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/emulator_programming08.htm#func2

        Parameters ........: $_sSessionName - session letter (presentation space id)
        Return values .....: Success - True; 
                             Failure - False + @error + @extended

        Called functions...: _IBMACS_DLLCall
        Author ............: exorcistas@github.com
        Modified...........: 2020-01-07
    #ce =====================================================================================================================================
    Func _IBMACS_ConnectPresentationSpace($_sSessionName = "A", $_bConnect = True)
        Local $_iFunction = 1                   ;-- connect
        If NOT $_bConnect Then $_iFunction = 2  ;-- disconnect
        
        $_Result = _IBMACS_DLLCall($_iFunction, $_sSessionName, 0, "")
            If @error Then Return SetError(1, @error, False)
            If NOT IsArray($_Result) Then Return SetError(2, $_Result, False)
            If $_Result[4] <> 0 Then Return SetError(3, $_Result[4], False)

        Return True
    EndFunc

    #cs #FUNCTION# ==========================================================================================================================
        Name...............: _IBMACS_SendKey($_sKey)
        Description .......: The Send Key function is used to send either a keystroke or a string of keystrokes to the host presentation space.
        Syntax.............: https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/emulator_programming08.htm#func3
                             https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/emulator_programming08.htm#kbmnem

        Parameters ........: $_sKey - key mnemonic to send
        Return values .....: Success - True; 
                             Failure - False + @error + @extended

        Called functions...: _IBMACS_DLLCall
        Author ............: exorcistas@github.com
        Modified...........: 2020-01-07
    #ce =====================================================================================================================================
    Func _IBMACS_SendKey($_sKey)
        If StringLeft($_sKey, 1) <> "@" Then Return SetError(1, 0, False)
        Local $_iFunction = 3
        Local $_iDataLength = StringLen($_sKey)

        $_Result = _IBMACS_DLLCall($_iFunction, $_sKey, $_iDataLength, "")
            If @error Then Return SetError(1, @error, False)
            If NOT IsArray($_Result) Then Return SetError(2, $_Result, False)
            If $_Result[4] <> 0 Then Return SetError(3, $_Result[4], False)

        Return True
    EndFunc

    #cs #FUNCTION# ==========================================================================================================================
        Name...............: _IBMACS_CopyFieldToString($_iRow, $_iCol, $iLength)
        Description .......: The Copy Field to String function transfers characters from a field in the host-connected presentation space into a string.
        Syntax.............: https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/emulator_programming08.htm#func34

        Parameters ........: $_iRow - X coord; $_iCol - Y coord; $iLength - string length
        Return values .....: Success - string; 
                             Failure - False + @error + @extended

        Called functions...: _IBMACS_DLLCall
        Author ............: exorcistas@github.com
        Modified...........: 2020-01-07
    #ce =====================================================================================================================================
    Func _IBMACS_CopyFieldToString($_iRow, $_iCol, $_iLength)
        If (($_iRow = 0) OR ($_iCol = 0) OR ($_iLength = 0)) Then Return SetError(1, 0, False)
        $_iPos = $_iCol + (($_iRow-1)*80)
        Local $_iFunction = 34

        $_Result = _IBMACS_DLLCall($_iFunction, "", $_iLength, $_iPos)
            If @error Then Return SetError(2, @error, False)
            If NOT IsArray($_Result) Then Return SetError(3, $_Result, False)
            If $_Result[4] = 6 AND $_Result[2] <> "" Then
                ;-- skip error
            ElseIf $_Result[4] <> 0 Then 
                Return SetError(4, $_Result[4], False)
            EndIf

        $_sRawString = $_Result[2]
        $_sString = StringLeft($_sRawString, $_iLength)
        return $_sString
    EndFunc

    #cs #FUNCTION# ==========================================================================================================================
        Name...............: _IBMACS_CopyStringToField($_sString, $_iRow, $_iCol)
        Description .......: The Copy String to Field function transfers a string of characters into a specified field in the host-connected presentation space.
        Syntax.............: https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/emulator_programming08.htm#func33

        Parameters ........: $_iRow - X coord; $_iCol - Y coord; $_sString - string to set
        Return values .....: Success - True; 
                             Failure - False + @error + @extended
                             
        Called functions...: _IBMACS_DLLCall
        Author ............: exorcistas@github.com
        Modified...........: 2021-03-05
    #ce =====================================================================================================================================
    Func _IBMACS_CopyStringToField($_sString, $_iRow, $_iCol)
        If (($_iRow = 0) OR ($_iCol = 0) OR (StringLen($_sString) = 0)) Then Return SetError(1, 0, False)
        $_iPos = $_iCol + (($_iRow-1)*80)
        Local $_iFunction = 33
        $_iDataLength = StringLen($_sString) + 4
        ;~ $_sString = __ReplaceChars2ASCII($_sString)

        $_Result = _IBMACS_DLLCall($_iFunction, $_sString, $_iDataLength, $_iPos)
            If @error Then Return SetError(2, @error, False)
            If NOT IsArray($_Result) Then Return SetError(3, $_Result, False)
            If $_Result[4] <> 0 Then Return SetError(4, $_Result[4], False)

        Return True
    EndFunc

    #cs #FUNCTION# ==========================================================================================================================
        Name...............: _IBMACS_SetCursor($_iRow, $_iCol)
        Description .......: The Set Cursor function is used to set the position of the cursor within the host presentation space.
        Syntax.............: https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/emulator_programming08.htm#func40

        Parameters ........: $_iRow - X coord; $_iCol - Y coord
        Return values .....: Success - True; 
                             Failure - False + @error + @extended
                             
        Called functions...: _IBMACS_DLLCall
        Author ............: exorcistas@github.com
        Modified...........: 2020-01-07
    #ce =====================================================================================================================================
    Func _IBMACS_SetCursor($_iRow, $_iCol)
        If (($_iRow = 0) OR ($_iCol = 0)) Then Return SetError(1, 0, False)
        $_iPos = $_iCol + (($_iRow-1)*80)
        Local $_iFunction = 40

        $_Result = _IBMACS_DLLCall($_iFunction, "", 0, $_iPos)
            If @error Then Return SetError(2, @error, False)
            If NOT IsArray($_Result) Then Return SetError(3, $_Result, False)
            If $_Result[4] <> 0 Then Return SetError(4, $_Result[4], False)

        Return True
    EndFunc

    #cs #FUNCTION# ==========================================================================================================================
        Name...............: _IBMACS_Wait()
        Description .......: The Wait function checks the status of the host-connected presentation space. 
                             If the session is waiting for a host response (indicated by XCLOCK (X []) or XSYSTEM), 
                             the Wait function causes EHLLAPI to wait up to 1 minute to see if the condition clears.
                            
        Syntax.............: https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/emulator_programming08.htm#func4
        Parameters ........: -
        Return values .....: Success - True; 
                             Failure - False + @error + @extended
                             
        Called functions...: _IBMACS_DLLCall
        Author ............: exorcistas@github.com
        Modified...........: 2020-01-07
    #ce =====================================================================================================================================
    Func _IBMACS_Wait()
        Local $_iFunction = 4

        $_Result = _IBMACS_DLLCall($_iFunction, "", 0, "")
            If @error Then Return SetError(1, @error, False)
            If NOT IsArray($_Result) Then Return SetError(2, $_Result, False)
            If $_Result[4] <> 0 Then Return SetError(3, $_Result[4], False)

        Return True
    EndFunc

    #cs #FUNCTION# ==========================================================================================================================
        Name...............: _IBMACS_CopyPresentationScreen()
        Description .......: The Copy Presentation Space function copies the contents of the host-connected presentation space into a data string that you define in your EHLLAPI application program.
        Syntax.............: https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/emulator_programming08.htm#func5
        Parameters ........: -
        Return values .....: Success - screen text as string; 
                             Failure - False + @error + @extended
                             
        Called functions...: _IBMACS_DLLCall
        Author ............: exorcistas@github.com
        Modified...........: 2020-01-07
    #ce =====================================================================================================================================
    Func _IBMACS_CopyPresentationScreen()
        Local $_iFunction = 5

        $_Result = _IBMACS_DLLCall($_iFunction, "", 0, "")
            If @error Then Return SetError(1, @error, False)
            If NOT IsArray($_Result) Then Return SetError(2, $_Result, False)
            If $_Result[4] <> 0 Then Return SetError(3, $_Result[4], False)

        Return $_Result[2]
    EndFunc

    #cs #FUNCTION# ==========================================================================================================================
        Name...............: _IBMACS_ResetSystem()
        Description .......: The Reset System function reinitializes EHLLAPI to its starting state. 
                             The session parameter options are reset to their defaults. Event notification is stopped. 
                             The reserved host session is released. The host presentation space is disconnected. Keystroke intercept is disabled.
                             You can use the Reset System function during initialization or at program termination to reset the system to a known initial condition.

        Syntax.............: https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/emulator_programming08.htm#func21
        Parameters ........: -
        Return values .....: Success - True; 
                             Failure - False + @error + @extended
                             
        Called functions...: _IBMACS_DLLCall
        Author ............: exorcistas@github.com
        Modified...........: 2020-01-07
    #ce =====================================================================================================================================
    Func _IBMACS_ResetSystem()
        Local $_iFunction = 21

        $_Result = _IBMACS_DLLCall($_iFunction, "", 0, "")
            If @error Then Return SetError(1, @error, False)
            If NOT IsArray($_Result) Then Return SetError(2, $_Result, False)
            If $_Result[4] <> 0 Then Return SetError(3, $_Result[4], False)

        Return True
    EndFunc

    #cs #FUNCTION# ==========================================================================================================================
        Name...............: _IBMACS_SetSessionParameters($_sParams)
        Description .......: The Set Session Parameters function lets you change certain default session options in EHLLAPI for all sessions.
        Syntax.............: https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/emulator_programming08.htm#func9
        Parameters ........: -
        Return values .....: Success - True; 
                             Failure - False + @error + @extended
                             
        Called functions...: _IBMACS_DLLCall
        Author ............: exorcistas@github.com
        Modified...........: 2020-01-07
    #ce =====================================================================================================================================
    Func _IBMACS_SetSessionParameters($_sParams)
        Local $_iFunction = 9

        $_Result = _IBMACS_DLLCall($_iFunction, $_sParams, 0, "")
            If @error Then Return SetError(1, @error, False)
            If NOT IsArray($_Result) Then Return SetError(2, $_Result, False)
            If $_Result[4] <> 0 Then Return SetError(3, $_Result[4], False)

        Return True
    EndFunc

#Region API_CORE_FUNCTIONS