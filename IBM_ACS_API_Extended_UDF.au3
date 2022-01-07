#cs # IBM_ACS_Extended_UDF # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Name..................: IBM_ACS_Extended_UDF
    Description...........: Extended UDF for 'IBM i Access Client Solution' API
    Dependencies..........: IBM_ACS_API_Core_UDF.au3, corresponding DLLs
    Documentation.........: https://www.ibm.com/support/knowledgecenter/en/ssw_ibm_i_71/rzaik/rzaikemulator.htm
                            https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/emulator_programming08.htm

    Author................: exorcistas@github.com
    Modified..............: 2020-01-03
    Version...............: v0.3.1rc
#ce ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#AutoIt3Wrapper_UseX64=n    ;-- required to run in 32bit environment
#include-once
#include <IBM_ACS_API_Core_UDF.au3>

#Region FUNCTIONS_LIST
#cs	===================================================================================================================================
    _IBMACS_Ext_ConnectSession($_sSession = "A")
    _IBMACS_Ext_DisconnectSession($_sSession = "A")
    _IBMACS_Ext_GetText($_iRow, $_iCol, $iLength)
    _IBMACS_Ext_SetText($_sString, $_iRow, $_iCol)
    _IBMACS_Ext_SendKeyAndWait($_sKey)
    _IBMACS_Ext_SendReset()
    _IBMACS_Ext_SendEnter()
    _IBMACS_Ext_SendFunctionKey($_iFnKey)
    _IBMACS_Ext_EraseField($_iRow = 0, $_iCol = 0)
#ce	===================================================================================================================================
#EndRegion FUNCTIONS_LIST

#Region IBM_ACS_EXTENDED_FUNCTIONS

    Func _IBMACS_Ext_ConnectSession($_sSession = "A")
        _IBMACS_OpenDll()
            If @error Then Return SetError(@error, @extended, False)
        _IBMACS_ConnectPresentationSpace($_sSession, True)
            If @error Then Return SetError(@error, @extended, False)
        _IBMACS_Wait()
        Return True
    EndFunc

    Func _IBMACS_Ext_DisconnectSession($_sSession = "A")
        _IBMACS_Wait()
        _IBMACS_ConnectPresentationSpace($_sSession, False)
            If @error Then Return SetError(@error, @extended, False)
        _IBMACS_CloseDll()
            If @error Then Return SetError(@error, @extended, False)
        Return True
    EndFunc

    Func _IBMACS_Ext_GetText($_iRow, $_iCol, $iLength)
        $_sString = _IBMACS_CopyFieldToString($_iRow, $_iCol, $iLength)
            If @error Then Return SetError(@error, @extended, False)

        Return $_sString
    EndFunc

    Func _IBMACS_Ext_SetText($_sString, $_iRow, $_iCol)
        _IBMACS_CopyStringToField($_sString, $_iRow, $_iCol)
            If @error Then Return SetError(@error, @extended, False)

        Return True
    EndFunc

    Func _IBMACS_Ext_SendKeyAndWait($_sKey)
        _IBMACS_Wait()
        _IBMACS_SendKey($_sKey)
            If @error Then Return SetError(@error, @extended, False)
        _IBMACS_Wait()

        Return True
    EndFunc

    Func _IBMACS_Ext_SendReset()
        _IBMACS_Ext_SendKeyAndWait("@R")
            If @error Then Return SetError(@error, @extended, False)

        Return True
    EndFunc

    Func _IBMACS_Ext_SendEnter()
        _IBMACS_Ext_SendKeyAndWait("@E")
            If @error Then Return SetError(@error, @extended, False)
                
        Return True
    EndFunc

    Func _IBMACS_Ext_SendFunctionKey($_iFnKey)
        Switch $_iFnKey
            Case 1 To 9
                _IBMACS_Ext_SendKeyAndWait("@" & $_iFnKey)
            Case 10
                _IBMACS_Ext_SendKeyAndWait("@a")
            Case 11
                _IBMACS_Ext_SendKeyAndWait("@b")
            Case 12
                _IBMACS_Ext_SendKeyAndWait("@c")
            Case 13
                _IBMACS_Ext_SendKeyAndWait("@d")
            Case 14
                _IBMACS_Ext_SendKeyAndWait("@e")
            Case 15
                _IBMACS_Ext_SendKeyAndWait("@f")
            Case 16
                _IBMACS_Ext_SendKeyAndWait("@g")
            Case 17
                _IBMACS_Ext_SendKeyAndWait("@h")
            Case 18
                _IBMACS_Ext_SendKeyAndWait("@i")
            Case 19
                _IBMACS_Ext_SendKeyAndWait("@j")
            Case 20
                _IBMACS_Ext_SendKeyAndWait("@k")
            Case 21
                _IBMACS_Ext_SendKeyAndWait("@l")
            Case 22
                _IBMACS_Ext_SendKeyAndWait("@m")
            Case 23
                _IBMACS_Ext_SendKeyAndWait("@n")
            Case 24
                _IBMACS_Ext_SendKeyAndWait("@o")
            Case Else
                Return SetError(25, 0, False)
        EndSwitch
        If @error Then Return SetError(@error, @extended, False)

        Return True
    EndFunc

    Func _IBMACS_Ext_EraseField($_iRow = 0, $_iCol = 0)
        If NOT (($_iRow = 0) AND ($_iCol = 0)) Then 
            _IBMACS_SetCursor($_iRow, $_iCol)
            If @error Then Return SetError(100+@error, @extended, False)
        EndIf
        _IBMACS_Ext_SendKeyAndWait("@F")
            If @error Then Return SetError(200+@error, @extended, False)
                
        Return True
    EndFunc

#EndRegion IBM_ACS_EXTENDED_FUNCTIONS