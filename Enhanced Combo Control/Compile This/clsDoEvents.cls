VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDoEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================
' Class: clsDoEvents
' Author: John Baughman
' Date: 2002-04-04
' Purpose: Provide a means for determining if we actually need to run a DoEvents in a loop.
'          Uses a Win32 API call (GetQueueStatus) to examine the Windows message queue
'          and return what messages are waiting. The method (GetInputState) then fires a
'          DoEvents if a desired message is waiting.
'
' Source: An article from Planet Source Code and John Galanopoulos
'         http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?_
'         lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=206793&strAuthorName=John%20Galanopoulos&_
'         txtMaxNumberOfEntriesPerPage=25
'         (Remove trailing "_" characters from URL before using)
' ======================================================================================

' API Constants and declare
' Constants used by GetQueueStatus API function
Private Const QS_HOTKEY = &H80
Private Const QS_KEY = &H1
Private Const QS_MOUSEBUTTON = &H4
Private Const QS_MOUSEMOVE = &H2
Private Const QS_PAINT = &H20
Private Const QS_POSTMESSAGE = &H8
Private Const QS_SENDMESSAGE = &H40
Private Const QS_TIMER = &H10
Private Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT = (QS_MOUSE Or QS_KEY)
Private Const QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or _
                             QS_HOTKEY)
Private Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or _
                            QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Private Const QS_MESSAGES = (QS_POSTMESSAGE Or QS_SENDMESSAGE)                      ' Not MS standard constant
Private Const QS_STANDARD = (QS_HOTKEY Or QS_KEY Or QS_MOUSEBUTTON Or QS_PAINT)     ' Not MS standard constant

' What will determine if we will need to DoEvents
Private Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long

' Enumerator to determine what messages are watched
Public Enum QueueMessagesUsed
    All_Inputs = QS_ALLINPUT
    All_Events = QS_ALLEVENTS
    Standard = QS_STANDARD
    Messages = QS_MESSAGES
    InputOnly = QS_INPUT
    Mouse = QS_MOUSE
    MouseMove = QS_MOUSEMOVE
    Timer = QS_TIMER
End Enum

' Local variable for our messages watched property
Private m_lQueueUsed As QueueMessagesUsed

' Our property to determine what messages to watch.
' Can be '+' to set multiple messages, but we do have a couple of
' these preset in the constants.
Public Property Let QueueUsed(ByVal vData As QueueMessagesUsed)
    m_lQueueUsed = vData
End Property

Public Property Get QueueUsed() As QueueMessagesUsed
    QueueUsed = m_lQueueUsed
End Property

' Our only method that determines if we have any messages in the queue
' waiting to be processed. These messages are determined by constants and
' our property. Fires a DoEvents if GetQueueStatus <> 0
Public Sub GetInputState()
    If GetQueueStatus(m_lQueueUsed) = 0 Then Exit Sub
    
    DoEvents
End Sub


