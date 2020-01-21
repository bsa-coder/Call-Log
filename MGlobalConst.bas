Attribute VB_Name = "MGlobalConst"
Option Explicit

'This module contains global constants for CLIENT

'FMain constants
Global Const clDEFAULTLASTACTION As Integer = -1
Global Const clDEFAULTLASTCOMPANY As Integer = 0
Global Const clDEFAULTLASTCONTACT As Integer = 10

'True if BLServer connection is active
Global ServerConnectionUp As Boolean

'Position or percentage through upload
Global LoadProgress As Integer
Global LoadProgressMax As Integer
Global LoadProgressActive As Boolean

Global lCurrentRecordID As Long
Global fEditCaseID As Boolean
