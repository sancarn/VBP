VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

#if VBA7 then
  Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As LongPtr) As Long
#else
  Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#end if

Private Enum EDependencyType
  DTypeNormal
  DTypeClass
  DTypeUserform
  DTypeReference
  DTypeSheet
End Enum

Private Enum ELocationType
  LocationFileSystem
  LocationInternet
End Enum

Public ErrorDescription as string

Public Function InstallDependencies() as Boolean

End Function

'@param {string} Url either `C:\...\xxx.cls` or `http://.../xxx.cls` or `http://.../xxx.xls[ma]?`
Public Function Requires(ByVal sUrl as string) as Boolean
  Dim sFile as string
  Dim iLocType as ELocationType: iLocType = getLocationFromURL(sUrl)
  select case iLocType
    case LocationFileSystem
      sFile = sUrl
    case LocationInternet
      if not downloadInternetFile(sUrl, sFile) then
        ErrorDescription = "Cannot download file."
        Requires = false
        Exit Function
      end if
  end select
  
  Dim iType as EDependencyType: iType = getTypeFromURL(sUrl)
  select case iType
    case DTypeNormal
      Requires = RequireModule(sFile, iLocType)
    case DTypeClass
      Requires = RequireClass(sFile, iLocType)
    case DTypeUserform
      Dim sFRM as string: sFRM = sFile
      Dim sFRX as string: sFRX = replace(sFile, ".frm", ".frx")
      Requires = RequireUserForm(sFRM, sFRX, iLocType)
    case DTypeReference
      Requires = RequireReference(sFile, iLocType)
    case DTypeSheet
      Requires = RequireSheet(sFile, iLocType)
    case else
      ErrorDescription = "Dependency type couldn't be deciphered from URL. Use `RequireClass` or one of the alternatives"
      Requires = false
  end select
End Function

'@param {string} URL of class
Public Function RequireClass(ByVal sUrl as string, ByVal iLocType as ELocationType) as Boolean
  if not checkModuleExists(sURL) then
    '..
  end if
End Function

'@param {string} URL of module
Public Function RequireModule(ByVal sUrl as string, ByVal iLocType as ELocationType) as Boolean
  if not checkModuleExists(sURL) then
    '..
  end if
End Function

'@param {string} URL of frm file
'@param {string} URL of frx file
Public Function RequireUserForm(ByVal sUrlFrm as string, ByVal sUrlFrx as string, ByVal iLocType as ELocationType) as Boolean
  if not checkModuleExists(sURL) then
    '..
  end if
End Function

'@param {string} Url of xlsm/xlsa/xls
Public Function RequireReference(ByVal sUrl as string, ByVal iLocType as ELocationType) as Boolean
  if not checkModuleExists(sURL) then
    '..
  end if
End Function

'@param {string} Url like "C:\...\xxx.json" or "http://.../xxx.json"
Public Function RequireSheet(ByVal sURL as string, ByVal iLocType as ELocationType) as Boolean
  if not checkModuleExists(sURL) then
    '..
  end if
End Function




Private Function checkModuleExists(ByVal sFilePath as string) as Boolean
  Dim bIsPredeclared as boolean
  Dim sModuleName as string: sModuleName = getModuleName(sFilePath, bIsPredeclared)
End Function

'Download a file from the internet to file
'@param {string} The input url to download the file from.
'@param {ByRef string} The output file destination.
'@returns {Boolean} Whether download was successful or not.
Private Function downloadInternetFile(ByVal sURL as string, ByRef sFileDest as string) as boolean
  On Error GoTo Failed
    sFileDest = Environ("Temp") & "\" & getGUID()
    Dim iRetVal as long: iRetVal = URLDownloadToFile(0, sURL, sFileDest, )
    downloadInternetFile = (iRetVal = 0) and (Len(Dir(sFileDest))>0)
    Exit Function
Failed:
  On Error GoTo -1
End Function

'Obtains a new GUID
Private Function getGUID() As String
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(CLng(Rnd() * 15.9999)), 1, 1)
  Next
End Function

'Obtain the module name from a class name from file content
'@param {string} File to obtain content from
'@param {boolean} Whether the file is predeclared or not.
'@remarks
'  # Class example
'      VERSION 1.0 CLASS
'      BEGIN
'        MultiUse = -1  'True
'      END
'      Attribute VB_Name = "VBPM"
'      Attribute VB_GlobalNameSpace = False
'      Attribute VB_Creatable = False
'      Attribute VB_PredeclaredId = True
'      Attribute VB_Exposed = False
'  # Module example
'      Attribute VB_Name = "Main"
'  # UserForm example
'      Version 5#
'      Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUserform_Icon
'         Caption = "Icon in a userform"
'         ClientHeight = 4824
'         ClientLeft = 108
'         ClientTop = 456
'         ClientWidth = 9024.001
'         OleObjectBlob   =   "frmUserform_Icon.frx":0000
'         StartUpPosition = 1    'CenterOwner
'      End
'      Attribute VB_Name = "frmUserform_Icon"
'      Attribute VB_GlobalNameSpace = False
'      Attribute VB_Creatable = False
'      Attribute VB_PredeclaredId = True
'      Attribute VB_Exposed = False
Private Function getModuleName(ByVal sFilePath as string, ByRef isPredeclared as boolean) as string
  Dim ff as Long: ff=FreeFile()
  Open sFilePath for Input Access Read as #ff
    bSearching = true
    While bSearching
      Dim sLine as string: Line Input #FileNum, sLine 
      select case true
        case sLine = "VERSION 1.0 CLASS"
          
        case sLine = "Version 5#" 'Userform
          isPredeclared = true
      end select

    Wend
  Close #ff
End Function