VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBP"
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

Private Type tModuleInfo
  name as string
  kind as EDependencyType
  isPredeclared as boolean
  path as string
End Type

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
  if checkFileExists(sUrl) then
    Dim info as tModuleInfo: info = getModuleInfo(sUrl)
  end if
End Function

'@param {string} URL of module
Public Function RequireModule(ByVal sUrl as string, ByVal iLocType as ELocationType) as Boolean
  Dim sName as string
  if not checkFileExists(sURL) then
    Dim info as tModuleInfo: info = getModuleInfo(sUrl)
  end if
End Function

'@param {string} URL of frm file
'@param {string} URL of frx file
Public Function RequireUserForm(ByVal sUrlFrm as string, ByVal sUrlFrx as string, ByVal iLocType as ELocationType) as Boolean
  Dim sName as string
  if not checkFileExists(sURLFrm) then
    Dim info as tModuleInfo: info = getModuleInfo(sUrl)
  end if
End Function

'@param {string} Url of xlsm/xlsa/xls
Public Function RequireReference(ByVal sUrl as string, ByVal iLocType as ELocationType) as Boolean
  if not checkFileExists(sURL) then
    Dim ref as Reference: set ref = ThisWorkbook.VBProject.References.AddFromFile(sURL)
    Dim sName as string: sName = ref.name
    On Error Resume Next
    Application.Run sName & ".InstallDependencies"
  end if
End Function

'@param {string} Url like "C:\...\xxx.json" or "http://.../xxx.json"
Public Function RequireSheet(ByVal sURL as string, ByVal iLocType as ELocationType) as Boolean
  if not checkFileExists(sURL) then
    '..
  end if
End Function




Private Function checkModuleExists(ByVal sFilePath as string, ByRef sModuleName as string) as Boolean
  Dim bIsPredeclared as boolean
  sModuleName = getModuleName(sFilePath, bIsPredeclared)
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
'      Attribute VB_Name = "VBP"
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
Private Function getModuleInfo(ByVal sFilePath as string) as tModuleInfo
  Dim ff as Long: ff=FreeFile()
  Dim info as tModuleInfo
  info.path = sFilePath
  Open sFilePath for Input Access Read as #ff
    Dim bRunning as boolean: bRunning = true
    Dim iState as Long: iState = 1
    While bRunning
      Dim sLine as string: Line Input #FileNum, sLine 
      Dim iType as EDependencyType
      select case iState
        case 1
          select case true
            case sLine = "VERSION 1.0 CLASS"
              info.kind = EDependencyType.DTypeClass
              iState = 2
            case sLine = "Version 5#" 'Userform
              info.kind = EDependencyType.DTypeUserform
              iState = 2
            case else
              info.kind = EDependencyType.DTypeNormal
              iState = 2
          end select
        case 2
          Dim iStart as Long, iEnd as Long, iStatePart as long
          select case info.kind
            case EDependencyType.DTypeClass
              if sLine like "Attribute*VB_Name*" then
                iStart = instr(1,sLine,"""")
                iEnd = instr(iStart+1,sLine,"""")
                info.name = mid(sLine, iStart, iEnd-iStart)
                iStatePart=iStatePart+1
              elseif sLine like "Attribute*VB_PredeclaredId*" then
                info.isPredeclared = sLine like "*[Tt][Rr][Uu][Ee]"
                iStatePart=iStatePart+1
              end if
              if iStatePart = 2 then iState = 3
            case EDependencyType.DTypeUserform
              if sLine like "Attribute*VB_Name*" then
                iStart = instr(1,sLine,"""")
                iEnd = instr(iStart+1,sLine,"""")
                info.name = mid(sLine, iStart, iEnd-iStart)
                iStatePart=iStatePart+1
              elseif sLine like "Attribute*VB_PredeclaredId*" then
                info.isPredeclared = sLine like "*[Tt][Rr][Uu][Ee]"
                iStatePart=iStatePart+1
              end if
              if iStatePart = 2 then iState = 3
            case EDependencyType.DTypeNormal
              if sLine like "Attribute*VB_Name*" then
                iStart = instr(1,sLine,"""")
                iEnd = instr(iStart+1,sLine,"""")
                info.name = mid(sLine, iStart, iEnd-iStart)
              end if
              if iStatePart = 1 then iState = 3
          end select
        case 3
          bRunning = false
      end select
    Wend
  Close #ff
End Function
