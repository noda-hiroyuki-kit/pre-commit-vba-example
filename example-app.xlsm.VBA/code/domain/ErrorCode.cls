VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("domain")
Option Explicit

Public Enum AppError
    [_Base_Error] = vbObjectError + 1024
    INVALID_SEMVER
    INVALID_VALUE
    HAS_CREATED
    MISSING_PROPERTY
    NOT_CREATED
    NOT_FOUND_CAN_REF
    RUBBERDUCK_REFERENCE_REMOVE_FAILED
    RUBBERDUCK_TYPELIB_NOT_FOUND
    VERSION_ROLLBACK_NOT_ALLOWED
End Enum
