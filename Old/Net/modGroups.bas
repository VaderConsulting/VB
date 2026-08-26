Attribute VB_Name = "modGroups"
Public Declare Function NetLocalGroupAddMembers Lib "Netapi32" (ByVal psServer As Long, ByVal psLocalGroupName As Long, ByVal Level As Long, pPtrBuffer As Long, ByVal membercount As Long) As Long

Public Type LOCALGROUP_MEMBERS_INFO_3
    DomainAndName As Long
End Type
