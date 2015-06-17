' ��������� UTF-8 ��� ANSI 
Attribute VB_Name = "DBCreateMDWSystem"
Option Explicit
Option Base 0
'12345678901234567890123456789012345678901234567890123456789012345678901234567890
'FullPath=Environ("UserProfile")&"\Application Data\Microsoft\Access\System.mdw"

Private ArrByte(126975) As Byte, iND As Long

Public Sub ByteArrayToSystemMdw(ByVal FullPath As String) ' rev.200
Attribute ByteArrayToSystemMdw.VB_Description = "������ ������� 12/02/2014 (danilin)"
  iND = LBound(ArrByte)
  Debug.Print "Loading  0%": Part00: Part01: Part02: Part03: Part04: Part05
  Debug.Print "Loading 14%": Part06: Part07: Part08: Part09: Part0A: Part0B
  Debug.Print "Loading 27%": Part0C: Part0D: Part0E: Part0F: Part10: Part11
  Debug.Print "Loading 41%": Part12: Part13: Part14: Part15: Part16: Part17
  Debug.Print "Loading 54%": Part18: Part19: Part1A: Part1B: Part1C: Part1D
  Debug.Print "Loading 68%": Part1E: Part1F: Part20: Part21: Part22: Part23
  Debug.Print "Loading 82%": Part24: Part25: Part26: Part27: Part28: Part29
  Debug.Print "Loading 95%": Part2A: Part2B: Debug.Print "Loaded 100%"
  If Err.Number = 6 Then Err.Clear ' ������� ������ ������������
  Open FullPath For Binary Access Write As #1: Put #1, , ArrByte: Close #1
End Sub

Private Function SetAsByte(ByRef ArrVar As Variant) As Long
Dim Counter As Integer
  On Error Resume Next ' �����! ��������� Value = Missing, �������� ����
    For Counter = LBound(ArrVar) To UBound(ArrVar)
      ArrByte(Counter + iND) = CByte(ArrVar(Counter))
    Next Counter: SetAsByte = Counter + iND
End Function

Private Sub Part00()
Dim ArrVar As Variant
  ArrVar = Array(, 1, , , 74, 101, 116, 32, 83, 121, 115, 116, 101, 109, 32, _
    68, 66, 32, 32, , 1, , , , 181, 110, 3, 98, 96, 9, 194, 85, 233, 169, 103, _
    114, 64, 63, , 156, 126, 159, 144, 255, 133, 154, 49, 197, 127, 186, 237, _
    48, 187, 223, 204, 157, 99, 217, 228, 195, 152, 70, 128, 128, 23, 238, _
    8, 89, 236, 55, 211, 230, 156, 250, 72, 252, 40, 230, 157, 20, 138, 96, _
    218, 54, 123, 54, 123, 208, 223, 177, 249, 86, 19, 67, 65, 13, 177, 51, _
    186, 195, 121, 91, 28, 23, 124, 42, 163, 224, 124, 153, 24, 31, 152, 253, _
    74, 26, 243, 152, 189, 111, 132, 102, 95, 149, 248, 208, 137, 36, 133, _
    103, 198, 31, 39, 68, 210, 238, 207, 101, 237, 255, 7, 199, 70, 161, 120, _
    22, 12, 237, 233, 45, 98, 212, 84, 6, , , 52, 46, 48)
  iND = SetAsByte(ArrVar) + 3425 ' ���������� ������
  ArrVar = Array(1, 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , _
    1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, _
    , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , _
    1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, _
    , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , _
    1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, _
    , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , _
    1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, _
    , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , _
    1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, _
    , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , _
    1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, _
    , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , _
    1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, _
    , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , _
    1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, _
    , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , _
    1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1, , 1)
  iND = SetAsByte(ArrVar)
  ArrVar = Array(215, 125, 255, 183, 213, 199, 66, 125, 234, 120, 163, 68, 52, _
    47, 140, 8, 152, 152, 48, 180, 123, 251, 99, 176, 196, 19, 250, 239, 180, _
    236, 109, 74, 50, 119, 159, 18, 116, 10, 6, 230, 200, 91, 239, 226, 172, _
    247, 157, 184, 252, 119, 233, 243, 132, 237, 197, 223, 206, 241, 239, 69, _
    229, 94, 232, 179, 172, 169, 231, 96, 224, 205, 4, 130, 186, 214, 97, 140, _
    142, 161, 24, 248, 183, 57, 174, 95, 67, 35, 55, 182, 77, 141, 117, 234, _
    240, 209, 235, 51, 149, 199, 14, 193, 121, 13, 181, 141, 78, 8, 220, 96, _
    62, 124, 17, 226, 112, 144, 252, 161, 86, 115, 54, 87, 182, 17, 114, 45, _
    76, 173, 42, 172, 9, 130, 169, 179, 140, 96, 206, 186, 19, 142, 162, 30, _
    96, 3, 46, 176, 6, 11, 228, 52, 163, 128, 148, 206, 54, 214, 224, 142, _
    167, 123, 176, 171, 36, 109, 248, 226, 195, 134, 98, 253, 161, 232, 3, _
    176, 205, 181, 74, 106, 130, 142, 135, 147, 93, 94, 122, 241, 69, 104, _
    152, 246, 182, 249, 91, 111, 158, 114, 82, 188, 201, 106, 213, 35, 61, 19, _
    93, 2, 197, 174, 248, 38, 153, 211, 157, 227, 36, 169, 180, 187, 212, 33, _
    232, 209, 202, 99, 29, 101, 42, 109, 239, 199, 157, 121, 67, 147, 232, _
    136, 211, 176, 12, 219, 24, 46, 86, 53, 62, 176, 192, 78, 134, 218, 202, _
    177, 185, 115, 230, 6, 206, 8, 41, 15, 208, 72, 1, 118, 103, 101, 119, 94, _
    137, 140, 29, 78, 160, 246, 47, 144, 150, 183, 11, 90, 11, 204, 89, 254)
  iND = SetAsByte(ArrVar)
  ArrVar = Array(42, 52, 156, 111, 159, 40, 227, 197, 39, 206, 226, 58, 159, _
    42, 65, 14, 104, 217, 11, 125, 83, 221, 244, 16, 97, 158, 178, 55, 90, 16, _
    8, 113, 211, 134, 177, 18, 139, 44, 90, 180, 221, 228, 174, 4, 224, 75, _
    138, 34, 4, 216, 188, 47, 120, 99, 166, 177, 239, 34, 183, 214, 221, 175, _
    54, 235, 94, 147, 234, 219, 176, 68, 26, 238, 22, 14, 35, 61, 103, 63, _
    44, 37, 38, 2, 108, 101, 148, 37, 47, 181, 206, 77, 212, 32, 34, 26, 133, _
    239, 172, 190, 118, 14, 14, 109, 176, 212, , 112, 179, 131, 159, 143, 116, _
    52, 97, 148, 101, 86, 190, 99, 50, 162, 219, 102, 195, 166, 176, 160, 168, _
    13, 105, 127, 27, 226, 160, 230, 115, 164, 166, 145, 1, 17, 83, 186, 200, _
    163, 40, 132, 157, 136, 196, 197, 110, 152, 227, 54, 33, 115, 163, 182, _
    246, 135, 123, 68, 176, 131, 22, 24, 182, 151, 76, 3, 108, 88, 241, 218, _
    181, 104, 4, 22, 177, 212, 167, 40, 154, 183, 242, 105, 229, 48, 181, 133, _
    114, 190, 210, 171, 249, 203, 99, 101, 254, 110, 36, 3, 136, 234, 45, 14, _
    85, 16, 179, 18, 149, 244, 92, 59, 21, 176, 198, 237, 206, 135, 168, 39, _
    4, 143, 255, 217, 56, 78, 250, 116, 55, 7, 102, 119, 91, 24, 151, 191, _
    195, 228, 224, 133, 172, 79, 183, 95, 56, 138, 83, 89, 170, 185, 3, 143, _
    242, 116, 238, 17, 77, 7, 208, 132, 141, 246, 55, 150, 115, 4, 28, 204, _
    230, 109, 141, 226, 194, 59, 188, 230, 130, 118, 187, 38, 1, 180, 45, 26)
  iND = SetAsByte(ArrVar)
  ArrVar = Array(20, 204, 169, 83, 140, 178, 100, 120, 234, 103, 45, 224, _
    74, 152, 3, 117, 15, 241, 77, 206, 176, 202, 220, 78, 157, 35, 229, 87, _
    72, 198, 125, 18, 133, 90, 46, 23, 120, 199, 1, 40, 99, 120, 133, 8, 65, _
    166, 137, 21, 233, 38, 87, 235, 109, 234, 228, 162, 222, 76, 192, 97, 111, _
    168, 88, 74, 237, 59, 9, 151, 215, 24, 71, 248, 232, 75, 153, 89, 90, 207, _
    245, 217, 117, 134, 87, 156, 135, 246, 235, 25, 177, 227, 119, 31, 208, _
    44, 143, 134, 243, 44, 248, 236, 131, 215, 88, 61, 63, 51, 125, 185, 28, _
    163, 41, 54, 162, 79, , 129, 27, 199, 59, 44, 39, 177, 225, 254, 87, 146, _
    33, 225, 48, 130, 231, 116, 231, 7, 89, 111, 225, 136, 79, 100, 22, 156, _
    73, 216, 164, 133, 63, 62, 92, 235, 110, 82, 23, 197, 220, 113, 58, 111, _
    217, 62, 6, 197, 217, 241, 229, 25, 114, 17, 213, , 111, 115, 240, 8, 101, _
    112, 64, 231, 179, 39, 32, 103, 59, 217, 204, 142, 223, 164, 187, 161, _
    174, 27, 18, 251, 178, 5, 181, 67, 227, 221, 149, 180, 105, 204, 197, 67, _
    202, 7, 117, 54, 191, 76, 225, 184, 7, 128, 251, 237, 233, 124, 194, 32, _
    203, 128, 226, 237, 235, 178, 254, 140, 16, 82, 84, 152, 68, 200, 38, 10, _
    111, 6, 100, 201, 241, 239, 143, 187, 85, 104, 26, 202, 127, 152, 143, _
    171, 82, 48, 161, 157, 140, 95, 161, 151, 156, 18, 101, 97, 120, 182, 224, _
    196, 203, 240, 203, 10, 77, 24, 106, 45, 156, 201, 177, 122, 56, 118, 118)
  iND = SetAsByte(ArrVar)
  ArrVar = Array(230, 181, 7, 206, 83, 37, 186, 233, 131, 13, 166, 224, 235, _
    254, 250, 133, 36, 99, 103, 131, 142, 48, 154, 228, 212, 33, 216, 50, 132, _
    214, 235, 102, 200, 244, 240, 112, 118, 193, 84, 114, 136, 81, 103, 53, _
    112, 230, 34, 235, 20, 19, 16, 31, 253, 249, 176, 48, 47, 131, 130, 224, _
    204, 161, 90, 89, 107, 249, 12, 98, 105, 30, 51, 136, 186, 233, 138, 106, _
    117, 141, 66, 187, 3, 117, 115, 64, 216, 23, 65, 223, 238, 161, 43, 47, _
    191, 100, 208, 138, 1, 114, 141, 124, 125, 95, 255, 234, 81, 137, 48, 201, _
    204, 52, 82, 232, 241, 151, 137, 136, 14, 6, 166, 52, 87, 42, 73, 7, 161, _
    111, 202, 122, 207, 148, 16, 159, 152, 172, 210, 135, 21, 168, 177, 28, _
    56, 47, 96, 249, 229, 122, 246, 73, 9, 134, 15, , 86, 158, 57, 181, 138, _
    233, 60, 76, 233, 177, 118, 212, 23, 195, 121, 162, 185, 68, 118, 9, 132, _
    71, 138, 233, 155, 187, 223, 171, 72, 117, 246, 105, 37, 74, 52, 199, 137, _
    211, 160, 26, 172, 73, 148, 151, 247, 40, 33, 119, 153, 199, 222, 170, _
    180, 18, 118, 102, 172, 199, 158, 179, 229, 200, 80, 217, 103, 126, 232, _
    94, 64, 16, 241, 227, 159, 119, 167, 123, 41, 45, 209, 207, 145, 67, 75, _
    113, 250, 95, 191, 130, 236, 113, 54, 141, 195, 5, 124, 226, 239, 158, _
    171, 27, 217, 163, 113, 141, 21, 81, 244, 205, 1, 149, 175, 251, 156, 37, _
    , 46, 90, 188, 151, 229, 75, 236, 198, 56, 227, 170, 228, 175, 104, 183)
  iND = SetAsByte(ArrVar)
  ArrVar = Array(96, 139, 180, 182, 234, 200, 164, 211, 226, 150, 94, 60, _
    148, 229, 88, 32, 21, 165, 255, 31, 76, 93, 110, 208, 240, 250, 110, 148, _
    205, 249, 81, 44, 32, 5, 250, 212, 42, 96, 165, 188, 248, 199, 78, 186, _
    145, 74, 166, 167, 17, 116, 194, 26, 225, 189, 125, 253, 78, 2, 221, 129, _
    243, 31, 253, 38, 104, 164, 228, 184, 129, 123, 230, 226, 135, 57, 5, 129, _
    19, 5, 74, 235, 223, 120, 26, 138, 65, 207, 153, 7, 93, 55, 87, 232, 94, _
    241, 8, 167, 4, 120, 213, 94, 146, 180, 220, 133, 79, 43, 46, 118, 200, _
    63, 58, 108, 31, 207, 16, 226, 231, 31, 35, 94, 220, 137, 82, 156, 199, _
    78, 227, 3, 109, 214, 69, 168, 102, 112, 111, 208, 255, 191, 163, 109, _
    119, 177, 87, 215, 8, 152, 68, 184, 117, 4, 138, 34, 107, 154, 72, 185, _
    71, 34, 39, 203, 106, 218, 168, 62, 55, 125, 37, 105, 23, 230, 72, 98, _
    31, 30, 171, 154, 171, 18, 222, 170, 128, 40, 15, 14, 249, 170, 213, 175, _
    24, 200, 79, 114, 51, 243, 217, 119, 154, 131, 188, 205, 172, 208, 232, _
    32, 222, 180, 30, 21, 99, 194, 12, 7, 80, 187, 198, 82, 110, 227, 113, _
    189, 102, 208, 26, 166, 144, 244, 49, 88, 231, 212, 160, 34, 158, 99, 97, _
    84, 90, 192, , 192, 159, 205, 80, 153, 168, 59, 197, 56, 174, 35, 239, _
    31, 173, 32, 149, 196, 140, 18, 18, 121, 231, 184, 55, 71, 78, 89, 11, _
    187, 127, 173, 138, 158, 171, 137, 116, 247, 244, 205, 112, 197, 195, 57)
  iND = SetAsByte(ArrVar)
  ArrVar = Array(163, 178, 72, 205, 184, 70, 49, 218, 27, 113, 119, 58, 75, _
    117, 101, 9, 51, 124, 83, 252, 185, 202, 99, 111, 174, 51, 73, 172, 68, _
    78, 20, 27, 125, 220, 21, 35, 81, 123, 6, 92, 102, 1, 16, 67, 130, 246, _
    206, 72, 190, 165, 115, 64, 64, 160, 250, 242, 130, 117, 22, 32, 114, 1, _
    52, 184, 174, 51, 66, 199, 149, 171, 133, 225, 43, 255, 91, 221, 236, 36, _
    31, 160, 53, 100, 136, 152, 242, 11, 62, 90, 77, 179, 3, 117, 116, 36, _
    58, 53, 80, 89, 29, 177, 126, 144, 166, 56, 19, 222, 253, 14, 103, 69, _
    99, 249, 51, 47, 130, 36, 64, 67, 1, 207, 157, 238, 117, 159, 4, 217, 218, _
    152, 40, 161, 52, 226, 250, 218, 213, 210, 242, 114, 234, 68, 197, 18, _
    64, 156, 110, 123, 71, 231, 27, 118, 93, 13, 139, 71, 218, 65, 166, 50, _
    171, 109, 143, 163, 148, 170, 212, 186, 101, 202, 239, 30, , 220, 10, 174, _
    244, 8, 139, 117, 226, 70, 214, 192, 69, 99, 127, 83, 220, 144, 130, 134, _
    252, 41, 138, 224, 124, 204, 146, 194, 152, 211, 67, 207, 226, 229, 112, _
    156, 167, 97, 75, 22, 138, 251, 64, 240, 184, 221, 48, 106, 94, 93, 255, _
    39, 159, 42, 165, 217, 254, 194, 44, 196, 235, 8, 244, 200, 51, 148, 109, _
    110, 176, 100, 111, 60, 71, 139, 162, 36, 107, 3, 70, 134, 188, 28, 203, _
    42, 16, 23, 164, 178, 68, 91, 78, 46, 171, 155, 156, 194, 32, 174, 2, 120, _
    152, 8, 178, 113, 191, 174, 170, 237, 222, 242, 66, 83, 16, 215, 15, 191)
  iND = SetAsByte(ArrVar)
  ArrVar = Array(129, 199, 148, 63, 214, 230, 231, 149, 222, 170, 8, 34, 177, _
    14, 110, 27, 53, 47, 163, 11, 241, 234, 240, 243, 234, 67, 179, 217, 254, _
    130, 112, 64, 19, 238, 96, 33, 164, 135, 28, 248, 212, 42, 96, 19, 198, _
    94, 80, 173, 66, 128, 103, 138, 24, 20, 169, 211, 240, 44, 138, 253, 100, _
    195, 116, 41, 194, 92, 44, 106, 253, 224, 44, 120, 80, 25, 2, 79, 112, _
    116, 143, 132, 121, 24, 52, 112, 168, 216, 193, 202, 216, 195, 98, 199, _
    201, 15, 178, 13, 141, 61, 211, 74, 227, 104, 166, 202, 219, 82, 100, 179, _
    122, 45, 90, 18, 126, 195, 179, 164, 28, 40, 23, 187, 12, 59, 239, 92, _
    19, 35, 143, 204, 48, 213, 248, 52, 39, 130, 140, 114, 8, 63, 40, 161, _
    85, 217, 207, 237, 182, , 8, 221, 232, 237, 6, 208, 106, 137, 11, 29, 156, _
    236, 51, 34, 118, 210, 207, 160, 181, 150, 164, 122, 231, 125, 236, 140, _
    82, 20, 219, 182, 194, 22, 145, 40, 43, 5, 143, 206, 108, 45, 211, 171, _
    189, 138, 103, 110, 110, 45, 124, 211, 174, 3, 89, 25, 73, 26, 148, 120, _
    24, 186, 226, 22, 244, 229, 8, 109, 180, 231, 167, 164, 64, 155, 79, 70, _
    30, 71, 243, 236, 106, 74, 206, 116, 254, 73, 198, 16, 63, 153, 67, 202, _
    131, 247, 24, 194, 184, 41, 210, 197, 93, 221, 183, 92, 65, 67, 159, 112, _
    255, 145, 252, 104, 25, 122, 6, 238, 142, 156, 117, 56, 40, 113, 18, 241, _
    38, 247, 118, 64, 134, 96, 133, 68, 34, 163, 130, 185, 216, 163, 158, 67)
  iND = SetAsByte(ArrVar)
  ArrVar = Array(4, 32, 182, 158, 135, 15, 178, 51, 112, 165, 93, 121, 138, _
    144, 91, 69, 217, 191, 115, 181, 189, 181, 50, 105, 170, 117, 234, 148, _
    100, 246, 245, 237, 37, 24, 128, 51, 186, 148, 140, 249, 187, 223, 238, _
    141, 230, 237, 206, 11, 84, 176, 24, 32, 18, 71, 166, 5, 103, 204, 30, _
    79, 90, 182, 251, 99, 107, 214, 12, 184, 137, 212, 147, 51, 77, 35, 27, _
    86, 41, 17, 19, 224, 19, 61, 210, 1, 117, 33, 86, 70, 31, 186, 41, 201, _
    212, 207, 201, 146, 62, 77, 194, 7, 10, 196, 178, 4, 165, 87, 154, 23, _
    154, 145, 173, 102, 34, 197, 203, 222, 104, 206, 212, 146, 151, 252, 40, _
    108, 233, 101, 177, 28, 151, 198, 141, 177, 78, 138, 217, 7, 226, 178, _
    126, 5, 94, 206, 25, 150, 106, 172, 56, 180, 140, 244, 70, 222, 252, 252, _
    196, 189, 224, 19, 28, 217, 112, 136, 230, 8, 180, 223, 207, 56, 105, 106, _
    17, 145, 73, 39, 70, 110, 121, 81, 18, 219, 25, 227, 195, 19, 110, 69, _
    45, 7, 50, 234, 150, 16, 107, 92, 149, 179, 8, 242, 154, 95, 224, 78, 86, _
    161, 53, 123, 116, 220, 153, 95, 86, 56, 88, 156, 91, 182, 43, 144, 167, _
    71, 197, 13, 114, 43, 198, 136, 145, 39, 34, 234, 136, 246, 46, 251, 119, _
    4, 233, 46, 73, 54, 248, 192, 220, 245, 107, 240, 199, 146, 223, 248, 178, _
    141, 112, 83, 173, 246, 151, 19, 208, 110, 30, 188, 246, 154, , 93, 136, _
    , 166, 165, 89, 201, 223, 38, 224, 117, 147, 12, 195, 27, 168, 159, 31)
  iND = SetAsByte(ArrVar)
  ArrVar = Array(15, 72, 86, 43, 62, 95, 200, 29, 191, 42, 101, 150, 134, _
    242, 222, 13, 50, 203, 1, 58, 255, 250, 242, 229, 55, 94, 144, 224, 9, _
    115, 68, 116, 63, 64, 105, 89, 82, 138, 8, 201, 145, 138, 162, 244, 159, _
    200, 150, 82, 147, 234, 88, 39, 57, 189, 220, 68, 212, 54, 125, 66, 3, _
    102, 220, 214, 195, 44, 100, 81, 126, 248, 96, 106, 89, 159, 207, 61, 248, _
    40, 31, 5, 65, 112, 131, 44, 38, 141, 155, 200, 40, 157, 21, 22, 109, 88, _
    174, 175, 209, 191, 214, 104, 127, 103, 19, 129, 105, 223, 181, 237, 62, _
    212, 39, 141, 149, 39, 248, 194, 50, 75, 53, 61, 13, 90, 140, 201, 5, 97, _
    254, 131, 173, 179, 90, 34, 40, 200, 131, 236, 42, 205, 109, 55, 99, 107, _
    128, 168, 119, 158, 145, 116, 252, 28, 127, 176, 224, 38, 1, 176, 169, _
    19, 42, 143, 9, 30, 11, 224, 127, 255, 124, 253, 103, 164, 235, 189, 156, _
    131, 169, 220, 19, 232, 236, 233, 180, 7, 40, 230, 38, 101, 160, 35, 174, _
    128, 186, 79, 147, 55, 229, 159, 13, 74, 70, 115, 129, 155, 213, 25, 245, _
    55, 165, 252, 241, 124, 175, 217, 186, 94, 51, 169, 42, 158, 101, 3, 193, _
    202, 233, 18, 146, 172, 93, 84, 130, 70, 123, 160, 188, 122, 249, 29, 161, _
    193, 169, 105, 157, 70, 196, 136, 179, 6, 24, 150, 134, 199, 85, 58, 150, _
    205, 192, 122, 16, 114, 91, 122, 173, 90, 227, 37, 47, 150, 179, 20, 40, _
    137, 57, 27, 229, 190, 106, 223, 55, 105, 115, 147, 252, 94, 46)
  iND = SetAsByte(ArrVar)
End Sub
