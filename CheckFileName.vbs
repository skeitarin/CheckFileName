Option Explicit
 
Dim objFileSys
Dim objFolder
Dim objFile
Dim objOutputTextStream
 
'�t�@�C���V�X�e���������I�u�W�F�N�g���쐬
Set objFileSys = CreateObject("Scripting.FileSystemObject")
 
'���O�o�͗p TextStream �I�u�W�F�N�g���쐬
'��2������ 1 �F�ǂݎ��A2 �F�㏑���A3 �F�ǋL�B
Dim logFileNm
Dim nowdate
nowdate= Year(Now())
nowdate= nowdate & Right("0" & Month(Now()) , 2)
nowdate= nowdate & Right("0" & Day(Now()) , 2)
nowdate= nowdate & Right("0" & Hour(Now()) , 2)
nowdate= nowdate & Right("0" & Minute(Now()) , 2)
nowdate= nowdate & Right("0" & Second(Now()) , 2)
logFileNm = "log_" & CStr(nowdate) & ".txt"
' logFileNm = "loga" & ".txt"
Set objOutputTextStream = objFileSys.OpenTextFile(logFileNm, 2, True)

Dim filePath
filePath = InputBox("��r�Ώۂ̃t�@�C���p�X�̓��́i�I�������t�@�C�����ɋL�ڂ̂Ȃ��t�@�C���������O�ɏo�͂��܂��j")
Dim objFSO, objFile2, line
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile2 = objFSO.OpenTextFile(filePath)

'�t�H���_�̃I�u�W�F�N�g���擾
Dim dirPath
dirPath = InputBox("�����Ώۂ̃t�H���_�p�X�̓���")
Set objFolder = objFileSys.GetFolder(dirPath)

Dim exist
 objOutputTextStream.WriteLine "��`�t�@�C���p�X�F" & filePath
 objOutputTextStream.WriteLine "�����Ώۃt�H���_�p�X�F" & dirPath
 objOutputTextStream.WriteLine "---- ���݂��Ȃ��t�@�C�����̃��X�g ----"

Do Until objFile2.AtEndOfStream
    exist = False
    line = objFile2.ReadLine
    ' objOutputTextStream.WriteLine "log�F" & line
    For Each objFile In objFolder.Files
        ' objOutputTextStream.WriteLine "_log�F" & objFile.Name
        if line = objFile.Name then
            exist = True
            Exit For
        end if
    Next
    if exist = False then
        '���O�t�@�C���ɏo��
        objOutputTextStream.WriteLine line
    end if
Loop
 
'TextStream �� Close ��Y�ꂸ��
objOutputTextStream.Close
 
Set objOutputTextStream = Nothing
Set objFolder  = Nothing
Set objFileSys = Nothing 

WScript.Echo "�������I�����܂����B���ʂ����O�t�@�C�����m�F���Ă��������B"