@echo ��ʼע�� --����Ҫ�Ҽ��Թ���Ա������Ŷ���ұ�����MSHFLXGD.ocx�ļ���һ��
Rd "%WinDir%\system32\test_permissions" >NUL 2>NUL
Md "%WinDir%\System32\test_permissions" 2>NUL||(Echo ��⵽��δ�ù���Ա��ݣ���ʹ���Ҽ�����Ա������У������޷�����ע�������&&Pause >nul&&Exit)
Rd "%WinDir%\System32\test_permissions" 2>NUL
%~d0
cd %~dp0 
if exist %windir%\SysWOW64 (
	copy wow64ext.dll %windir%\syswow64\
	%windir%\syswow64\regsvr32 %windir%\syswow64\wow64ext.dll
)else (
	copy wow64ext.dll %windir%\system32\ 
	%windir%\system32\regsvr32 %windir%\system32\wow64ext.dll
)
@echo ע����ɡ���CH
@pause 