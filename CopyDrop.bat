IF "%PROCESSOR_ARCHITECTURE%" == "x86" (
	xcopy /y "%~dp0\ExcelExtension\bin\Debug\Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelExtension.*" "%CommonProgramFiles%\Microsoft Shared\VSTT\10.0\UITestExtensionPackages\*.*"
	xcopy /y "%~dp0\ExcelExtension\bin\Debug\Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelCommunication.*" "%ProgramFiles%\Microsoft Visual Studio 10.0\Common7\IDE\PrivateAssemblies\*.*"
) ELSE (
	xcopy /y "%~dp0\ExcelExtension\bin\Debug\Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelExtension.*" "%CommonProgramFiles(x86)%\Microsoft Shared\VSTT\10.0\UITestExtensionPackages\*.*"
	xcopy /y "%~dp0\ExcelExtension\bin\Debug\Microsoft.VisualStudio.TestTools.UITest.Sample.ExcelCommunication.*" "%ProgramFiles(x86)%\Microsoft Visual Studio 10.0\Common7\IDE\PrivateAssemblies\*.*"
)


