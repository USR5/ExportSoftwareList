# ExportSoftwareList
Windows10 Pro使用 PowerShell 导出计算机上安装的软件列表及其详细信息到桌面
大家好，欢迎来到本期的教程！今天我将向大家介绍如何使用 PowerShell 脚本来导出计算机上已安装的软件列表，并获取每个软件的详细信息。让我们开始吧！

步骤一：复制代码 首先，复制以下代码并将其保存为一个 PowerShell 脚本文件（例如：ExportSoftwareList.ps1）。


$softwareList = Get-ItemProperty $uninstallKey |

                Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |

                Where-Object {$_.DisplayName -ne $null}



$excel = New-Object -ComObject Excel.Application

$excel.Visible = $true



$workbook = $excel.Workbooks.Add()

$worksheet = $workbook.Worksheets.Item(1)



$worksheet.Cells.Item(1,1) = "软件名称"

$worksheet.Cells.Item(1,2) = "版本"

$worksheet.Cells.Item(1,3) = "发布者"

$worksheet.Cells.Item(1,4) = "安装日期"



$row = 2

foreach ($software in $softwareList) {

    $worksheet.Cells.Item($row,1) = $software.DisplayName

    $worksheet.Cells.Item($row,2) = $software.DisplayVersion

    $worksheet.Cells.Item($row,3) = $software.Publisher

    $worksheet.Cells.Item($row,4) = $software.InstallDate

    $row++

}



$savePath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'SoftwareList.xlsx')

$workbook.SaveAs($savePath)

$excel.Quit()


步骤二：运行脚本 现在，打开 PowerShell 终端，并导航到保存了脚本文件的位置。运行以下命令以执行脚本：


.\ExportSoftwareList.ps1


脚本将开始运行，并在计算机上搜索已安装的软件。它会获取每个软件的显示名称、版本号、发布者和安装日期。

步骤三：查看导出结果 脚本执行完成后，你可以在桌面上找到一个名为 "SoftwareList.xlsx" 的 Excel 文件。双击打开该文件，你将看到一个表格，其中包含了计算机上已安装软件的详细信息。

每一列显示了软件的不同属性：软件名称、版本号、发布者和安装日期。你可以根据自己的需求，对表格进行进一步的格式化和调整。

通过使用这个脚本，你可以轻松地获取计算机上已安装软件的详细信息，并将其保存到一个方便查看的 Excel 文件中。

这就是使用 PowerShell 导出计算机上已安装软件列表及其详细信息的方法了！如果你有任何问题或疑问，请随时向我提问。谢谢大家的阅读！

