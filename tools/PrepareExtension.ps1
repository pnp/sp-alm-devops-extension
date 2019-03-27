
$currentPath = (Split-Path -Parent $MyInvocation.MyCommand.Path)

$extensionFileJson = Get-Content -Path "$currentPath\..\src\vss-extension.json" | Out-String | ConvertFrom-Json

#copy only to used extension paths

$extensionTasks = $extensionFileJson.contributions | Where-Object { $_.type -eq "ms.vss-distributed-task.task"}
$extensionIds = $extensionTasks.properties.name

$extensionIds | ForEach-Object {

    $taskIdName = $_

    $destinationFolder = "$currentPath\..\src\$taskIdName\$taskIdNameV1\ps_modules"

    #remove any content from those folder, as they are temporary
    Remove-Item -Path $destinationFolder -Recurse -Force -ErrorAction SilentlyContinue
    #copy and overwrite all
    Copy-Item -Path "$currentPath\..\src\ps_modules" -Destination $destinationFolder -Recurse -Force


}