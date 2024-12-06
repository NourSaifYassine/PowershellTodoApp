param (
    [switch]$AddTodo,
    [switch]$ViewTasks,
    [switch]$RemoveTask,
    [switch]$FindWordFile,
    [switch]$AddTodoToWord
)

$todoListFile = "$env:USERPROFILE\Documents\TodoList.txt"

if (-not (Test-Path -Path $todoListFile)) {
    New-Item -Path $todoListFile -ItemType File
}

function Add-Todo {
    $task = Read-Host "Enter a Task"
    Add-Content -Path $todoListFile -Value $task
    Write-Host "Your task has been added: $task"
}

function Show-Tasks {
    if (Test-Path -Path $todoListFile) {
        $tasks = Get-Content -Path $todoListFile
        if ($tasks.Count -eq 0) {
            Write-Host "No tasks found."
        }
        else {
            $tasks | ForEach-Object { Write-Host $_ }
        }
    }
    else {
        Write-Host "No tasks found."
    }
}

function Remove-Task {
    if (Test-Path -Path $todoListFile) {
        $tasks = Get-Content -Path $todoListFile
        if ($tasks.Count -eq 0) {
            Write-Host "No tasks to remove."
            return
        }
        
        for ($i = 0; $i -lt $tasks.Count; $i++) {
            Write-Host "$i. $($tasks[$i])"
        }

        $taskIndex = Read-Host "Enter the number of the task to remove"
        if ($taskIndex -ge 0 -and $taskIndex -lt $tasks.Count) {
            $tasks = $tasks | Where-Object { $_ -ne $tasks[$taskIndex] }
            Set-Content -Path $todoListFile -Value $tasks
            Write-Host "Task removed."
            Start-Process explorer "$todoListFile"
            View-Tasks
        }
        else {
            Write-Host "Invalid task number."
        }
    }
    else {
        Write-Host "No tasks to remove."
    }
}

function Find-WordFile {
    $wordInput = Read-Host "Enter the name of the Word file (without the .docx extension)"

    $directories = Get-ChildItem -Path $env:USERPROFILE -Recurse -Directory -ErrorAction SilentlyContinue
    
    $wordFiles = @()
    
    foreach ($dir in $directories) {
        $foundFiles = Get-ChildItem -Path $dir.FullName -Filter "$wordInput.docx" -ErrorAction SilentlyContinue
        $wordFiles += $foundFiles
    }
    
    if ($wordFiles.Count -gt 0) {
        Write-Host "Found the following Word files:"
        
        for ($i = 0; $i -lt $wordFiles.Count; $i++) {
            Write-Host "$i. $($wordFiles[$i].FullName)"
        }
    
        $selectedIndex = Read-Host "Enter the number of the file you want to open"
    
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $wordFiles.Count) {
            $selectedFile = $wordFiles[$selectedIndex]
            Write-Host "You selected: $($selectedFile.FullName)"
            Invoke-Item $selectedFile.FullName
        } else {
            Write-Host "Invalid selection. Please try again."
        }
    } else {
        Write-Host "No Word files matching '$wordInput.docx' were found."
    }    
}


function Add-TodoToWord {
    $wordInput = Read-Host "Enter the name of the Word file (without the .docx extension)"
    $wordFilePath = (Get-ChildItem -Recurse -Filter "$wordInput.docx" -ErrorAction SilentlyContinue).FullName

    if (-not $wordFilePath) {
        Write-Host "Word file not found. Creating a new file: $wordInput.docx"
        $wordFilePath = "$($pwd.Path)\$wordInput.docx"
    }

    $task = Read-Host "Enter the task to add to the Word document"

    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $true

    if (Test-Path $wordFilePath) {
        $doc = $wordApp.Documents.Open($wordFilePath)
    }
    else {
        $doc = $wordApp.Documents.Add()
        $doc.SaveAs([ref]$wordFilePath)
    }

    $range = $doc.Content
    $range.InsertAfter([char]0x2022 + " $task`r`n")

    $doc.Save()
    Write-Host "Task added to the Word document: $task"

    $doc.Close()
    $wordApp.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp) | Out-Null
}


if ($AddTodo) {
    Add-Todo
}
elseif ($ViewTasks) {
    Show-Tasks
}
elseif ($RemoveTask) {
    Remove-Task
}
elseif ($FindWordFile) {
    Find-WordFile
}
elseif ($AddTodoToWord) {
    Add-TodoToWord
}
else {
    Write-Host "Please provide a valid action: -AddTodo, -ViewTasks, -RemoveTask, -FindWord, -FindWordFile, or -AddTodoToWord"
}
