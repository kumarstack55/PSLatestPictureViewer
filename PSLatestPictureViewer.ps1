Add-Type -AssemblyName System.Windows.Forms

# 定数
$InitialFormWidth = 1024
$InitialFormHeight = 768
$MiniThumbnailWidth = 200
$MiniThumbnailHeight = 100
$IntervalInMilliseconds = 5000
$NumberOfPicturesToDisplay = 1+20
$LabelTextNoInfo = "(no info)"

# 変数
[Pictures]$global:Pictures = $null
$global:CustomRecords = $null
$global:Form = $null
$global:LastPositionX = $null
$global:LastPositionY = $null
$global:LastWidth = $null
$global:LastHeight = $null

# Drawing.Image の処理を関数で定義する。
# PowerShell のクラス内で Add-Type した型を利用できないため。
# https://stackoverflow.com/questions/34625440
function New-ImageFromMemoryStream {
    param([Parameter(Mandatory)]$MemoryStream)
    [System.Drawing.Image]::FromStream($MemoryStream)
}

function Write-CustomHost {
    param([Parameter(Mandatory)][string]$Message)
    Write-Host ("PID:{0} {1}" -f $PID, $Message)
}

# ピクチャ
class Picture {
    [string]$FilePath
    [string]$DirectoryPath
    [string]$FileName
    [System.DateTime]$LastWriteTime
    hidden $Image
    Picture([string]$FilePath, [string]$DirectoryPath, [string]$FileName, [System.DateTime]$LastWriteTime, $Image) {
        $this.FilePath = $FilePath
        $this.DirectoryPath = $DirectoryPath
        $this.FileName = $FileName
        $this.LastWriteTime = $LastWriteTime
        $this.Image = $Image
    }
    Dispose() {
        if ($null -ne $this.Image) {
            $this.Image.Dispose()
            $this.Image = $null
        }
    }
}

# ピクチャ・ファクトリ
class PictureFactory {
    static [Picture]CreatePicture([System.IO.FileInfo]$Item) {
        $FilePath = $Item.FullName
        $DirectoryPath = $Item.Directory.FullName
        $FileName = $Item.Name
        $LastWriteTime = $Item.LastWriteTime

        $ImageStream = [System.IO.File]::OpenRead($Item.FullName)
        $MemoryStream = New-Object System.IO.MemoryStream
        $ImageStream.CopyTo($MemoryStream)
        $ImageStream.Close()
        $MemoryStream.Position = 0
        $Image = New-ImageFromMemoryStream -MemoryStream $MemoryStream

        return [Picture]::new($FilePath, $DirectoryPath, $FileName, $LastWriteTime, $Image)
    }
}

# LRUキャッシュ
class LruCache {
    hidden [int]$Capacity
    hidden [scriptblock]$RemoveScriptBlock
    hidden [hashtable]$Cache = @{}
    hidden [System.Collections.Generic.List[string]]$OrderList = [System.Collections.Generic.List[string]]::new()
    LruCache([int]$Capacity, [scriptblock]$RemoveScriptBlock) {
        $this.Capacity = $Capacity
        $this.RemoveScriptBlock = $RemoveScriptBlock
    }
    Update([string]$Key, [Picture]$Value) {
        if ($this.Cache.ContainsKey($Key)) {
            $OldValue = $this.Cache[$Key]
            Invoke-Command -ScriptBlock $this.RemoveScriptBlock -ArgumentList $Key, $OldValue
            $this.OrderList.Remove($Key)
        } elseif ($this.Cache.Count -ge $this.Capacity) {
            $ExpiredKey = $this.OrderList[0]
            $ExpiredValue = $this.Cache[$ExpiredKey]
            Invoke-Command -ScriptBlock $this.RemoveScriptBlock -ArgumentList $ExpiredKey, $ExpiredValue
            $this.OrderList.Remove($ExpiredKey)
            $this.Cache.Remove($ExpiredKey)
        }
        $this.Cache[$Key] = $Value
        $this.OrderList.Add($Key)
    }
    [bool] ContainsKey([string]$Key) {
        return $this.Cache.ContainsKey($Key)
    }
    [Picture] Get([string]$Key) {
        return $this.Cache[$Key]
    }
    RemoveLatest() {
        $Key = $this.OrderList[$this.OrderList.Count - 1]
        $Value = $this.Cache[$Key]
        Invoke-Command -ScriptBlock $this.RemoveScriptBlock -ArgumentList $Key, $Value
        $this.OrderList.Remove($Key)
        $this.Cache.Remove($Key)
    }
    RemoveAll() {
        while ($this.OrderList.Count -gt 0) {
            $this.RemoveLatest()
        }
    }
}

# 複数のピクチャ
class Pictures {
    hidden [LruCache]$LruCache
    Pictures([int]$CacheCapacity) {
        $RemoveScriptBlock = {
            param([string]$Key, [Picture]$Value)
            $Value.Image.Dispose()
        }
        $this.LruCache = [LruCache]::new($CacheCapacity, $RemoveScriptBlock)
    }
    [Picture] GetPicture([System.IO.FileInfo]$Item) {
        $FilePath = $Item.FullName
        if ($this.LruCache.ContainsKey($FilePath)) {
            $Picture = $this.LruCache.Get($FilePath)
            if ($Item.LastWriteTime -gt $Picture.LastWriteTime) {
                $Picture = [PictureFactory]::CreatePicture($Item)
                $this.LruCache.Update($FilePath, $Picture)
            }
        } else {
            $Picture = [PictureFactory]::CreatePicture($Item)
            $this.LruCache.Update($FilePath, $Picture)
        }
        return $this.LruCache.Get($FilePath)
    }
    RemoveLatestItemFromCache() {
        $this.LruCache.RemoveLatest()
    }
    RemoveAllItemsFromCache() {
        $this.LruCache.RemoveAll()
    }
}

# 画面を構成するピクチャを含む各種オブジェクトのレコード
class CustomRecord {
    [Picture]$Picture
    $Panel
    $Label
    $PictureBox
    $Button
    $Block
}

function Get-MyPicturesFolderPath {
    param()
    [Environment]::GetFolderPath("MyPictures")
}

function Get-LatestItemsInPictures {
    param([int]$Count = -1)

    $PicturesFolderPath = Get-MyPicturesFolderPath
    $FileItems = Get-ChildItem -LiteralPath $PicturesFolderPath -File
    $ImageItems = $FileItems | Where-Object { $_ -match '^.*\.(jpg|png|bmp)$' }
    $SortedItems = $ImageItems | Sort-Object -Property LastWriteTime -Descending
    if ($Count -eq -1) {
        $SortedItems
    } else {
        $SortedItems | Select-Object -First $Count
    }
}

function New-CustomPanelControl {
    param([Parameter(Mandatory)][CustomRecord]$CustomRecord)

    $CustomRecord.Panel = New-Object System.Windows.Forms.Panel
    $CustomRecord.PictureBox = New-Object System.Windows.Forms.PictureBox
    $CustomRecord.Label = New-Object System.Windows.Forms.Label
    $CustomRecord.Button = New-Object System.Windows.Forms.Button

    $c = $CustomRecord.Panel
    $c.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $c.Controls.Add(
        (&{
            $c = New-Object System.Windows.Forms.Panel
            $c.Dock = [System.Windows.Forms.DockStyle]::Fill
            $c.AutoSize = $true
            $c.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
            $c.Controls.Add(
                (&{
                    $c = $CustomRecord.PictureBox
                    $c.Dock = [System.Windows.Forms.DockStyle]::Fill
                    $c.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
                    $c
                })
            )
            $c.Controls.Add(
                (&{
                    $c = $CustomRecord.Label
                    $c.Dock = [System.Windows.Forms.DockStyle]::Top
                    $c.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
                    $c.BackColor = [System.Drawing.Color]::AliceBlue
                    $c.Text = $LabelTextNoInfo
                    $c
                })
            )
            $c.Controls.Add(
                (&{
                    $c = $CustomRecord.Button
                    $c.Text = "画像をごみ箱に入れる"
                    $c.Dock = [System.Windows.Forms.DockStyle]::Bottom
                    $c.Enabled = $false
                    $c.Add_Click($CustomRecord.Block)
                    $c
                })
            )
            $c
        })
    )
    $c
}

function New-ViewerForm {
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Latest Images in Pictures"
    if ($null -ne $global:LastPositionX) {
        $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
        $Form.Left = $global:LastPositionX
        $Form.Top = $global:LastPositionY
    }

    if ($null -ne $global:LastWidth) {
        $Form.Width = $global:LastWidth
        $Form.Height = $global:LastHeight
    } else {
        $Form.Width = $InitialFormWidth
        $Form.Height = $InitialFormHeight
    }

    $CustomRecord = $global:CustomRecords[0]
    $Form.Controls.Add( (&{
            $c = New-CustomPanelControl -CustomRecord $CustomRecord
            $c.Dock = [System.Windows.Forms.DockStyle]::Fill
            $c
        })
    )
    $Form.Controls.Add(
        (&{
            $c = New-Object System.Windows.Forms.FlowLayoutPanel
            $c.Dock = [System.Windows.Forms.DockStyle]::Bottom
            $c.AutoSize = $true
            $c.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
            $c.WrapContents = $false
            foreach ($Index in 1 .. ($NumberOfPicturesToDisplay - 1)) {
                $CustomRecord = $global:CustomRecords[$Index]
                $c.Controls.Add(
                    (&{
                        $c = New-CustomPanelControl -CustomRecord $CustomRecord
                        $c.Width = $MiniThumbnailWidth
                        $c.Height = $MiniThumbnailHeight
                        $c
                    })
                )
            }
            $c
        })
    )

    $Form.Add_Move({
        $Position = $Form.Location
        Write-CustomHost "Window moved to X: $($Position.X), Y: $($Position.Y)"
        $global:LastPositionX = $Position.X
        $global:LastPositionY = $Position.Y
    })

    $Form.Add_Resize({
        $Size = $Form.Size
        Write-CustomHost "Window resized to Width: $($Size.Width), Height: $($Size.Height)"
        $global:LastWidth = $Size.Width
        $global:LastHeight = $Size.Height
    })

    $Form
}

function Update-AllCustomPanels {
    # 最近の画像を最大 $NumberOfPicturesToDisplay 件、得る。
    $Items = Get-LatestItemsInPictures -Count $NumberOfPicturesToDisplay

    # 画像、ラベルを設定する。
    foreach ($Index in 0 .. ($NumberOfPicturesToDisplay - 1)) {
        $CustomRecord = $global:CustomRecords[$Index]
        $PictureBox = $CustomRecord.PictureBox
        $Label = $CustomRecord.Label
        $Button = $CustomRecord.Button
        if ($Index -lt $Items.Count) {
            $Item = $Items[$Index]
            $Picture = $global:Pictures.GetPicture($Item)
            $CustomRecord.Picture = $Picture

            $LabelText = "{0}: {1}" -f ($Index + 1), $Picture.FileName
            $Label.Text = $LabelText

            # 期待した動作とならない場合は Form を作り直す。
            try {
                $PictureBox.Image = $Picture.Image
                $Button.Enabled = $true
            } catch {
                Write-Host $_.ScriptStackTrace
                $Button.Enabled = $false
                $PictureBox.Image = $null
                $global:Pictures.RemoveLatestItemFromCache()
            }
        } else {
            $PictureBox.Image = $null
            $Label.Text = ""
            $Button.Enabled = $false
        }
    }
}

function New-ViewerTimer {
    $Timer = New-Object System.Windows.Forms.Timer
    $Timer.Interval = $IntervalInMilliseconds
    $Timer.Add_Tick({
        Write-CustomHost "Interval Timer: Calling Update-AllCustomPanels..."
        Update-AllCustomPanels
    })
    $Timer
}

function Invoke-Application {
    $global:CustomRecords = [System.Collections.Generic.List[CustomRecord]]::new()

    # 表示する画像数より十分に大きい数をキャッシュ対象とする。
    $CacheSize = [System.Math]::Ceiling($NumberOfPicturesToDisplay * 1.5)
    if ($null -ne $global:Pictures) {
        $global:Pictures.RemoveAllItemsFromCache()
    }
    $global:Pictures = [Pictures]::new($CacheSize)

    foreach ($Index in 0 .. ($NumberOfPicturesToDisplay - 1)) {
        $Block = {
            Write-CustomHost "block is invoked. (Index:$Index)"
            $Shell = New-Object -ComObject Shell.Application
            $CustomRecord = $global:CustomRecords[$Index]

            $Picture = $CustomRecord.Picture
            if ($null -eq $Picture) {
                Write-CustomHost "Picture is null."
                return
            }

            $Folder = $Shell.Namespace($Picture.DirectoryPath)
            if ($null -eq $Folder) {
                Write-CustomHost "Folder is null."
                return
            }

            $Item = $Folder.ParseName($Picture.FileName)
            if ($null -eq $Item) {
                Write-CustomHost "Item is null."
                return
            }

            $Item.InvokeVerb("delete")
        }
        $Closure = $Block.GetNewClosure()

        $CustomRecord = [CustomRecord]::new()
        $CustomRecord.Picture = $null
        $CustomRecord.Label = $null
        $CustomRecord.Block = $Closure
        $global:CustomRecords.Add($CustomRecord)
    }

    $global:Form = New-ViewerForm

    # フォームを表示する前に、画像を更新しておく。
    Update-AllCustomPanels

    [System.Windows.Forms.Application]::Run($Form)
}

Write-Host "Current PowerShell process ID: ${PID}"

# FileSystemWatcher で指定ディレクトリの変更を検知する。
$PathToWatch = Get-MyPicturesFolderPath
$Watcher = New-Object System.IO.FileSystemWatcher
$Watcher.Path = $PathToWatch
$Watcher.EnableRaisingEvents = $true
Register-ObjectEvent $Watcher Created -SourceIdentifier FileCreated -Action {
    Write-CustomHost "Action: Created: Calling Update-AllCustomPanels..."
    Update-AllCustomPanels
} | Out-Null
Register-ObjectEvent $Watcher Changed -SourceIdentifier FileChanged -Action {
    Write-CustomHost "Action: Changed: Calling Update-AllCustomPanels..."
    Update-AllCustomPanels
} | Out-Null
Register-ObjectEvent $Watcher Deleted -SourceIdentifier FileDeleted -Action {
    Write-CustomHost "Action: Deleted: Calling Update-AllCustomPanels..."
    Update-AllCustomPanels
} | Out-Null
Register-ObjectEvent $Watcher Renamed -SourceIdentifier FileRenamed -Action {
    Write-CustomHost "Action: Renamed: Calling Update-AllCustomPanels..."
    Update-AllCustomPanels
} | Out-Null

# FileSystemWatcher に加えて、タイマーで定期的に変更を検知する。
# 別のプロセスがピクチャ・ファイルに変更を加えるとき、FileSystemWatcher がイベントを通知するタイミングが、
# ファイルを書き込み中か、ファイルの書き込み後のクローズ処理か、仕様から判断できなかったため。
# 実装として、少なくとも1回のピクチャ・ファイルの保存で、複数回の Changed が通知されたことがある。
$Timer = New-ViewerTimer
$Timer.Start()

Write-CustomHost "Starting Application..."
Invoke-Application
Write-CustomHost "Terminated."

[System.Windows.Forms.Application]::Exit()
