Add-Type -AssemblyName System.Windows.Forms

# 定数
$InitialFormWidth = 1024
$InitialFormHeight = 768
$MiniThumbnailWidth = 200
$MiniThumbnailHeight = 100
$IntervalInMilliseconds = 2000
$NumberOfPicturesToDisplay = 1+20
$LabelTextNoInfo = "(no info)"

# 変数
[Pictures]$global:Pictures = $null
$global:CompornentRecords = [System.Collections.Generic.List[CompornentRecord]]::new()
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

class CompornentRecord {
    [Picture]$Picture
    $Index
    $Panel
    $Label
    $PictureBox
    $Button
    $Block
}

function Get-LatestItemsInPictures {
    param([int]$Count = -1)

    $PicturesFolderPath = [Environment]::GetFolderPath("MyPictures")
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
    param([Parameter(Mandatory)][CompornentRecord]$CompornentRecord)

    $CompornentRecord.Panel = New-Object System.Windows.Forms.Panel
    $CompornentRecord.PictureBox = New-Object System.Windows.Forms.PictureBox
    $CompornentRecord.Label = New-Object System.Windows.Forms.Label
    $CompornentRecord.Button = New-Object System.Windows.Forms.Button

    $c = $CompornentRecord.Panel
    $c.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $c.Controls.Add(
        (&{
            $c = New-Object System.Windows.Forms.Panel
            $c.Dock = [System.Windows.Forms.DockStyle]::Fill
            $c.AutoSize = $true
            $c.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
            $c.Controls.Add(
                (&{
                    $c = $CompornentRecord.PictureBox
                    $c.Dock = [System.Windows.Forms.DockStyle]::Fill
                    $c.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
                    $c
                })
            )
            $c.Controls.Add(
                (&{
                    $c = $CompornentRecord.Label
                    $c.Dock = [System.Windows.Forms.DockStyle]::Top
                    $c.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
                    $c.BackColor = [System.Drawing.Color]::AliceBlue
                    $c.Text = $LabelTextNoInfo
                    $c
                })
            )
            $c.Controls.Add(
                (&{
                    $c = $CompornentRecord.Button
                    $c.Text = "画像をごみ箱に入れる"
                    $c.Dock = [System.Windows.Forms.DockStyle]::Bottom
                    $c.Enabled = $false
                    $c.Add_Click($CompornentRecord.Block)
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

    $CompornentRecord = $global:CompornentRecords[0]
    $Form.Controls.Add( (&{
            $c = New-CustomPanelControl -CompornentRecord $CompornentRecord
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
                $CompornentRecord = $global:CompornentRecords[$Index]
                $c.Controls.Add(
                    (&{
                        $c = New-CustomPanelControl -CompornentRecord $CompornentRecord
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

function Invoke-Tick {
    Write-CustomHost "Invoke-Tick is called."

    # 最近の画像を最大 $NumberOfPicturesToDisplay 件、得る。
    $Items = Get-LatestItemsInPictures -Count $NumberOfPicturesToDisplay

    # 画像、ラベルを設定する。
    foreach ($Index in 0 .. ($NumberOfPicturesToDisplay - 1)) {
        $CompornentRecord = $global:CompornentRecords[$Index]
        $PictureBox = $CompornentRecord.PictureBox
        $Label = $CompornentRecord.Label
        $Button = $CompornentRecord.Button
        if ($Index -lt $Items.Count) {
            $Item = $Items[$Index]
            $Picture = $global:Pictures.GetPicture($Item)
            $CompornentRecord.Picture = $Picture

            $LabelText = "{0}: {1}" -f ($Index + 1), $Picture.FileName
            $Label.Text = $LabelText

            # 期待した動作とならない場合は Form を作り直す。
            # 特に 1 tick 内に複数の画像が追加あるいは削除されたときに例外が発生しそうだが原因はわかってない。
            # 出力中の未完全な画像や、削除中の画像を読んでいるため？
            try {
                $PictureBox.Image = $Picture.Image
            } catch {
                Write-Host $_.ScriptStackTrace

                $global:Pictures.RemoveLatestItemFromCache()

                Write-CustomHost "Set NeedRestarting is true."
                $global:NeedRestarting = $true

                $global:Form.Close()
            }

            $Button.Enabled = $true
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
    $Timer.Add_Tick({ Invoke-Tick })
    $Timer
}

function Invoke-Application {
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
            $CompornentRecord = $global:CompornentRecords[$Index]

            $Picture = $CompornentRecord.Picture
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

        $CompornentRecord = [CompornentRecord]::new()
        $CompornentRecord.Picture = $null
        $CompornentRecord.Label = $null
        $CompornentRecord.Index = $Index
        $CompornentRecord.Block = $Closure
        $global:CompornentRecords.Add($CompornentRecord)
    }

    $global:Form = New-ViewerForm

    # フォームを表示する前に、画像を更新しておく。
    Invoke-Tick

    [System.Windows.Forms.Application]::Run($Form)
}

Write-Host "Current PowerShell process ID: ${PID}"

$Timer = New-ViewerTimer
$Timer.Start()

$global:NeedRestarting = $false
Write-Host "Starting Application..."
Invoke-Application
while ($global:NeedRestarting) {
    Write-Host "Restarting Application..."
    $global:NeedRestarting = $false
    Invoke-Application
    Write-Host "Terminated."
    Start-Sleep 1
}

[System.Windows.Forms.Application]::Exit()
