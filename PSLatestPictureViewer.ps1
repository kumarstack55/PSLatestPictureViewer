Add-Type -AssemblyName System.Windows.Forms

# 定数
$InitialFormWidth = 1024
$InitialFormHeight = 768
$MiniThumbnailWidth = 200
$MiniThumbnailHeight = 100
$IntervalInMilliseconds = 2000
$NumberOfPicturesToDisplay = 20
$LabelTextNoInfo = "(no info)"

# 変数
[Pictures]$global:Pictures = $null
[Picture]$global:LatestPicture = $null
$global:PictureBoxList = [System.Collections.Generic.List[object]]::new()
$global:LabelList = [System.Collections.Generic.List[object]]::new()

# Drawing.Image の処理を関数で定義する。
# PowerShell のクラス内で Add-Type した型を利用できないため。
# https://stackoverflow.com/questions/34625440
function New-ImageFromMemoryStream {
    param([Parameter(Mandatory)]$MemoryStream)
    [System.Drawing.Image]::FromStream($MemoryStream)
}

# ピクチャ
class Picture {
    [string]$FilePath
    [string]$DirectoryPath
    [string]$FileName
    hidden $Image
    Picture([string]$FilePath, [string]$DirectoryPath, [string]$FileName, $Image) {
        $this.FilePath = $FilePath
        $this.DirectoryPath = $DirectoryPath
        $this.FileName = $FileName
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

        $ImageStream = [System.IO.File]::OpenRead($Item.FullName)
        $MemoryStream = New-Object System.IO.MemoryStream
        $ImageStream.CopyTo($MemoryStream)
        $ImageStream.Close()
        $MemoryStream.Position = 0
        $Image = New-ImageFromMemoryStream -MemoryStream $MemoryStream

        return [Picture]::new($FilePath, $DirectoryPath, $FileName, $Image)
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
            $this.OrderList.Remove($Key)
        } elseif ($this.Cache.Count -ge $this.Capacity) {
            $Key = $this.OrderList[0]
            $Value = $this.Cache[$Key]
            Invoke-Command -ScriptBlock $this.RemoveScriptBlock -ArgumentList $Key, $Value

            $this.OrderList.RemoveAt(0)
            $this.Cache.Remove($Key)
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
        if (-not $this.LruCache.ContainsKey($FilePath)) {
            $Picture = [PictureFactory]::CreatePicture($Item)
            $this.LruCache.Update($FilePath, $Picture)
        }
        return $this.LruCache.Get($FilePath)
    }
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
        $SortedItems[0 .. ($Count - 1)]
    }
}

function New-ViewerForm {
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Latest Images in Pictures"
    $Form.Width = $InitialFormWidth
    $Form.Height = $InitialFormHeight

    # Fill を最初に Add() する。
    # Top や Bottom より後に Fill を Add() すると、Panel 内の PictureBox の上側、下側が欠損してしまう。これを回避する。
    $Form.Controls.Add(
        (&{
            $c = New-Object System.Windows.Forms.Panel
            $c.Dock = [System.Windows.Forms.DockStyle]::Fill
            $c.Controls.Add(
                (&{
                    $c = New-Object System.Windows.Forms.PictureBox
                    $global:PictureBoxList.Add($c)
                    $c.Dock = [System.Windows.Forms.DockStyle]::Fill
                    $c.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
                    $c
                })
            )
            $c.Controls.Add(
                (&{
                    $c = New-Object System.Windows.Forms.Label
                    $global:LabelList.Add($c)
                    $c.Dock = [System.Windows.Forms.DockStyle]::Bottom
                    $c.AutoSize = $false
                    $c.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
                    $c.Text = $LabelTextNoInfo
                    $c
                })
            )
            $c
        })
    )

    $Form.Controls.Add(
        (&{
            $c = New-Object System.Windows.Forms.FlowLayoutPanel
            $c.Dock = [System.Windows.Forms.DockStyle]::Top
            $c.AutoSize = $true
            $c.Controls.Add(
                (&{
                    $c = New-Object System.Windows.Forms.Button
                    $c.Text = "画像をごみ箱に入れる"
                    $c.AutoSize = $true
                    $c.Add_Click({
                        if ($null -ne $global:LatestPicture) {
                            $Shell = New-Object -ComObject Shell.Application
                            $Folder = $Shell.Namespace($global:LatestPicture.DirectoryPath)
                            $Item = $Folder.ParseName($global:LatestPicture.FileName)
                            $Item.InvokeVerb("delete")
                        }
                    })
                    $c
                })
            )
            $c
        })
    )

    $Form.Controls.Add(
        (&{
            $c = New-Object System.Windows.Forms.FlowLayoutPanel
            $c.Dock = [System.Windows.Forms.DockStyle]::Bottom

            foreach ($_ in 2 .. $NumberOfPicturesToDisplay) {
                $c.Controls.Add(
                    (&{
                        $c = New-Object System.Windows.Forms.Panel
                        $c.Width = $MiniThumbnailWidth
                        $c.Height = $MiniThumbnailHeight
                        $c.Controls.Add(
                            (&{
                                $c = New-Object System.Windows.Forms.PictureBox
                                $c.Dock = [System.Windows.Forms.DockStyle]::Fill
                                $c.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
                                $global:PictureBoxList.Add($c)
                                $c
                            })
                        )
                        $c.Controls.Add(
                            (&{
                                $c = New-Object System.Windows.Forms.Label
                                $global:LabelList.Add($c)
                                $c.Dock = [System.Windows.Forms.DockStyle]::Bottom
                                $c.AutoSize = $false
                                $c.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
                                $c.Text = $LabelTextNoInfo
                                $c
                            })
                        )
                        $c
                    })
                )
            }
            $c
        })
    )

    $Form
}

function Invoke-Tick {
    # 最近の画像を最大 $NumberOfPicturesToDisplay 件、得る。
    $Items = Get-LatestItemsInPictures -Count $NumberOfPicturesToDisplay

    # 画像、ラベルを設定する。
    foreach ($Index in 0 .. ($NumberOfPicturesToDisplay - 1)) {
        $PictureBox = $PictureBoxList[$Index]
        $Label = $LabelList[$Index]

        if ($Index -lt $Items.Count) {
            $Item = $Items[$Index]
            $Picture = $global:Pictures.GetPicture($Item)
            if ($Index -eq 0) {
                $global:LatestPicture = $Picture
            }

            try {
                $PictureBox.Image = $Picture.Image
            } catch {
                Write-Host $_.ScriptStackTrace
            }

            $LabelText = "{0}: {1}" -f ($Index + 1), $Picture.FileName
            $Label.Text = $LabelText
        } else {
            if ($Index -eq 0) {
                $global:LatestPicture = $null
            }
            $PictureBox.Image = $null
            $Label.Text = ""
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
    $global:Pictures = [Pictures]::new($CacheSize)

    Write-Host "Current PowerShell process ID: ${PID}"

    $Form = New-ViewerForm
    $Timer = New-ViewerTimer

    $Timer.Start()

    # フォームを表示する前に、画像を更新しておく。
    Invoke-Tick

    [System.Windows.Forms.Application]::Run($Form)
}

Invoke-Application
