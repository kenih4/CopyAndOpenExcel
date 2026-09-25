Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --------------------------------------------------
# 初期値（デフォルト値）の自動取得
# --------------------------------------------------
$CurrentYear  = (Get-Date).ToString("yyyy")
$CurrentMonth = (Get-Date).ToString("MM")

# --------------------------------------------------
# 画面（フォーム）の作成
# --------------------------------------------------
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Log Note Copy Tool"
$Form.Size = New-Object System.Drawing.Size(350, 300) # チェックボックス用に少し縦を広げました
$Form.StartPosition = "CenterScreen"
$Form.FormBorderStyle = "FixedDialog"
$Form.MaximizeBox = $false

# 1. 種類（Type）のラベルとコンボボックス
$LabelType = New-Object System.Windows.Forms.Label
$LabelType.Text = "種類 (Type):"
$LabelType.Location = New-Object System.Drawing.Point(30, 30)
$LabelType.Size = New-Object System.Drawing.Size(100, 20)
$Form.Controls.Add($LabelType)

$ComboType = New-Object System.Windows.Forms.ComboBox
$ComboType.Location = New-Object System.Drawing.Point(140, 27)
$ComboType.Size = New-Object System.Drawing.Size(150, 20)
$ComboType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$ComboType.Items.Add("SACLA")
[void]$ComboType.Items.Add("SCSS")
[void]$ComboType.Items.Add("SP8")
$ComboType.SelectedItem = "SACLA"
$Form.Controls.Add($ComboType)

# 2. 年（Year）のラベルとテキストボックス
$LabelYear = New-Object System.Windows.Forms.Label
$LabelYear.Text = "西暦 (Year):"
$LabelYear.Location = New-Object System.Drawing.Point(30, 70)
$LabelYear.Size = New-Object System.Drawing.Size(100, 20)
$Form.Controls.Add($LabelYear)

$TxtYear = New-Object System.Windows.Forms.TextBox
$TxtYear.Location = New-Object System.Drawing.Point(140, 67)
$TxtYear.Size = New-Object System.Drawing.Size(150, 20)
$TxtYear.Text = $CurrentYear
$Form.Controls.Add($TxtYear)

# 3. 月（Month）のラベルとテキストボックス
$LabelMonth = New-Object System.Windows.Forms.Label
$LabelMonth.Text = "月 (Month):"
$LabelMonth.Location = New-Object System.Drawing.Point(30, 110)
$LabelMonth.Size = New-Object System.Drawing.Size(100, 20)
$Form.Controls.Add($LabelMonth)

$TxtMonth = New-Object System.Windows.Forms.TextBox
$TxtMonth.Location = New-Object System.Drawing.Point(140, 107)
$TxtMonth.Size = New-Object System.Drawing.Size(150, 20)
$TxtMonth.Text = $CurrentMonth
$Form.Controls.Add($TxtMonth)

# 4. 【追加】サーバーからコピーするかのチェックボックス
$CheckCopy = New-Object System.Windows.Forms.CheckBox
$CheckCopy.Text = "サーバーから最新版をコピーする"
$CheckCopy.Location = New-Object System.Drawing.Point(30, 145)
$CheckCopy.Size = New-Object System.Drawing.Size(260, 24)
$CheckCopy.Checked = $false # デフォルトでチェックtreu（サーバーからコピーする）
$Form.Controls.Add($CheckCopy)

# 5. 実行ボタン
$BtnSubmit = New-Object System.Windows.Forms.Button
$BtnSubmit.Text = "エクセルを開く"
$BtnSubmit.Location = New-Object System.Drawing.Point(100, 195)
$BtnSubmit.Size = New-Object System.Drawing.Size(140, 35)

# --------------------------------------------------
# ボタンが押された時の処理
# --------------------------------------------------
$BtnSubmit.Add_Click({
    $Type  = $ComboType.SelectedItem.ToString()
    $Year  = $TxtYear.Text.Trim()
    $Month = $TxtMonth.Text.Trim()

    # 月のゼロ埋め
    if ($Month.Length -eq 1) { $Month = "0" + $Month }

    # コピー元・コピー先のパスを設定
    $DestFolder = "C:\Users\kenic\Documents\operation_log_NEW\$Type"
    $DestPath   = "$DestFolder\${Year}_${Month}_${Type}.xlsm"


    # --- チェックあり（サーバーからコピーして開く） ---
    if ($CheckCopy.Checked) {
        $SourcePath = "\\saclaoprfs01.spring8.or.jp\log_note\$Type\operation_log\$Year\$Month\${Year}_${Month}.xlsm"
        $SourcePath_2 = "\\saclaoprfs01.spring8.or.jp\log_note\$Type\operation_log\$Year\$Month\${Year}_${Month}_${Type}.xlsm"
        # サーバー上のファイルが存在するか確認
        if (Test-Path $SourcePath) {
            $SourcePath_Exist = $SourcePath
        }
        elseif (Test-Path $SourcePath_2) {
            $SourcePath_Exist = $SourcePath_2
        }
        else {
            $SourcePath_Exist = $null
        }
        Write-Host "コピー元: $SourcePath_Exist"
        if (Test-Path $SourcePath_Exist) {
            try {
                if (-not (Test-Path $DestFolder)) {
                    New-Item -ItemType Directory -Path $DestFolder | Out-Null
                }
                Copy-Item -Path $SourcePath_Exist -Destination $DestPath -Force
                
                Invoke-Item -Path $DestPath
                $Form.Close()
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error")
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("サーバー上にファイルが見つかりません。`n`nPath: $SourcePath", "Warning")
        }
    } 
    # --- チェックなし（ローカルのファイルをそのまま開く） ---
    else {
        if (Test-Path $DestPath) {
            Invoke-Item -Path $DestPath
            $Form.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("ローカルにファイルが見つかりません。一度チェックを入れてサーバーからコピーしてください。`n`nPath: $DestPath", "Warning")
        }
    }
})

$Form.Controls.Add($BtnSubmit)

# 画面を表示
[void]$Form.ShowDialog()