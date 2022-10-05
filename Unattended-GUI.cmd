<# :
@pushd "%~dp0"
@set "__CMD__=%~f0" &powershell -noprofile -exec bypass -c "iex([io.file]::ReadAllText(\"%~f0\"))" &exit /b
#>
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(544))
{
start cmd -ArgumentList ("/c `"$Env:__CMD__`"") -Verb runAs ;exit
}
#
#
#
$Global:Choice=0
# Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
#
# Generated Form Objects
$form = New-Object System.Windows.Forms.Form
$textBox1 = New-Object System.Windows.Forms.TextBox
$checkbox7 = New-Object System.Windows.Forms.RadioButton
$checkbox6 = New-Object System.Windows.Forms.RadioButton
$checkbox29 = New-Object System.Windows.Forms.RadioButton
$checkbox28 = New-Object System.Windows.Forms.RadioButton
$checkbox27 = New-Object System.Windows.Forms.CheckBox
$checkbox26 = New-Object System.Windows.Forms.CheckBox
$checkbox25 = New-Object System.Windows.Forms.CheckBox
$checkbox24 = New-Object System.Windows.Forms.CheckBox
$checkbox23 = New-Object System.Windows.Forms.CheckBox
$checkbox22 = New-Object System.Windows.Forms.CheckBox
$checkbox21 = New-Object System.Windows.Forms.CheckBox
$checkbox8 = New-Object System.Windows.Forms.CheckBox
$checkBox5 = New-Object System.Windows.Forms.CheckBox
$checkBox4 = New-Object System.Windows.Forms.CheckBox
$checkBox3 = New-Object System.Windows.Forms.CheckBox
$checkBox2 = New-Object System.Windows.Forms.CheckBox
$checkBox1 = New-Object System.Windows.Forms.CheckBox
$label1 = New-Object System.Windows.Forms.Label
$CancelButton = New-Object System.Windows.Forms.Button
$Choice1 = New-Object System.Windows.Forms.Button
$Choice2 = New-Object System.Windows.Forms.Button
$OKButton = New-Object System.Windows.Forms.Button
#
#
$Form.KeyPreview = $True
$Form.TopMost = $True
$Form.FormBorderStyle = "none"
$Form.Add_KeyDown({if ($_.KeyCode -eq "Enter"){$OKButton.PerformClick()}})         # Enter = click OK button
$Form.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$CancelButton.PerformClick()}})    # Escape = click on Cancel button
$Form.BackColor = "SteelBlue"
$Form.ForeColor = "WhiteSmoke"
$form.ClientSize = New-Object System.Drawing.Size(750,150)
$form.Location = New-Object System.Drawing.Point(452,0)
$form.Name = "form"
$form.Text = "KMS_VL_ALL"

$label1.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",12,0,3,1)
$label1.Location = New-Object System.Drawing.Point(250,50)
$label1.Name = "label1"
$label1.Size = New-Object System.Drawing.Size(300,23)
$label1.Text = "KMS_VL_ALL unattended parameters"
$form.Controls.Add($label1)

# $OKButton.Location = New-Object System.Drawing.Point(562,100)
$OKButton.Name = "OKButton"
$OKButton.Size = New-Object System.Drawing.Size(140,25)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

$CancelButton.Location = New-Object System.Drawing.Point(83,100)
$CancelButton.Name = "CancelButton"
$CancelButton.Size = New-Object System.Drawing.Size(140,25)
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.Controls.Add($CancelButton)

# Autorenewal Setup

$Choice1.Location = New-Object System.Drawing.Point(305,100)
$Choice1.Name = "Choice1"
$Choice1.Size = New-Object System.Drawing.Size(140,25)
$Choice1.Text = "AutoRenewal Setup"
$Choice1.Add_Click({
    $form.Controls.remove($checkbox21)
    $form.Controls.remove($checkbox22)
    $form.Controls.remove($checkbox23)
    $form.Controls.remove($checkbox24)
    $form.Controls.remove($checkbox25)
    $form.Controls.remove($checkbox26)
    $form.Controls.remove($checkbox27)
    $form.Controls.remove($checkbox28)
    $form.Controls.remove($checkbox29)
    $form.Controls.remove($textBox1)
    $form.Controls.Add($checkBox1)
    $form.Controls.Add($checkBox2)
    $form.Controls.Add($checkBox3)
    $form.Controls.Add($checkBox4)
    $form.Controls.Add($checkBox5)
    $form.Controls.Add($checkbox6)
    $form.Controls.Add($checkbox7)
    $form.Controls.Add($checkbox8)
    $form.Controls.Add($OKButton)
    $Checkbox6.Checked=$False
    $Checkbox7.Checked=$False
    $CancelButton.Location = New-Object System.Drawing.Point(38,530)
    $Choice1.Location = New-Object System.Drawing.Point(216,530)
    $Choice2.Location = New-Object System.Drawing.Point(394,530)
    $OKButton.Location = New-Object System.Drawing.Point(572,530)
    $form.ClientSize = New-Object System.Drawing.Size(750,650)
    $form.refresh()
    $Global:Choice=1
})

$form.Controls.Add($Choice1)

$checkBox1.Location = New-Object System.Drawing.Point(80,140)
$checkBox1.Name = "checkBox1"
$checkBox1.Size = New-Object System.Drawing.Size(240,60)
$checkBox1.TabIndex = 14
$checkBox1.Text = "Unattended
(auto install or remove and exit)"

$checkBox2.Location = New-Object System.Drawing.Point(455,140)
$checkBox2.Name = "checkBox2"
$checkBox2.Size = New-Object System.Drawing.Size(240,60)
$checkBox2.TabIndex = 15
$checkBox2.Text = "Silent
(implies Unattended)"

$checkBox3.Location = New-Object System.Drawing.Point(80,200)
$checkBox3.Name = "checkBox3"
$checkBox3.Size = New-Object System.Drawing.Size(240,60)
$checkBox3.TabIndex = 16
$checkBox3.Text = "Debug mode
(implies Unattended)"

$checkBox4.Location = New-Object System.Drawing.Point(455,200)
$checkBox4.Name = "checkBox4"
$checkBox4.Size = New-Object System.Drawing.Size(240,60)
$checkBox4.TabIndex = 17
$checkBox4.Text = "Silent and create simple log"

$checkBox5.Location = New-Object System.Drawing.Point(80,260)
$checkBox5.Name = "checkBox5"
$checkBox5.Size = New-Object System.Drawing.Size(240,60)
$checkBox5.TabIndex = 18
$checkBox5.Text = "Silent Debug mode"

$checkbox6.Location = New-Object System.Drawing.Point(80,320)
$checkbox6.Name = "checkbox6"
$checkbox6.Size = New-Object System.Drawing.Size(240,60)
$checkbox6.TabIndex = 19
$checkbox6.Text = "Force installation regardless detection
(implies Unattended)"

$checkbox7.Location = New-Object System.Drawing.Point(455,320)
$checkbox7.Name = "checkbox7"
$checkbox7.Size = New-Object System.Drawing.Size(240,60)
$checkbox7.TabIndex = 20
$checkbox7.Text = "Force removal regardless detection
(implies Unattended)"

$checkbox8.Location = New-Object System.Drawing.Point(80,380)
$checkbox8.Name = "checkbox8"
$checkbox8.Size = New-Object System.Drawing.Size(240,60)
$checkbox8.TabIndex = 21
$checkbox8.Text = "Do not clear KMS cache"

#
#    Activate
#

$Choice2.Location = New-Object System.Drawing.Point(527,100)
$Choice2.Name = "Choice2"
$Choice2.Size = New-Object System.Drawing.Size(140,25)
$Choice2.Text = "Activate"
$Choice2.Add_Click({
    $form.Controls.remove($checkBox1)
    $form.Controls.remove($checkBox2)
    $form.Controls.remove($checkBox3)
    $form.Controls.remove($checkBox4)
    $form.Controls.remove($checkBox5)
    $form.Controls.remove($checkbox6)
    $form.Controls.remove($checkbox7)
    $form.Controls.remove($checkbox8)
    $form.Controls.Add($checkbox21)
    $form.Controls.Add($checkbox22)
    $form.Controls.Add($checkbox23)
    $form.Controls.Add($checkbox24)
    $form.Controls.Add($checkbox25)
    $form.Controls.Add($checkbox26)
    $form.Controls.Add($checkbox27)
    $form.Controls.Add($checkbox28)
    $form.Controls.Add($checkbox29)
    $form.Controls.Add($textBox1)
    $form.Controls.Add($OKButton)
    $Checkbox28.Checked=$False
    $Checkbox29.Checked=$False
    $CancelButton.Location = New-Object System.Drawing.Point(38,530)
    $Choice1.Location = New-Object System.Drawing.Point(216,530)
    $Choice2.Location = New-Object System.Drawing.Point(394,530)
    $OKButton.Location = New-Object System.Drawing.Point(572,530)
    $form.ClientSize = New-Object System.Drawing.Size(750,650)
    $form.refresh()
    $Global:Choice=2
})

$form.Controls.Add($Choice2)

$checkbox21.Location = New-Object System.Drawing.Point(80,140)
$checkbox21.Name = "checkbox21"
$checkbox21.Size = New-Object System.Drawing.Size(240,60)
$checkbox21.TabIndex = 31
$checkbox21.Text = "Unattended
(auto exit)"

$checkbox22.Location = New-Object System.Drawing.Point(455,140)
$checkbox22.Name = "checkbox22"
$checkbox22.Size = New-Object System.Drawing.Size(240,60)
$checkbox22.TabIndex = 32
$checkbox22.Text = "Silent
(implies Unattended)"

$checkbox23.Location = New-Object System.Drawing.Point(80,200)
$checkbox23.Name = "checkbox23"
$checkbox23.Size = New-Object System.Drawing.Size(240,60)
$checkbox23.TabIndex = 33
$checkbox23.Text = "Debug mode
(implies Unattended)"

$checkbox24.Location = New-Object System.Drawing.Point(455,200)
$checkbox24.Name = "checkbox24"
$checkbox24.Size = New-Object System.Drawing.Size(240,60)
$checkbox24.TabIndex = 34
$checkbox24.Text = "Silent and create simple log"

$checkbox25.Location = New-Object System.Drawing.Point(80,260)
$checkbox25.Name = "checkbox25"
$checkbox25.Size = New-Object System.Drawing.Size(240,60)
$checkbox25.TabIndex = 35
$checkbox25.Text = "Silent Debug mode"

$checkbox26.Location = New-Object System.Drawing.Point(80,320)
$checkbox26.Name = "checkbox26"
$checkbox26.Size = New-Object System.Drawing.Size(240,60)
$checkbox26.TabIndex = 36
$checkbox26.Text = "External activation mode"

$textBox1.Location = New-Object System.Drawing.Point(455,340)
$textBox1.Name = "textBox1"
$textBox1.Size = New-Object System.Drawing.Size(240,60)
$textBox1.TabIndex = 37
$textBox1.Text = "KMS server address"

$checkbox27.Location = New-Object System.Drawing.Point(80,380)
$checkbox27.Name = "checkbox27"
$checkbox27.Size = New-Object System.Drawing.Size(240,60)
$checkbox27.TabIndex = 38
$checkbox27.Text = "Revert Windows 10 KMS38 to normal KMS"

$checkbox28.Location = New-Object System.Drawing.Point(80,440)
$checkbox28.Name = "checkbox28"
$checkbox28.Size = New-Object System.Drawing.Size(240,60)
$checkbox28.TabIndex = 39
$checkbox28.Text = "Activate Office only"

$checkbox29.Location = New-Object System.Drawing.Point(455,440)
$checkbox29.Name = "checkbox29"
$checkbox29.Size = New-Object System.Drawing.Size(240,60)
$checkbox29.TabIndex = 40
$checkbox29.Text = "Activate Windows only"

$result=$form.ShowDialog()
$param_u=$param_s=$param_l=$param_d=$param_i=$param_r=$param_k=$param_o=$param_w=$param_x=$param_e=$param_a=""

if ($result -eq "Cancel"){exit}

#    Autorenewal Setup
if ($Global:Choice -eq 1){

    if ($checkbox1.checked){$param_u=" /u"}
    if ($checkbox2.checked){$param_s=" /s"}
    if ($checkbox3.checked){$param_d=" /d"}
    if ($checkbox4.checked){$param_s=" /s";$param_l=" /l"}
    if ($checkbox5.checked){$param_s=" /s";$param_l=" /d"}
    if ($checkbox6.checked){$param_i=" /i"}
    if ($checkbox7.checked){$param_r=" /r"}
    if ($checkbox8.checked){$param_k=" /k"}
    $cmd="AutoRenewal-Setup.cmd"
}
#    Activate
if ($Global:Choice -eq 2){
    if ($result -eq "Cancel"){exit}
    if ($checkbox21.checked){$param_u=" /u"}
    if ($checkbox22.checked){$param_s=" /s"}
    if ($checkbox23.checked){$param_d=" /d"}
    if ($checkbox24.checked){$param_s=" /s";$param_l=" /l"}
    if ($checkbox25.checked){$param_s=" /s";$param_l=" /d"}
    if ($checkbox27.checked){$param_x=" /x"}
    if ($checkbox28.checked){$param_o=" /o"}
    if ($checkbox29.checked){$param_w=" /w"}
    if ($checkbox26.checked){$a=$textbox1.text;$param_e=" /e " + $a}
    $cmd="Activate.cmd"
}
#
$param=$param_u+$param_s+$param_l+$param_d+$param_i+$param_r+$param_k+$param_o+$param_w+$param_x+$param_e
& .\$cmd $param
# done