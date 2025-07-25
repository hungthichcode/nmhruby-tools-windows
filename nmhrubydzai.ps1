<#
.SYNOPSIS
    Công cụ tối ưu và tùy chỉnh Windows đa năng.
.DESCRIPTION
    Tập lệnh này cung cấp một giao diện dòng lệnh để thực hiện nhiều tác vụ quản trị hệ thống trên Windows,
    bao gồm quản lý chế độ năng lượng, bật/tắt các tính năng hệ thống, cài đặt phần mềm và tùy chỉnh khác.
    Tác giả: Nguyễn Mạnh Hùng (NMHRUBY)
    Phiên bản: 1.0
.NOTES
    Cần chạy với quyền Administrator để có đầy đủ chức năng.
    Hỗ trợ Windows 7 trở lên.
#>

#================================================================================================
# KHỞI TẠO VÀ KIỂM TRA QUYỀN ADMIN
#================================================================================================

# Kiểm tra xem script có đang chạy với quyền Administrator không
if (-Not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Nếu không, tự khởi chạy lại với quyền Admin
    Write-Warning "Yêu cầu quyền Administrator để chạy công cụ này."
    Write-Warning "Vui lòng chấp nhận yêu cầu UAC để tiếp tục."
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Đặt chính sách thực thi cho phiên hiện tại để đảm bảo các lệnh chạy được
Set-ExecutionPolicy Bypass -Scope Process -Force

#================================================================================================
# ĐỊNH NGHĨA NGÔN NGỮ (GUI TEXT)
#================================================================================================

# Sử dụng Hashtable để lưu trữ chuỗi văn bản cho đa ngôn ngữ
$lang_en = @{
    # Menu Chính
    mainTitle             = "==== NMHRUBY TOOLS FOR WINDOWS ===="
    powerOptimize         = "1. Windows Optimize"
    windowsDisable        = "2. Windows Disable"
    windowsOptions        = "3. Windows Options & Fix"
    changeThemes          = "4. Change Themes"
    installSoftware       = "5. Install Software"
    activateTools         = "6. Activate Windows & Office"
    fixGui                = "7. Language (GUI)"
    aboutMe               = "8. About Me"
    exitScript            = "Q. Exit"
    enterChoice           = "Enter your choice: "
    invalidChoice         = "Invalid choice. Please try again."
    pressEnter            = "Press Enter to continue..."

    # Tối ưu Power
    powerMenuTitle        = "==== WINDOWS OPTIMIZE ===="
    powerSaver            = "1. Power saver"
    balanced              = "2. Balanced"
    highPerformance       = "3. High performance"
    ultimatePerformance   = "4. Ultimate performance"
    deletePerformance     = "5. Delete a power plan"
    backToMain            = "B. Back to main menu"
    currentPlan           = "Current active plan"
    chooseYesNo           = "Choose (Y/N): "
    chooseAllNo           = "Choose (A=All / N=No): "
    planActivated         = "Power plan activated successfully."
    planCreated           = "Ultimate Performance plan created and activated."
    noAction              = "No action taken."
    listPlans             = "Available power plans on your system:"
    choosePlanToDelete    = "Enter the number of the plan to delete (or B to back):"
    confirmDelete         = "Are you sure you want to delete this plan?"
    planDeleted           = "Power plan deleted successfully."
    cannotDeleteActive    = "Cannot delete the currently active power plan."

    # Vô hiệu hóa Windows
    disableMenuTitle      = "==== WINDOWS DISABLE ===="
    winUpdate             = "1. Windows Update"
    winSecurity           = "2. Windows Security (Defender)"
    chooseEnableDisable   = "Choose (E=Enable / D=Disable): "
    enabledStatus         = "[Enabled]"
    disabledStatus        = "[Disabled]"
    rebootRequired        = "A system reboot is required for changes to take full effect."
    
    # Tùy chọn & Sửa lỗi
    optionsMenuTitle      = "==== WINDOWS OPTIONS & FIX ===="
    fixTime               = "1. Fix Time & Date"
    changeDeviceName      = "2. Change Device Name"
    changeUserName        = "3. Change User Full Name"
    enterTimeZone         = "Enter the Time Zone ID (e.g., 'SE Asia Standard Time'): "
    timeZoneUpdated       = "Time zone updated successfully."
    invalidTimeZone       = "Error: Invalid Time Zone ID."
    enterNewDeviceName    = "Enter the new device name: "
    confirmDeviceName     = "Are you sure? This requires a restart."
    enterNewFullName      = "Enter the new full name for the user '$env:USERNAME': "
    confirmFullName       = "Are you sure you want to change the full name?"
    fullNameChanged       = "Full name changed successfully."

    # Chủ đề
    themeMenuTitle        = "==== CHANGE THEMES ===="
    yellow                = "Y. Yellow"
    red                   = "R. Red"
    blue                  = "B. Blue"
    green                 = "G. Green"
    purple                = "P. Purple"
    themeChanged          = "Theme color changed."

    # Cài đặt phần mềm
    installMenuTitle      = "==== INSTALL SOFTWARE ===="
    unikey                = "1. Download Unikey"
    evkey                 = "2. Download EVKey"
    winrar                = "3. Install WinRAR"
    systemTypeDetected    = "Detected System Type"
    chooseBit             = "Choose version (64B/32B/ARM): "
    downloading           = "Downloading..."
    downloadedToDesktop   = "has been downloaded to your Desktop."
    installing            = "Installing..."
    installedSuccessfully = "installed successfully."
    activateWinRAR        = "Activate WinRAR? (Y/N): "
    winrarActivated       = "WinRAR activated successfully. Please restart WinRAR."
    winrarNotFound        = "WinRAR installation not found at 'C:\Program Files\WinRAR'."
    
    # Kích hoạt
    activateMenuTitle     = "==== ACTIVATE WINDOWS & OFFICE ===="
    runAsAdmin            = "R. Run with Administrator"
    runAndClean           = "The tool will be deleted after closing."
    
    # Ngôn ngữ
    guiMenuTitle          = "==== LANGUAGE (GUI) ===="
    vietnamese            = "1. Vietnamese"
    english               = "2. English (Default)"
    
    # Giới thiệu
    aboutTitle            = "==== ABOUT ME ===="
    aboutLine1            = "Hello, thank you for trusting and using this tool."
    aboutLine2            = "Rest assured, this tool is coded, built, updated, and shared by Nguyen Manh Hung."
    aboutLine3            = "No malware is installed."
    aboutLine4            = "For inquiries, please contact the links below:"
}

$lang_vi = @{
    # Menu Chính
    mainTitle             = "==== NMHRUBY TOOLS CHO WINDOWS ===="
    powerOptimize         = "1. Tối ưu Windows"
    windowsDisable        = "2. Vô hiệu hóa Windows"
    windowsOptions        = "3. Tùy chọn & Sửa lỗi Windows"
    changeThemes          = "4. Thay đổi Chủ đề"
    installSoftware       = "5. Cài đặt Phần mềm"
    activateTools         = "6. Kích hoạt Windows & Office"
    fixGui                = "7. Ngôn ngữ (GUI)"
    aboutMe               = "8. Về tác giả"
    exitScript            = "Q. Thoát"
    enterChoice           = "Nhập lựa chọn của bạn: "
    invalidChoice         = "Lựa chọn không hợp lệ. Vui lòng thử lại."
    pressEnter            = "Nhấn Enter để tiếp tục..."

    # Tối ưu Power
    powerMenuTitle        = "==== TỐI ƯU WINDOWS ===="
    powerSaver            = "1. Tiết kiệm pin (Power saver)"
    balanced              = "2. Cân bằng (Balanced)"
    highPerformance       = "3. Hiệu năng cao (High performance)"
    ultimatePerformance   = "4. Hiệu năng tối thượng (Ultimate performance)"
    deletePerformance     = "5. Xóa một chế độ năng lượng"
    backToMain            = "B. Quay lại menu chính"
    currentPlan           = "Chế độ đang hoạt động"
    chooseYesNo           = "Chọn (Y/N): "
    chooseAllNo           = "Chọn (A=Tất cả / N=Không): "
    planActivated         = "Kích hoạt chế độ năng lượng thành công."
    planCreated           = "Đã tạo và kích hoạt chế độ Ultimate Performance."
    noAction              = "Không thực hiện hành động nào."
    listPlans             = "Các chế độ năng lượng có sẵn trên máy của bạn:"
    choosePlanToDelete    = "Nhập số của chế độ cần xóa (hoặc B để quay lại):"
    confirmDelete         = "Bạn có chắc chắn muốn xóa chế độ này không?"
    planDeleted           = "Xóa chế độ năng lượng thành công."
    cannotDeleteActive    = "Không thể xóa chế độ năng lượng đang được kích hoạt."

    # Vô hiệu hóa Windows
    disableMenuTitle      = "==== VÔ HIỆU HÓA WINDOWS ===="
    winUpdate             = "1. Windows Update"
    winSecurity           = "2. Windows Security (Defender)"
    chooseEnableDisable   = "Chọn (E=Bật / D=Tắt): "
    enabledStatus         = "[Đã bật]"
    disabledStatus        = "[Đã tắt]"
    rebootRequired        = "Cần khởi động lại máy để thay đổi có hiệu lực hoàn toàn."

    # Tùy chọn & Sửa lỗi
    optionsMenuTitle      = "==== TÙY CHỌN & SỬA LỖI WINDOWS ===="
    fixTime               = "1. Sửa Thời gian & Ngày tháng"
    changeDeviceName      = "2. Đổi tên Thiết bị"
    changeUserName        = "3. Đổi Tên đầy đủ Người dùng"
    enterTimeZone         = "Nhập ID Múi giờ (ví dụ: 'SE Asia Standard Time'): "
    timeZoneUpdated       = "Cập nhật múi giờ thành công."
    invalidTimeZone       = "Lỗi: ID Múi giờ không hợp lệ."
    enterNewDeviceName    = "Nhập tên thiết bị mới: "
    confirmDeviceName     = "Bạn có chắc không? Hành động này yêu cầu khởi động lại máy."
    enterNewFullName      = "Nhập tên đầy đủ mới cho người dùng '$env:USERNAME': "
    confirmFullName       = "Bạn có chắc muốn đổi tên đầy đủ không?"
    fullNameChanged       = "Đổi tên đầy đủ thành công."

    # Chủ đề
    themeMenuTitle        = "==== THAY ĐỔI CHỦ ĐỀ ===="
    yellow                = "Y. Vàng"
    red                   = "R. Đỏ"
    blue                  = "B. Xanh dương"
    green                 = "G. Xanh lá"
    purple                = "P. Tím"
    themeChanged          = "Đã đổi màu chủ đề."

    # Cài đặt phần mềm
    installMenuTitle      = "==== CÀI ĐẶT PHẦN MỀM ===="
    unikey                = "1. Tải Unikey"
    evkey                 = "2. Tải EVKey"
    winrar                = "3. Cài đặt WinRAR"
    systemTypeDetected    = "Kiến trúc hệ thống"
    chooseBit             = "Chọn phiên bản (64B/32B/ARM): "
    downloading           = "Đang tải xuống..."
    downloadedToDesktop   = "đã được tải về Desktop của bạn."
    installing            = "Đang cài đặt..."
    installedSuccessfully = "đã được cài đặt thành công."
    activateWinRAR        = "Kích hoạt WinRAR? (Y/N): "
    winrarActivated       = "Kích hoạt WinRAR thành công. Vui lòng khởi động lại WinRAR."
    winrarNotFound        = "Không tìm thấy thư mục cài đặt WinRAR tại 'C:\Program Files\WinRAR'."

    # Kích hoạt
    activateMenuTitle     = "==== KÍCH HOẠT WINDOWS & OFFICE ===="
    runAsAdmin            = "R. Chạy với quyền Administrator"
    runAndClean           = "Công cụ sẽ bị xóa sau khi bạn đóng nó."

    # Ngôn ngữ
    guiMenuTitle          = "==== NGÔN NGỮ (GUI) ===="
    vietnamese            = "1. Tiếng Việt"
    english               = "2. Tiếng Anh (Mặc định)"

    # Giới thiệu
    aboutTitle            = "==== VỀ TÁC GIẢ ===="
    aboutLine1            = "Xin chào các bạn, cảm ơn các bạn đã tin tưởng và sử dụng."
    aboutLine2            = "Yên tâm Tools được Nguyễn Mạnh Hùng Code, Build, Update, Share."
    aboutLine3            = "Không cài mã độc vào."
    aboutLine4            = "Thắc mắc xin liên hệ các link bên dưới"
}

# Đặt ngôn ngữ mặc định
$lang = $lang_en

#================================================================================================
# ĐỒNG HỒ ĐẾM NGƯỢC VÀ THỜI GIAN
#================================================================================================

# Chạy một công việc nền để cập nhật tiêu đề cửa sổ console
$countdownJob = Start-Job -ScriptBlock {
    $endTime = (Get-Date).AddHours(24)
    while ($true) {
        $remaining = $endTime - (Get-Date)
        $dateStr = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
        $title = "NMHRUBY Tools | Countdown: $($remaining.ToString('hh\:mm\:ss')) | Current: $($dateStr)"
        $Host.UI.RawUI.WindowTitle = $title
        Start-Sleep -Seconds 1
    }
}

#================================================================================================
# CÁC HÀM CHỨC NĂNG
#================================================================================================

# Hàm hiển thị banner
function Show-Banner {
    # Màu cho banner
    $bannerColor = "Cyan"
    Write-Host @"
███╗   ██╗███╗   ███╗██╗  ██╗██████╗ ██╗   ██╗██████╗ ██╗   ██╗     
████╗  ██║████╗ ████║██║  ██║██╔══██╗██║   ██║██╔══██╗╚██╗ ██╔╝     
██╔██╗ ██║██╔████╔██║███████║██████╔╝██║   ██║██████╔╝ ╚████╔╝      
██║╚██╗██║██║╚██╔╝██║██╔══██║██╔══██╗██║   ██║██╔══██╗  ╚██╔╝       
██║ ╚████║██║ ╚═╝ ██║██║  ██║██║  ██║╚██████╔╝██████╔╝   ██║        
╚═╝  ╚═══╝╚═╝     ╚═╝╚═╝  ╚═╝╚═╝  ╚═╝ ╚═════╝ ╚═════╝    ╚═╝        
                                                                    
████████╗ ██████╗  ██████╗ ██╗     ███████╗     ██╗   ██╗   ██████╗ 
╚══██╔══╝██╔═══██╗██╔═══██╗██║     ██╔════╝     ███║  ███║   ╚════██╗
   ██║   ██║   ██║██║   ██║██║     ███████╗     ╚██║  ╚██║    █████╔╝
   ██║   ██║   ██║██║   ██║██║     ╚════██║      ██║   ██║    ╚═══██╗
   ██║   ╚██████╔╝╚██████╔╝███████╗███████║      ██║██╗██║██╗██████╔╝
   ╚═╝    ╚═════╝  ╚═════╝ ╚══════╝╚══════╝      ╚═╝╚═╝╚═╝╚═╝╚═════╝ 
                                                                    
"@ -ForegroundColor $bannerColor
}

# Hàm tạm dừng màn hình
function Pause-Screen {
    Write-Host "`n$($lang.pressEnter)" -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Hàm chạy lệnh CMD ẩn
function Invoke-CmdHidden {
    param(
        [string]$Command
    )
    Start-Process cmd.exe -ArgumentList "/c $Command" -WindowStyle Hidden -Wait
}

#-------------------------------------------------
# 1. WINDOWS OPTIMIZE
#-------------------------------------------------
function Menu-WindowsOptimize {
    while ($true) {
        Clear-Host
        Show-Banner
        Write-Host "==== $($lang.powerMenuTitle) ====" -ForegroundColor Green
        
        # Lấy và hiển thị power plan hiện tại
        try {
            $currentPlan = powercfg /getactivescheme
            Write-Host "$($lang.currentPlan): $($currentPlan.Split('(')[1].Trim(')'))" -ForegroundColor Cyan
        } catch {
            Write-Warning "Could not retrieve active power plan."
        }

        Write-Host "1. $($lang.powerSaver)"
        Write-Host "2. $($lang.balanced)"
        Write-Host "3. $($lang.highPerformance)"
        Write-Host "4. $($lang.ultimatePerformance)"
        Write-Host "5. $($lang.deletePerformance)"
        Write-Host "B. $($lang.backToMain)"
        $choice = Read-Host "`n$($lang.enterChoice)"

        switch ($choice) {
            '1' {
                $confirm = Read-Host "$($lang.chooseYesNo)"
                if ($confirm -eq 'y') {
                    Invoke-CmdHidden "powercfg /setactive a1841308-3541-4fab-bc81-f71556f20b4a"
                    Write-Host "`n$($lang.planActivated)" -ForegroundColor Green
                } else { Write-Host "`n$($lang.noAction)" -ForegroundColor Yellow }
                Pause-Screen
            }
            '2' {
                $confirm = Read-Host "$($lang.chooseYesNo)"
                if ($confirm -eq 'y') {
                    Invoke-CmdHidden "powercfg /setactive 381b4222-f694-41f0-9685-ff5bb260df2e"
                    Write-Host "`n$($lang.planActivated)" -ForegroundColor Green
                } else { Write-Host "`n$($lang.noAction)" -ForegroundColor Yellow }
                Pause-Screen
            }
            '3' {
                $confirm = Read-Host "$($lang.chooseYesNo)"
                if ($confirm -eq 'y') {
                    Invoke-CmdHidden "powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
                    Write-Host "`n$($lang.planActivated)" -ForegroundColor Green
                } else { Write-Host "`n$($lang.noAction)" -ForegroundColor Yellow }
                Pause-Screen
            }
            '4' {
                $confirm = Read-Host "$($lang.chooseAllNo)"
                if ($confirm -eq 'a') {
                    Invoke-CmdHidden "powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61"
                    Invoke-CmdHidden "powercfg /setactive e9a42b02-d5df-448d-aa00-03f14749eb61"
                    Write-Host "`n$($lang.planCreated)" -ForegroundColor Green
                } else { Write-Host "`n$($lang.noAction)" -ForegroundColor Yellow }
                Pause-Screen
            }
            '5' {
                # Liệt kê các plan
                $plans = @()
                powercfg /list | ForEach-Object {
                    if ($_ -match 'GUID: (.*)  \((.*)\)') {
                        $plans += [PSCustomObject]@{
                            GUID = $matches[1].Trim()
                            Name = $matches[2].Trim()
                        }
                    }
                }

                Write-Host "`n$($lang.listPlans)" -ForegroundColor Cyan
                for ($i = 0; $i -lt $plans.Count; $i++) {
                    Write-Host "$($i+1). $($plans[$i].Name)"
                }

                $deleteChoice = Read-Host "`n$($lang.choosePlanToDelete)"
                if ($deleteChoice -eq 'b') { continue }
                $idx = [int]$deleteChoice - 1

                if ($idx -ge 0 -and $idx -lt $plans.Count) {
                    $planToDelete = $plans[$idx]
                    $activeGuid = (powercfg /getactivescheme).Split(' ')[3]
                    
                    if ($planToDelete.GUID -eq $activeGuid) {
                        Write-Warning $lang.cannotDeleteActive
                    } else {
                        $confirmDelete = Read-Host "$($lang.confirmDelete) '$($planToDelete.Name)'? ($($lang.chooseYesNo))"
                        if ($confirmDelete -eq 'y') {
                            Invoke-CmdHidden "powercfg /delete $($planToDelete.GUID)"
                            Write-Host "`n$($lang.planDeleted)" -ForegroundColor Green
                        } else {
                            Write-Host "`n$($lang.noAction)" -ForegroundColor Yellow
                        }
                    }
                } else {
                    Write-Warning $lang.invalidChoice
                }
                Pause-Screen
            }
            'b' { return }
            default { Write-Warning $lang.invalidChoice; Start-Sleep 1 }
        }
    }
}

#-------------------------------------------------
# 2. WINDOWS DISABLE
#-------------------------------------------------
function Get-ServiceStatus {
    param([string]$ServiceName)
    try {
        $service = Get-Service -Name $ServiceName -ErrorAction Stop
        return $service.Status
    } catch {
        return "Not Found"
    }
}

function Get-WindowsUpdateStatus {
    try {
        $key = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
        if (Test-Path $key) {
            $value = Get-ItemProperty -Path $key -Name "NoAutoUpdate" -ErrorAction SilentlyContinue
            if ($value -and $value.NoAutoUpdate -eq 1) {
                return $lang.disabledStatus
            }
        }
        # Kiểm tra trạng thái dịch vụ
        $wuauservStatus = Get-ServiceStatus -ServiceName "wuauserv"
        if ($wuauservStatus -eq "Disabled" -or $wuauservStatus -eq "Stopped") {
             return $lang.disabledStatus
        }
        return $lang.enabledStatus
    } catch {
        return $lang.enabledStatus
    }
}

function Get-WindowsSecurityStatus {
    try {
        $key = "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender"
        if (Test-Path $key) {
            $value = Get-ItemProperty -Path $key -Name "DisableAntiSpyware" -ErrorAction SilentlyContinue
            if ($value -and $value.DisableAntiSpyware -eq 1) {
                return $lang.disabledStatus
            }
        }
        # Kiểm tra dịch vụ
        $windefendStatus = Get-ServiceStatus -ServiceName "WinDefend"
        if ($windefendStatus -eq "Disabled" -or $windefendStatus -eq "Stopped") {
             return $lang.disabledStatus
        }
        return $lang.enabledStatus
    } catch {
        return $lang.enabledStatus
    }
}

function Menu-WindowsDisable {
    while ($true) {
        Clear-Host
        Show-Banner
        Write-Host "==== $($lang.disableMenuTitle) ====" -ForegroundColor Green
        
        # Hiển thị trạng thái hiện tại
        $updateStatus = Get-WindowsUpdateStatus
        $securityStatus = Get-WindowsSecurityStatus

        Write-Host "1. $($lang.winUpdate) - $updateStatus"
        Write-Host "2. $($lang.winSecurity) - $securityStatus"
        Write-Host "B. $($lang.backToMain)"
        $choice = Read-Host "`n$($lang.enterChoice)"

        switch ($choice) {
            '1' {
                $action = Read-Host "$($lang.chooseEnableDisable)"
                $updateKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
                if (!(Test-Path $updateKey)) { New-Item -Path $updateKey -Force | Out-Null }
                
                if ($action -eq 'd') {
                    # Tắt qua Registry và dịch vụ
                    Set-ItemProperty -Path $updateKey -Name "NoAutoUpdate" -Value 1 -Type DWord -Force
                    Set-ItemProperty -Path $updateKey -Name "AUOptions" -Value 1 -Type DWord -Force
                    Stop-Service -Name "wuauserv" -Force -ErrorAction SilentlyContinue
                    Set-Service -Name "wuauserv" -StartupType Disabled -ErrorAction SilentlyContinue
                    Write-Host "`nWindows Update has been disabled." -ForegroundColor Green
                } elseif ($action -eq 'e') {
                    # Bật lại
                    Set-ItemProperty -Path $updateKey -Name "NoAutoUpdate" -Value 0 -Type DWord -Force
                    Set-ItemProperty -Path $updateKey -Name "AUOptions" -Value 4 -Type DWord -Force
                    Set-Service -Name "wuauserv" -StartupType Automatic -ErrorAction SilentlyContinue
                    Start-Service -Name "wuauserv" -ErrorAction SilentlyContinue
                    Write-Host "`nWindows Update has been enabled." -ForegroundColor Green
                }
                Write-Warning $lang.rebootRequired
                Pause-Screen
            }
            '2' {
                $action = Read-Host "$($lang.chooseEnableDisable)"
                $defenderKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender"
                if (!(Test-Path $defenderKey)) { New-Item -Path $defenderKey -Force | Out-Null }

                if ($action -eq 'd') {
                    Set-ItemProperty -Path $defenderKey -Name "DisableAntiSpyware" -Value 1 -Type DWord -Force
                    Stop-Service -Name "WinDefend" -Force -ErrorAction SilentlyContinue
                    Set-Service -Name "WinDefend" -StartupType Disabled -ErrorAction SilentlyContinue
                    Write-Host "`nWindows Security has been disabled." -ForegroundColor Green
                } elseif ($action -eq 'e') {
                    Set-ItemProperty -Path $defenderKey -Name "DisableAntiSpyware" -Value 0 -Type DWord -Force
                    Set-Service -Name "WinDefend" -StartupType Automatic -ErrorAction SilentlyContinue
                    Start-Service -Name "WinDefend" -ErrorAction SilentlyContinue
                    Write-Host "`nWindows Security has been enabled." -ForegroundColor Green
                }
                Write-Warning $lang.rebootRequired
                Pause-Screen
            }
            'b' { return }
            default { Write-Warning $lang.invalidChoice; Start-Sleep 1 }
        }
    }
}

#-------------------------------------------------
# 3. WINDOWS OPTIONS & FIX
#-------------------------------------------------
function Menu-WindowsOptions {
     while ($true) {
        Clear-Host
        Show-Banner
        Write-Host "==== $($lang.optionsMenuTitle) ====" -ForegroundColor Green
        Write-Host "1. $($lang.fixTime)"
        Write-Host "2. $($lang.changeDeviceName)"
        Write-Host "3. $($lang.changeUserName)"
        Write-Host "B. $($lang.backToMain)"
        $choice = Read-Host "`n$($lang.enterChoice)"

        switch ($choice) {
            '1' {
                Write-Host "Current Time Zone: $((Get-TimeZone).DisplayName)" -ForegroundColor Cyan
                $tz = Read-Host "`n$($lang.enterTimeZone)"
                try {
                    Set-TimeZone -Id $tz -ErrorAction Stop
                    # Đồng bộ thời gian
                    w32tm /resync /force | Out-Null
                    Write-Host "`n$($lang.timeZoneUpdated)" -ForegroundColor Green
                } catch {
                    Write-Warning "`n$($lang.invalidTimeZone)"
                }
                Pause-Screen
            }
            '2' {
                Write-Host "Current Device Name: $env:COMPUTERNAME" -ForegroundColor Cyan
                $newName = Read-Host "`n$($lang.enterNewDeviceName)"
                if ([string]::IsNullOrWhiteSpace($newName)) {
                    Write-Warning "Device name cannot be empty."; Pause-Screen; continue
                }
                $confirm = Read-Host "$($lang.confirmDeviceName) ($($lang.chooseYesNo))"
                if ($confirm -eq 'y') {
                    try {
                        Rename-Computer -NewName $newName -Force -ErrorAction Stop
                        Write-Host "`nDevice name will be changed to '$newName' after restart." -ForegroundColor Green
                        Write-Warning "Please restart your computer for the change to take effect."
                    } catch {
                        Write-Error "Failed to rename computer. Error: $($_.Exception.Message)"
                    }
                } else { Write-Host "`n$($lang.noAction)" -ForegroundColor Yellow }
                Pause-Screen
            }
            '3' {
                Write-Host "Current User: $env:USERNAME" -ForegroundColor Cyan
                $newFullName = Read-Host "`n$($lang.enterNewFullName)"
                 if ([string]::IsNullOrWhiteSpace($newFullName)) {
                    Write-Warning "Full name cannot be empty."; Pause-Screen; continue
                }
                $confirm = Read-Host "$($lang.confirmFullName) ($($lang.chooseYesNo))"
                if ($confirm -eq 'y') {
                    try {
                        # Lệnh này an toàn và chỉ thay đổi tên hiển thị, không đổi tên thư mục profile
                        net user "$env:USERNAME" /fullname:"$newFullName"
                        Write-Host "`n$($lang.fullNameChanged)" -ForegroundColor Green
                    } catch {
                        Write-Error "Failed to change full name. Error: $($_.Exception.Message)"
                    }
                } else { Write-Host "`n$($lang.noAction)" -ForegroundColor Yellow }
                Pause-Screen
            }
            'b' { return }
            default { Write-Warning $lang.invalidChoice; Start-Sleep 1 }
        }
    }
}

#-------------------------------------------------
# 4. CHANGE THEMES
#-------------------------------------------------
function Menu-ChangeThemes {
    Clear-Host
    Show-Banner
    Write-Host "==== $($lang.themeMenuTitle) ====" -ForegroundColor Green
    Write-Host "Y. $($lang.yellow)"
    Write-Host "R. $($lang.red)"
    Write-Host "B. $($lang.blue)"
    Write-Host "G. $($lang.green)"
    Write-Host "P. $($lang.purple)"
    Write-Host "B. $($lang.backToMain)"
    $choice = Read-Host "`n$($lang.enterChoice)"
    
    $color = $Host.UI.RawUI.ForegroundColor
    switch ($choice.ToLower()) {
        'y' { $color = "Yellow" }
        'r' { $color = "Red" }
        'b' { $color = "Blue" }
        'g' { $color = "Green" }
        'p' { $color = "Magenta" } # Purple is Magenta in console
        'b' { return }
        default { Write-Warning $lang.invalidChoice; Start-Sleep 1; return }
    }
    $Host.UI.RawUI.ForegroundColor = $color
    Write-Host "`n$($lang.themeChanged)" -ForegroundColor $color
    Start-Sleep 1
}

#-------------------------------------------------
# 5. INSTALL SOFTWARE
#-------------------------------------------------
function Download-File {
    param(
        [string]$Url,
        [string]$FileName
    )
    $desktopPath = [System.Environment]::GetFolderPath('Desktop')
    $outputPath = Join-Path -Path $desktopPath -ChildPath $FileName
    Write-Host "`n$($lang.downloading) $FileName..." -ForegroundColor Cyan
    try {
        # Sử dụng Invoke-WebRequest với các tùy chọn để tương thích tốt hơn
        Invoke-WebRequest -Uri $Url -OutFile $outputPath -UseBasicParsing
        Write-Host "$FileName $($lang.downloadedToDesktop)" -ForegroundColor Green
        return $outputPath
    } catch {
        Write-Error "Download failed for $Url. Error: $($_.Exception.Message)"
        return $null
    }
}

function Menu-InstallSoftware {
    while ($true) {
        Clear-Host
        Show-Banner
        Write-Host "==== $($lang.installMenuTitle) ====" -ForegroundColor Green
        
        $arch = (Get-CimInstance Win32_OperatingSystem).OSArchitecture
        Write-Host "$($lang.systemTypeDetected): $arch" -ForegroundColor Cyan

        Write-Host "1. $($lang.unikey)"
        Write-Host "2. $($lang.evkey)"
        Write-Host "3. $($lang.winrar)"
        Write-Host "B. $($lang.backToMain)"
        $choice = Read-Host "`n$($lang.enterChoice)"

        switch ($choice) {
            '1' {
                $bitChoice = Read-Host "$($lang.chooseBit)"
                $url = ""
                $fileName = ""
                switch ($bitChoice.ToLower()) {
                    '64b' {
                        $url = "https://www.unikey.org/assets/release/unikey46RC2-230919-win64.zip"
                        $fileName = "unikey46RC2-win64.zip"
                    }
                    '32b' {
                        $url = "https://www.unikey.org/assets/release/unikey46RC2-230919-win32.zip"
                        $fileName = "unikey46RC2-win32.zip"
                    }
                    'arm' {
                        $url = "https://www.unikey.org/assets/release/unikey46RC2-250531-arm64.zip"
                        $fileName = "unikey46RC2-arm64.zip"
                    }
                    default { Write-Warning $lang.invalidChoice; Pause-Screen; continue }
                }
                $zipPath = Download-File -Url $url -FileName $fileName
                if ($zipPath) {
                    Expand-Archive -Path $zipPath -DestinationPath "$([System.Environment]::GetFolderPath('Desktop'))\Unikey" -Force
                }
                Pause-Screen
            }
            '2' {
                $confirm = Read-Host "$($lang.chooseYesNo)"
                if ($confirm -eq 'y') {
                    $url = "https://github.com/lamquangminh/EVKey/releases/download/Release/EVKey.zip"
                    $fileName = "EVKey.zip"
                    $zipPath = Download-File -Url $url -FileName $fileName
                    if ($zipPath) {
                        Expand-Archive -Path $zipPath -DestinationPath "$([System.Environment]::GetFolderPath('Desktop'))\EVKey" -Force
                    }
                } else { Write-Host "`n$($lang.noAction)" -ForegroundColor Yellow }
                Pause-Screen
            }
            '3' {
                # WinRAR chỉ có bản 64-bit trên trang chủ chính thức gần đây
                $url = "https://www.win-rar.com/fileadmin/winrar-versions/winrar/winrar-x64-701.exe" # Link ổn định hơn
                $fileName = "winrar-x64.exe"
                $exePath = Download-File -Url $url -FileName $fileName
                if ($exePath) {
                    Write-Host "`n$($lang.installing) WinRAR..." -ForegroundColor Cyan
                    Start-Process -FilePath $exePath -ArgumentList "/S" -Wait
                    Remove-Item $exePath -Force
                    Write-Host "`nWinRAR $($lang.installedSuccessfully)" -ForegroundColor Green

                    $activate = Read-Host "$($lang.activateWinRAR)"
                    if ($activate -eq 'y') {
                        $keyUrl = "https://raw.githubusercontent.com/hungthichcode/nmhruby-tools-windows/master/rarreg.key"
                        $winrarPath = "C:\Program Files\WinRAR"
                        if (Test-Path $winrarPath) {
                            $keyPath = Join-Path $winrarPath "rarreg.key"
                            try {
                                Invoke-WebRequest -Uri $keyUrl -OutFile $keyPath -UseBasicParsing
                                Write-Host "`n$($lang.winrarActivated)" -ForegroundColor Green
                            } catch {
                                Write-Error "Failed to download activation key. Error: $($_.Exception.Message)"
                            }
                        } else {
                            Write-Warning $lang.winrarNotFound
                        }
                    }
                }
                Pause-Screen
            }
            'b' { return }
            default { Write-Warning $lang.invalidChoice; Start-Sleep 1 }
        }
    }
}

#-------------------------------------------------
# 6. ACTIVATE WINDOWS & OFFICE
#-------------------------------------------------
function Menu-Activate {
    while ($true) {
        Clear-Host
        Show-Banner
        Write-Host "==== $($lang.activateMenuTitle) ====" -ForegroundColor Green
        Write-Host "R. $($lang.runAsAdmin)"
        Write-Host "   ($($lang.runAndClean))"
        Write-Host "N. $($lang.noAction)"
        Write-Host "B. $($lang.backToMain)"
        $choice = Read-Host "`n$($lang.enterChoice)"

        switch ($choice.ToLower()) {
            'r' {
                $url = "https://raw.githubusercontent.com/hungthichcode/nmhruby-tools-windows/master/NMHRUBY.bat"
                $fileName = "NMHRUBY.bat"
                $batPath = Download-File -Url $url -FileName $fileName
                if ($batPath) {
                    Start-Process -FilePath $batPath -Verb RunAs -Wait
                    Remove-Item $batPath -Force -ErrorAction SilentlyContinue
                    Write-Host "`nActivation tool has been run and cleaned up." -ForegroundColor Green
                }
                Pause-Screen
            }
            'n' { Write-Host "`n$($lang.noAction)" -ForegroundColor Yellow; Pause-Screen }
            'b' { return }
            default { Write-Warning $lang.invalidChoice; Start-Sleep 1 }
        }
    }
}

#-------------------------------------------------
# 7. LANGUAGE (GUI)
#-------------------------------------------------
function Menu-FixGui {
    Clear-Host
    Show-Banner
    Write-Host "==== $($lang.guiMenuTitle) ====" -ForegroundColor Green
    Write-Host "1. $($lang_vi.vietnamese)"
    Write-Host "2. $($lang_en.english)"
    Write-Host "B. $($lang.backToMain)"
    $choice = Read-Host "`n$($lang.enterChoice)"
    
    switch ($choice) {
        '1' { $script:lang = $lang_vi }
        '2' { $script:lang = $lang_en }
        'b' { return }
        default { Write-Warning $lang.invalidChoice; Start-Sleep 1 }
    }
}

#-------------------------------------------------
# 8. ABOUT ME
#-------------------------------------------------
function Menu-AboutMe {
    Clear-Host
    Show-Banner
    Write-Host "==== $($lang.aboutTitle) ====" -ForegroundColor Green
    Write-Host "`n$($lang.aboutLine1)"
    Write-Host $lang.aboutLine2
    Write-Host $lang.aboutLine3
    Write-Host "`n$($lang.aboutLine4)`n"
    Write-Host "[Facebook] = https://www.facebook.com/NMHRUBY" -ForegroundColor Cyan
    Write-Host "[Github]   = https://github.com/hungthichcode/hungdevelopers2kar6" -ForegroundColor Cyan
    Write-Host "[Website]  = https://hungthichcode.github.io/nmhruby/" -ForegroundColor Cyan
    Pause-Screen
}


#================================================================================================
# VÒNG LẶP CHÍNH CỦA CHƯƠNG TRÌNH
#================================================================================================
try {
    while ($true) {
        Clear-Host
        Show-Banner
        # Đặt lại màu mặc định mỗi lần vẽ lại menu
        $Host.UI.RawUI.ForegroundColor = [System.ConsoleColor]::Gray
        
        Write-Host "==== $($lang.mainTitle) ====" -ForegroundColor Green
        Write-Host "1. $($lang.powerOptimize)"
        Write-Host "2. $($lang.windowsDisable)"
        Write-Host "3. $($lang.windowsOptions)"
        Write-Host "4. $($lang.changeThemes)"
        Write-Host "5. $($lang.installSoftware)"
        Write-Host "6. $($lang.activateTools)"
        Write-Host "7. $($lang.fixGui)"
        Write-Host "8. $($lang.aboutMe)"
        Write-Host "Q. $($lang.exitScript)"
        
        $choice = Read-Host "`n$($lang.enterChoice)"

        switch ($choice.ToLower()) {
            '1' { Menu-WindowsOptimize }
            '2' { Menu-WindowsDisable }
            '3' { Menu-WindowsOptions }
            '4' { Menu-ChangeThemes }
            '5' { Menu-InstallSoftware }
            '6' { Menu-Activate }
            '7' { Menu-FixGui }
            '8' { Menu-AboutMe }
            'q' { break }
            default {
                Write-Warning $lang.invalidChoice
                Start-Sleep -Seconds 1
            }
        }
    }
}
finally {
    # Dọn dẹp khi script kết thúc
    Write-Host "Exiting tool. Cleaning up background jobs..." -ForegroundColor Yellow
    Stop-Job -Job $countdownJob
    Remove-Job -Job $countdownJob
    # Đặt lại tiêu đề cửa sổ
    $Host.UI.RawUI.WindowTitle = "PowerShell"
}
