#requires -Version 5.1
<#
.SYNOPSIS
    Быстрые заметки (QuickNotes) — лёгкая утилита для Windows: создание текстовых
    файлов в один клик.

.DESCRIPTION
    Каждая пользовательская кнопка привязана к своей папке. Нажатие на кнопку
    открывает окно заметки, в котором имя файла уже заполнено текущей датой,
    а текст можно отредактировать и сохранить в .txt.

    Кнопки хранятся в файле buttons.json рядом со скриптом (если папка доступна
    для записи), иначе — в %APPDATA%\QuickNotes\buttons.json.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File QuickNotes.ps1

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File QuickNotes.ps1 -SelfTest
    Запуск встроенных проверок (без GUI).
#>

[CmdletBinding()]
param(
    # Запустить встроенные проверки логики и разметки, не открывая окно.
    [switch]$SelfTest
)

$ErrorActionPreference = 'Stop'

# ============================== НАСТРОЙКИ ====================================

$script:AppTitle   = 'Быстрые заметки'
$script:Extension  = '.txt'
$script:DateFormat = 'yyyy-MM-dd'      # формат имени новой заметки по умолчанию
$script:ConfigName = 'buttons.json'

# ========================== РАБОТА С ДАННЫМИ =================================

function Get-DataDirectory {
    # Возвращает папку для хранения конфигурации: рядом со скриптом, если она
    # доступна для записи, иначе %APPDATA%\QuickNotes.
    if ($script:DataDirectory) { return $script:DataDirectory }

    $dir = $PSScriptRoot
    $writable = $false
    if ($dir) {
        try {
            $probe = Join-Path $dir ('.qn-write-test-{0}.tmp' -f [guid]::NewGuid().ToString('N'))
            [System.IO.File]::WriteAllText($probe, 'ok')
            Remove-Item -LiteralPath $probe -Force
            $writable = $true
        } catch {
            $writable = $false
        }
    }
    if (-not $writable) {
        $base = $env:APPDATA
        if (-not $base) { $base = [System.IO.Path]::GetTempPath() }
        $dir = Join-Path $base 'QuickNotes'
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
    }
    $script:DataDirectory = $dir
    return $dir
}

function Get-ConfigFilePath {
    return (Join-Path (Get-DataDirectory) $script:ConfigName)
}

function Read-ButtonConfig {
    param([string]$Path = (Get-ConfigFilePath))

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @() }
    $raw = [System.IO.File]::ReadAllText($Path)
    if ([string]::IsNullOrWhiteSpace($raw)) { return @() }

    $parsed = $null
    try {
        $parsed = ConvertFrom-Json -InputObject $raw
    } catch {
        # Конфигурация повреждена — убираем её в резервный файл и начинаем с пустого списка.
        try { Move-Item -LiteralPath $Path -Destination ($Path + '.corrupt') -Force } catch { }
        return @()
    }
    if ($null -eq $parsed) { return @() }

    if ($parsed -is [System.Array]) {
        $rawItems = $parsed
    } elseif ($null -ne $parsed.buttons) {
        $rawItems = @($parsed.buttons)
    } else {
        $rawItems = @()
    }

    $result = @()
    foreach ($item in $rawItems) {
        if ($null -eq $item) { continue }
        $name   = [string]$item.name
        $folder = [string]$item.folder
        if ($name) {
            $result += [pscustomobject]@{ Name = $name; Folder = $folder }
        }
    }
    return ,$result
}

function Save-ButtonConfig {
    param(
        [object[]]$Buttons = @(),
        [string]$Path = (Get-ConfigFilePath)
    )

    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($b in $Buttons) {
        if ($null -ne $b) {
            $entries.Add([pscustomobject]@{ name = [string]$b.Name; folder = [string]$b.Folder })
        }
    }
    $config = [pscustomobject]@{ buttons = $entries.ToArray() }
    $json   = ConvertTo-Json -InputObject $config -Depth 4

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($Path, $json, (New-Object System.Text.UTF8Encoding($true)))
}

# ============================ ВСПОМОГАТЕЛЬНЫЕ ================================

function ConvertTo-SafeFileName {
    # Убирает из имени файла запрещённые символы и крайние точки/пробелы.
    param([string]$Name)

    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Name.ToCharArray()) {
        if ($invalid -notcontains $ch) { [void]$sb.Append($ch) }
    }
    $clean = $sb.ToString().Trim()
    $clean = $clean.TrimEnd('.', ' ')
    if ([string]::IsNullOrWhiteSpace($clean)) { return $null }
    return $clean
}

function Get-UniqueFilePath {
    # Возвращает несуществующий путь: "имя.txt" -> "имя (2).txt" -> "имя (3).txt" ...
    param(
        [string]$Folder,
        [string]$BaseName,
        [string]$Extension = $script:Extension
    )

    $candidate = Join-Path $Folder ($BaseName + $Extension)
    $i = 2
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $Folder ('{0} ({1}){2}' -f $BaseName, $i, $Extension)
        $i++
    }
    return $candidate
}

function Resolve-FolderPath {
    # Разворачивает переменные окружения и относительные пути.
    param([string]$Path)

    $p = [Environment]::ExpandEnvironmentVariables($Path.Trim())
    if (-not [System.IO.Path]::IsPathRooted($p)) {
        $p = Join-Path (Get-Location).Path $p
    }
    return $p
}

function Open-FolderInExplorer {
    param([string]$Folder)

    if (Test-Path -LiteralPath $Folder -PathType Container) {
        Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $Folder)
    } else {
        Show-MessageBox -Owner $script:Window -Message ("Папка не существует:`n{0}" -f $Folder) -Icon 'Warning'
    }
}

# ================================ РАЗМЕТКА ===================================

$script:CommonStyles = @'
        <Style x:Key="AccentButton" TargetType="{x:Type Button}">
            <Setter Property="Background" Value="#7B6CF6"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" RecognizesAccessKey="True"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#8F83F8"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#6A5BE0"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Bd" Property="Background" Value="#3A3A45"/>
                                <Setter Property="Foreground" Value="#71717C"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="GhostButton" TargetType="{x:Type Button}">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#C7C7D1"/>
            <Setter Property="BorderBrush" Value="#3A3A45"/>
            <Setter Property="Padding" Value="14,7"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" RecognizesAccessKey="True"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#2A2A33"/>
                                <Setter TargetName="Bd" Property="BorderBrush" Value="#4A4A57"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#26262E"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="DarkTextBox" TargetType="{x:Type TextBox}">
            <Setter Property="Background" Value="#232329"/>
            <Setter Property="Foreground" Value="#ECECF1"/>
            <Setter Property="BorderBrush" Value="#3A3A45"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,7"/>
            <Setter Property="CaretBrush" Value="#ECECF1"/>
            <Setter Property="SelectionBrush" Value="#7B6CF6"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type TextBox}">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsKeyboardFocusWithin" Value="True">
                                <Setter TargetName="Bd" Property="BorderBrush" Value="#7B6CF6"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Bd" Property="Background" Value="#26262E"/>
                                <Setter Property="Foreground" Value="#71717C"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
'@

function Get-MainWindowXaml {
    return (@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Быстрые заметки"
        Width="700" Height="540" MinWidth="540" MinHeight="430"
        WindowStartupLocation="CenterScreen"
        Background="#1B1B1F" Foreground="#ECECF1"
        FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" TextOptions.TextFormattingMode="Display">
    <Window.Resources>
[[STYLES]]
        <Style x:Key="NoteButton" TargetType="{x:Type Button}">
            <Setter Property="Background" Value="#2A2A33"/>
            <Setter Property="Foreground" Value="#ECECF1"/>
            <Setter Property="BorderBrush" Value="#3A3A45"/>
            <Setter Property="Width" Value="170"/>
            <Setter Property="Height" Value="92"/>
            <Setter Property="Margin" Value="0,0,12,12"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="10">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#33333E"/>
                                <Setter TargetName="Bd" Property="BorderBrush" Value="#7B6CF6"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#3C3C49"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="#232329" BorderBrush="#2E2E38" BorderThickness="0,0,0,1" Padding="18,14">
            <DockPanel LastChildFill="True">
                <Button x:Name="BtnAdd" DockPanel.Dock="Right" Style="{StaticResource AccentButton}" Content="+  Добавить кнопку" ToolTip="Создать новую кнопку (Ctrl+N)"/>
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="📝" FontSize="22" Margin="0,0,12,0" VerticalAlignment="Center"/>
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock Text="Быстрые заметки" FontSize="16" FontWeight="SemiBold"/>
                        <TextBlock Text="Текстовые файлы в один клик" FontSize="11" Foreground="#9A9AA5"/>
                    </StackPanel>
                </StackPanel>
            </DockPanel>
        </Border>

        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Padding="18,14,6,14">
            <StackPanel>
                <TextBlock x:Name="TxtEmpty" Text="Пока нет ни одной кнопки.&#x0a;Нажмите «+ Добавить кнопку», чтобы создать первую." TextAlignment="Center" Foreground="#9A9AA5" Margin="0,70,0,0" LineHeight="22"/>
                <WrapPanel x:Name="PnlButtons"/>
            </StackPanel>
        </ScrollViewer>

        <Border Grid.Row="2" Background="#232329" BorderBrush="#2E2E38" BorderThickness="0,1,0,0" Padding="18,9">
            <DockPanel LastChildFill="True">
                <TextBlock x:Name="TxtCount" DockPanel.Dock="Right" Foreground="#9A9AA5" VerticalAlignment="Center" Margin="18,0,0,0"/>
                <TextBlock x:Name="TxtStatus" Text="Готово к работе" Foreground="#9A9AA5" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" ToolTip="{Binding Text, RelativeSource={RelativeSource Self}}"/>
            </DockPanel>
        </Border>
    </Grid>
</Window>
'@).Replace('[[STYLES]]', $script:CommonStyles)
}

function Get-EditorWindowXaml {
    return (@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Новая заметка"
        Width="720" Height="620" MinWidth="560" MinHeight="480"
        WindowStartupLocation="CenterOwner"
        Background="#1B1B1F" Foreground="#ECECF1"
        FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" TextOptions.TextFormattingMode="Display">
    <Window.Resources>
[[STYLES]]
    </Window.Resources>
    <Grid Margin="18,16,18,16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*" MinHeight="140"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <DockPanel Grid.Row="0" LastChildFill="True" VerticalAlignment="Center">
            <TextBlock DockPanel.Dock="Left" Text="Имя:" Foreground="#C7C7D1" Margin="0,0,10,0" VerticalAlignment="Center"/>
            <TextBlock DockPanel.Dock="Right" Text=".txt" Foreground="#71717C" Margin="10,0,0,0" FontFamily="Consolas" VerticalAlignment="Center"/>
            <TextBox x:Name="TbFileName" Style="{StaticResource DarkTextBox}" MaxLength="180"/>
        </DockPanel>

        <TextBlock x:Name="TxtFolder" Grid.Row="1" Margin="2,9,2,0" Foreground="#9A9AA5" FontSize="11" TextTrimming="CharacterEllipsis"/>

        <TextBox x:Name="TbText" Grid.Row="2" Style="{StaticResource DarkTextBox}" Margin="0,10,0,0"
                 AcceptsReturn="True" AcceptsTab="True" TextWrapping="Wrap" VerticalContentAlignment="Stretch"
                 VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"
                 FontFamily="Consolas" FontSize="14" Padding="12,10"/>

        <DockPanel Grid.Row="3" LastChildFill="False" Margin="2,14,2,0">
            <TextBlock DockPanel.Dock="Left" Text="Ctrl+S — сохранить   •   Esc — закрыть" Foreground="#71717C" FontSize="11" VerticalAlignment="Center"/>
            <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="BtnCancel" Style="{StaticResource GhostButton}" Content="Отмена" MinWidth="98" IsCancel="True"/>
                <Button x:Name="BtnSave" Style="{StaticResource AccentButton}" Content="Сохранить" MinWidth="118" Margin="10,0,0,0" ToolTip="Ctrl+S"/>
            </StackPanel>
        </DockPanel>
    </Grid>
</Window>
'@).Replace('[[STYLES]]', $script:CommonStyles)
}

function Get-ButtonDialogXaml {
    return (@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Новая кнопка"
        Width="520" Height="300" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner"
        Background="#1B1B1F" Foreground="#ECECF1"
        FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" TextOptions.TextFormattingMode="Display">
    <Window.Resources>
[[STYLES]]
    </Window.Resources>
    <Grid Margin="20,18,20,18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Название кнопки:" Foreground="#C7C7D1"/>
        <TextBox Grid.Row="1" x:Name="TbName" Style="{StaticResource DarkTextBox}" Margin="0,8,0,0" MaxLength="60"/>

        <TextBlock Grid.Row="2" Text="Папка, в которую будут сохраняться .txt:" Foreground="#C7C7D1" Margin="0,16,0,0"/>
        <DockPanel Grid.Row="3" LastChildFill="True" Margin="0,8,0,0">
            <Button DockPanel.Dock="Right" x:Name="BtnBrowse" Style="{StaticResource GhostButton}" Content="Обзор…" MinWidth="96" Margin="10,0,0,0"/>
            <TextBox x:Name="TbFolder" Style="{StaticResource DarkTextBox}"/>
        </DockPanel>

        <DockPanel Grid.Row="4" LastChildFill="False" Margin="0,18,0,0">
            <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="BtnCancel" Style="{StaticResource GhostButton}" Content="Отмена" MinWidth="98" IsCancel="True"/>
                <Button x:Name="BtnSave" Style="{StaticResource AccentButton}" Content="Сохранить" MinWidth="118" Margin="10,0,0,0"/>
            </StackPanel>
        </DockPanel>

        <TextBlock Grid.Row="5" x:Name="TxtError" Foreground="#F2777A" TextWrapping="Wrap" Margin="0,12,0,0" Visibility="Collapsed"/>
    </Grid>
</Window>
'@).Replace('[[STYLES]]', $script:CommonStyles)
}

# ============================== ИНТЕРФЕЙС ====================================

function ConvertFrom-XamlMarkup {
    param([string]$Markup)
    return [System.Windows.Markup.XamlReader]::Parse($Markup)
}

function Show-MessageBox {
    param(
        [System.Windows.Window]$Owner,
        [string]$Message,
        [string]$Title = $script:AppTitle,
        [string]$ButtonsText = 'OK',
        [string]$IconText = 'Information'
    )

    $button = [System.Windows.MessageBoxButton]::($ButtonsText)
    $icon   = [System.Windows.MessageBoxImage]::($IconText)
    if ($Owner) {
        return [System.Windows.MessageBox]::Show($Owner, $Message, $Title, $button, $icon)
    }
    return [System.Windows.MessageBox]::Show($Message, $Title, $button, $icon)
}

function Update-Status {
    param([string]$Text, [switch]$Good)

    if ($script:TxtStatus) {
        $script:TxtStatus.Text = $Text
        if ($Good) { $script:TxtStatus.Foreground = '#7DD487' }
        else       { $script:TxtStatus.Foreground = '#9A9AA5' }
    }
}

function Refresh-Buttons {
    $script:PnlButtons.Children.Clear()

    $buttons = @($script:Buttons)
    $script:TxtCount.Text = ('Кнопок: {0}' -f $buttons.Count)

    if ($buttons.Count -eq 0) {
        $script:TxtEmpty.Visibility   = 'Visible'
        $script:PnlButtons.Visibility = 'Collapsed'
    } else {
        $script:TxtEmpty.Visibility   = 'Collapsed'
        $script:PnlButtons.Visibility = 'Visible'
    }

    foreach ($b in $buttons) {
        if ($null -eq $b) { continue }
        [void]$script:PnlButtons.Children.Add((New-ButtonCard -Button $b))
    }
}

function New-ButtonCard {
    param([object]$Button)

    $btn = New-Object System.Windows.Controls.Button
    $btn.Style = $script:Window.TryFindResource('NoteButton')
    $btn.Tag   = $Button
    $btn.ToolTip = ("{0}`n`nЛевый клик — создать заметку.`nПравый клик — меню." -f $Button.Folder)

    $icon = New-Object System.Windows.Controls.TextBlock
    $icon.Text = '📝'
    $icon.FontSize = 24
    $icon.HorizontalAlignment = 'Center'

    $label = New-Object System.Windows.Controls.TextBlock
    $label.Text = $Button.Name
    $label.TextWrapping = 'Wrap'
    $label.TextAlignment = 'Center'
    $label.TextTrimming = 'CharacterEllipsis'
    $label.FontWeight = 'SemiBold'
    $label.Margin = '10,7,10,0'
    $label.MaxWidth = 140
    $label.MaxHeight = 42

    $stack = New-Object System.Windows.Controls.StackPanel
    $stack.VerticalAlignment = 'Center'
    [void]$stack.Children.Add($icon)
    [void]$stack.Children.Add($label)
    $btn.Content = $stack

    # Левый клик — открыть редактор заметки.
    $btn.Add_Click({
        param($sender, $e)
        Open-Editor -ButtonObj $sender.Tag
    })

    # Правый клик — контекстное меню.
    $menu = New-Object System.Windows.Controls.ContextMenu

    $miEdit = New-Object System.Windows.Controls.MenuItem
    $miEdit.Header = '✏  Изменить'
    $miEdit.Tag = $Button
    $miEdit.Add_Click({
        param($sender, $e)
        Show-ButtonDialog -Existing $sender.Tag
    })
    [void]$menu.Items.Add($miEdit)

    $miDelete = New-Object System.Windows.Controls.MenuItem
    $miDelete.Header = '🗑  Удалить'
    $miDelete.Tag = $Button
    $miDelete.Add_Click({
        param($sender, $e)
        $target = $sender.Tag
        $answer = Show-MessageBox -Owner $script:Window -Message ("Удалить кнопку «{0}»?`n(Уже созданные файлы останутся в папке.)" -f $target.Name) -Title 'Удаление кнопки' -ButtonsText 'YesNo' -IconText 'Question'
        if ($answer -ne 'Yes') { return }
        $script:Buttons = @($script:Buttons | Where-Object { $_ -ne $target })
        try {
            Save-ButtonConfig -Buttons $script:Buttons
        } catch {
            Show-MessageBox -Owner $script:Window -Message ("Не удалось сохранить настройки:`n{0}" -f $_.Exception.Message) -Icon 'Error'
            return
        }
        Refresh-Buttons
        Update-Status ('Кнопка «{0}» удалена' -f $target.Name)
    })
    [void]$menu.Items.Add($miDelete)

    [void]$menu.Items.Add((New-Object System.Windows.Controls.Separator))

    $miOpen = New-Object System.Windows.Controls.MenuItem
    $miOpen.Header = '📂  Открыть папку'
    $miOpen.Tag = $Button
    $miOpen.Add_Click({
        param($sender, $e)
        Open-FolderInExplorer -Folder (Resolve-FolderPath $sender.Tag.Folder)
    })
    [void]$menu.Items.Add($miOpen)

    $btn.ContextMenu = $menu
    return $btn
}

function Show-ButtonDialog {
    # Диалог создания/редактирования пользовательской кнопки.
    param([object]$Existing = $null)

    $dlg = ConvertFrom-XamlMarkup (Get-ButtonDialogXaml)
    $dlg.Owner = $script:Window
    $tbName    = $dlg.FindName('TbName')
    $tbFolder  = $dlg.FindName('TbFolder')
    $btnBrowse = $dlg.FindName('BtnBrowse')
    $btnSave   = $dlg.FindName('BtnSave')
    $btnCancel = $dlg.FindName('BtnCancel')
    $txtError  = $dlg.FindName('TxtError')

    if ($null -ne $Existing) {
        $dlg.Title     = 'Изменить кнопку'
        $tbName.Text   = $Existing.Name
        $tbFolder.Text = $Existing.Folder
    } else {
        $dlg.Title = 'Новая кнопка'
    }

    $btnBrowse.Add_Click({
        param($sender, $e)
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = 'Выберите папку, в которую будут сохраняться текстовые файлы'
        $fbd.ShowNewFolderButton = $true
        $current = [Environment]::ExpandEnvironmentVariables($tbFolder.Text.Trim())
        if ($current -and (Test-Path -LiteralPath $current -PathType Container)) {
            try { $fbd.SelectedPath = (Resolve-Path -LiteralPath $current).Path } catch { }
        }
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $tbFolder.Text = $fbd.SelectedPath
        }
    })

    $btnCancel.Add_Click({
        param($sender, $e)
        $dlg.Close()
    })

    $saveAction = {
        $name   = $tbName.Text.Trim()
        $folder = [Environment]::ExpandEnvironmentVariables($tbFolder.Text.Trim())

        if (-not $name) {
            $txtError.Text = 'Введите название кнопки.'
            $txtError.Visibility = 'Visible'
            return
        }
        if (-not $folder) {
            $txtError.Text = 'Укажите папку для файлов.'
            $txtError.Visibility = 'Visible'
            return
        }
        if (-not [System.IO.Path]::IsPathRooted($folder)) {
            $folder = Join-Path (Get-Location).Path $folder
        }

        if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
            $answer = Show-MessageBox -Owner $dlg -Message ("Папка не существует:`n{0}`n`nСоздать её?" -f $folder) -ButtonsText 'YesNo' -IconText 'Question'
            if ($answer -ne 'Yes') { return }
            try {
                New-Item -ItemType Directory -Path $folder -Force | Out-Null
            } catch {
                Show-MessageBox -Owner $dlg -Message ("Не удалось создать папку:`n{0}" -f $_.Exception.Message) -Icon 'Error'
                return
            }
        }

        if ($null -ne $Existing) {
            $Existing.Name   = $name
            $Existing.Folder = $folder
        } else {
            $script:Buttons = @($script:Buttons) + [pscustomobject]@{ Name = $name; Folder = $folder }
        }

        try {
            Save-ButtonConfig -Buttons $script:Buttons
        } catch {
            Show-MessageBox -Owner $dlg -Message ("Не удалось сохранить настройки:`n{0}" -f $_.Exception.Message) -Icon 'Error'
            return
        }

        $dlg.DialogResult = $true
        $dlg.Close()
    }

    $btnSave.Add_Click($saveAction)
    $tbName.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
            $e.Handled = $true
            & $saveAction
        }
    })
    $tbFolder.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
            $e.Handled = $true
            & $saveAction
        }
    })

    [void]$dlg.ShowDialog()
    Refresh-Buttons
}

function Open-Editor {
    # Окно заметки: имя по умолчанию — текущая дата, текст редактируется, файл
    # создаётся по кнопке «Сохранить».
    param([object]$ButtonObj)

    $folder = Resolve-FolderPath $ButtonObj.Folder
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        $answer = Show-MessageBox -Owner $script:Window -Message ("Папка не существует:`n{0}`n`nСоздать её?" -f $folder) -ButtonsText 'YesNo' -IconText 'Question'
        if ($answer -ne 'Yes') { return }
        try {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
        } catch {
            Show-MessageBox -Owner $script:Window -Message ("Не удалось создать папку:`n{0}" -f $_.Exception.Message) -Icon 'Error'
            return
        }
    }

    # Имя по умолчанию — текущая дата (с номером, если файл уже есть).
    $today = Get-Date -Format $script:DateFormat
    $suggested = [System.IO.Path]::GetFileNameWithoutExtension((Get-UniqueFilePath -Folder $folder -BaseName $today))

    $dlg = ConvertFrom-XamlMarkup (Get-EditorWindowXaml)
    $dlg.Title = ('Новая заметка — {0}' -f $ButtonObj.Name)
    $dlg.Owner = $script:Window

    $tbFileName = $dlg.FindName('TbFileName')
    $tbText     = $dlg.FindName('TbText')
    $txtFolder  = $dlg.FindName('TxtFolder')
    $btnSave    = $dlg.FindName('BtnSave')
    $btnCancel  = $dlg.FindName('BtnCancel')

    $tbFileName.Text = $suggested
    $txtFolder.Text  = ('Папка: {0}' -f $folder)
    $txtFolder.ToolTip = $folder

    $saveAction = {
        Save-FromEditor -Editor $dlg -Folder $folder -TbName $tbFileName -TbText $tbText
    }

    $btnSave.Add_Click($saveAction)
    $btnCancel.Add_Click({
        param($sender, $e)
        $dlg.Close()
    })

    $dlg.Add_PreviewKeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::S -and
            (($e.KeyboardDevice.Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0)) {
            $e.Handled = $true
            & $saveAction
        }
    })

    # Enter в поле имени — перейти к тексту.
    $tbFileName.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
            $e.Handled = $true
            $tbText.Focus()
        }
    })

    $dlg.Add_ContentRendered({
        param($sender, $e)
        $tbFileName.Focus()
        $tbFileName.SelectAll()
    })

    $result = $dlg.ShowDialog()
    if ($result -and $script:LastSavedPath) {
        Update-Status ('✔ Сохранено: {0}' -f $script:LastSavedPath) -Good
    }
}

function Save-FromEditor {
    param(
        [System.Windows.Window]$Editor,
        [string]$Folder,
        [object]$TbName,
        [object]$TbText
    )

    $name = ConvertTo-SafeFileName $TbName.Text
    if (-not $name) {
        Show-MessageBox -Owner $Editor -Message 'Введите имя файла.' -Icon 'Warning'
        return
    }

    $path = Join-Path $Folder ($name + $script:Extension)

    if (Test-Path -LiteralPath $path) {
        $message = ("В папке уже есть файл:`n{0}{1}`n`nДа — перезаписать его.`nНет — сохранить с номером в имени.`nОтмена — вернуться к редактированию." -f $name, $script:Extension)
        $answer = Show-MessageBox -Owner $Editor -Message $message -Title 'Файл уже существует' -ButtonsText 'YesNoCancel' -IconText 'Warning'
        if ($answer -eq 'Cancel') { return }
        if ($answer -eq 'No') {
            $path = Get-UniqueFilePath -Folder $Folder -BaseName $name
        }
    }

    try {
        $encoding = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($path, $TbText.Text, $encoding)
    } catch {
        Show-MessageBox -Owner $Editor -Message ("Не удалось сохранить файл:`n{0}" -f $_.Exception.Message) -Icon 'Error'
        return
    }

    $script:LastSavedPath = $path
    $Editor.DialogResult = $true
    $Editor.Close()
}

# ============================== САМОТЕСТЫ ====================================

function Invoke-SelfTest {
    function Assert-True {
        param([bool]$Condition, [string]$Message)
        if (-not $Condition) { throw $Message }
    }

    $passed = 0
    $failed = 0
    $failures = @()

    $checks = @(
        @{ Name = 'safe-name: removes forbidden slash'; Test = { (ConvertTo-SafeFileName 'note/test') -eq 'notetest' } },
        @{ Name = 'safe-name: trims trailing dots/spaces'; Test = { (ConvertTo-SafeFileName '  report. ') -eq 'report' } },
        @{ Name = 'safe-name: null for empty result'; Test = { $null -eq (ConvertTo-SafeFileName '///') } },
        @{ Name = 'safe-name: keeps normal text'; Test = { (ConvertTo-SafeFileName 'Встреча 2026') -eq 'Встреча 2026' } },
        @{ Name = 'date format is yyyy-MM-dd'; Test = { (Get-Date -Format $script:DateFormat) -match '^\d{4}-\d{2}-\d{2}$' } },
        @{ Name = 'unique-file: appends (2)'; Test = {
            $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('qn-test-' + [guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $tmp | Out-Null
            try {
                Set-Content -LiteralPath (Join-Path $tmp '2026-01-01.txt') -Value 'x'
                $p = Get-UniqueFilePath -Folder $tmp -BaseName '2026-01-01'
                ([System.IO.Path]::GetFileName($p) -eq '2026-01-01 (2).txt')
            } finally { Remove-Item -LiteralPath $tmp -Recurse -Force }
        } },
        @{ Name = 'unique-file: free name returned as-is'; Test = {
            $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('qn-test-' + [guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $tmp | Out-Null
            try {
                $p = Get-UniqueFilePath -Folder $tmp -BaseName 'free'
                ([System.IO.Path]::GetFileName($p) -eq 'free.txt')
            } finally { Remove-Item -LiteralPath $tmp -Recurse -Force }
        } },
        @{ Name = 'config: json round-trip'; Test = {
            $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('qn-test-' + [guid]::NewGuid().ToString('N') + '.json')
            try {
                $src = @(
                    [pscustomobject]@{ Name = 'Работа'; Folder = 'C:\Notes' },
                    [pscustomobject]@{ Name = 'Дом';   Folder = 'D:\Личное' }
                )
                Save-ButtonConfig -Buttons $src -Path $tmp
                $restored = @(Read-ButtonConfig -Path $tmp)
                Assert-True ($restored.Count -eq 2) 'expected 2 buttons'
                Assert-True ($restored[0].Name -eq 'Работа' -and $restored[0].Folder -eq 'C:\Notes') 'first button mismatch'
                Assert-True ($restored[1].Name -eq 'Дом' -and $restored[1].Folder -eq 'D:\Личное') 'second button mismatch'
                $true
            } finally { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force } }
        } },
        @{ Name = 'config: empty list round-trip'; Test = {
            $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('qn-test-' + [guid]::NewGuid().ToString('N') + '.json')
            try {
                Save-ButtonConfig -Buttons @() -Path $tmp
                $restored = @(Read-ButtonConfig -Path $tmp)
                Assert-True ($restored.Count -eq 0) 'expected empty list'
                $true
            } finally { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force } }
        } },
        @{ Name = 'config: missing file returns empty'; Test = {
            $restored = @(Read-ButtonConfig -Path (Join-Path ([System.IO.Path]::GetTempPath()) ('qn-missing-' + [guid]::NewGuid().ToString('N') + '.json')))
            ($restored.Count -eq 0)
        } },
        @{ Name = 'main XAML is well-formed'; Test = {
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml((Get-MainWindowXaml))
            $true
        } },
        @{ Name = 'editor XAML is well-formed'; Test = {
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml((Get-EditorWindowXaml))
            $true
        } },
        @{ Name = 'button dialog XAML is well-formed'; Test = {
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml((Get-ButtonDialogXaml))
            $true
        } }
    )

    Write-Host 'QuickNotes self-test:'
    foreach ($check in $checks) {
        $ok = $false
        $err = $null
        try { $ok = [bool](& $check.Test) } catch { $err = $_.Exception.Message }
        if ($ok) {
            $passed++
            Write-Host ('  [OK]   ' + $check.Name)
        } else {
            $failed++
            if (-not $err) { $err = 'condition is false' }
            $failures += ('{0}: {1}' -f $check.Name, $err)
            Write-Host ('  [FAIL] ' + $check.Name + ' :: ' + $err)
        }
    }

    Write-Host ''
    Write-Host ('Passed: {0}, failed: {1}' -f $passed, $failed)
    if ($failed -gt 0) {
        throw ('Self-test failed: ' + ($failures -join '; '))
    }
}

# ================================ ЗАПУСК =====================================

if ($SelfTest) {
    try {
        Invoke-SelfTest
        exit 0
    } catch {
        Write-Error $_
        exit 1
    }
}

try {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Windows.Forms

    $script:DataDirectory  = $null
    $script:LastSavedPath  = $null
    $script:Buttons        = @(Read-ButtonConfig)
    $script:FirstRun       = -not (Test-Path -LiteralPath (Get-ConfigFilePath))

    $window = ConvertFrom-XamlMarkup (Get-MainWindowXaml)
    $script:Window     = $window
    $script:BtnAdd     = $window.FindName('BtnAdd')
    $script:TxtEmpty   = $window.FindName('TxtEmpty')
    $script:PnlButtons = $window.FindName('PnlButtons')
    $script:TxtStatus  = $window.FindName('TxtStatus')
    $script:TxtCount   = $window.FindName('TxtCount')

    Update-Status ('Настройки: {0}' -f (Get-ConfigFilePath))

    $script:BtnAdd.Add_Click({
        param($sender, $e)
        Show-ButtonDialog
    })

    $window.Add_PreviewKeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::N -and
            (($e.KeyboardDevice.Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0)) {
            $e.Handled = $true
            Show-ButtonDialog
        }
    })

    # При первом запуске сразу предлагаем создать кнопку.
    $window.Add_ContentRendered({
        param($sender, $e)
        if ($script:FirstRun) {
            $script:FirstRun = $false
            Show-ButtonDialog
        }
    })

    Refresh-Buttons
    [void]$window.ShowDialog()
} catch {
    $message = ('Ошибка: {0}' -f $_.Exception.Message)
    try {
        Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
        [void][System.Windows.MessageBox]::Show($message, $script:AppTitle, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    } catch {
        Write-Error $message
    }
    exit 1
}
