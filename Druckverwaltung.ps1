#Erstellen des Hashtables. Hauptvariable
$var = [hashtable]::Synchronized(@{})

#Hier die Druckserver angeben. Für 2 Stück vorbereitet:
$var.prints1 = "$env:COMPUTERNAME"
#$var.prints2 = "$env:COMPUTERNAME"


# Wenn das Script bereits läuft, schliessen
$runningtest = $($(get-wmiobject win32_process -Filter "name = 'powershell.exe'").commandline | select-string -Pattern "Druckerverwaltung.ps1").count
if ($runningtest -ge 2 ) { exit }

# ---------------------------------------------------------------------------------------------------
#Ausblenden des Powershell Fensters.
# ---------------------------------------------------------------------------------------------------
Add-Type -Name win -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);' -Namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle,0)

#Laden der Benötigten Assemblies
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName presentationframework

#Timer für die aktualisierung des Fensters.
$var.timer = $null
$var.timer = New-Object System.Windows.Forms.Timer

function mainwindow {
#Erstellen des Hauptfensters. Achtung formatierung muss so bleiben.
[xml]$var.maingui = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Druckverwaltung" Height="535" Width="800">
    <Grid>
        <ListBox x:Name="lb1" HorizontalAlignment="Left" Margin="10,69,0,0" VerticalAlignment="Top" Width="300" Height="425" SelectionMode="Extended"/>
        <ListBox x:Name="lb2" HorizontalAlignment="Left" Height="425" Margin="482,69,0,0" VerticalAlignment="Top" Width="300" SelectionMode="Extended"/>
        <TextBox x:Name="tb1" HorizontalAlignment="Left" Height="23" Margin="10,41,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="300"/>
        <Label x:Name="l1" Content="Filter:" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="tb2" HorizontalAlignment="Left" Height="23" Margin="482,41,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="300"/>
        <Label x:Name="l2" Content="Filter:" HorizontalAlignment="Left" Margin="482,10,0,0" VerticalAlignment="Top"/>
        <Button x:Name="b1" Content="           Hinzufügen-----------&gt;" HorizontalAlignment="Left" Margin="315,69,0,0" VerticalAlignment="Top" Width="162"/>
        <Button x:Name="b2" Content="&lt;-----------Löschen                " HorizontalAlignment="Left" Margin="315,94,0,0" VerticalAlignment="Top" Width="162"/>
        <Button x:Name="b3" Content="&lt;--------Als Standard--------&gt;" HorizontalAlignment="Left" Margin="315,119,0,0" VerticalAlignment="Top" Width="162"/>
        <Button x:Name="b4" Content="Geräte und Drucker" HorizontalAlignment="Left" Margin="315,41,0,0" VerticalAlignment="Top" Width="162" Height="23"/>
        <Button x:Name="b5" Content="&lt;-----Testseite Drucken          " HorizontalAlignment="Left" Margin="315,144,0,0" VerticalAlignment="Top" Width="162"/>
        <TextBlock x:Name="tbl1" HorizontalAlignment="Left" Margin="315,180,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Height="325" Width="162"/>
    </Grid>
</Window>

"@
#Einbinden der Objekte aus der GUI in veränderbare Variablen um diese nacher manipulieren zu können.
    $var.MainWindow=[Windows.Markup.XamlReader]::Load($(New-Object System.Xml.XmlNodeReader $var.maingui))
    $var.lb1 = $var.MainWindow.FindName('lb1')
    $var.lb2 = $var.MainWindow.FindName('lb2')
    $var.tb1 = $var.MainWindow.FindName('tb1')
    $var.tb2 = $var.MainWindow.FindName('tb2')
    $var.b1 = $var.MainWindow.FindName('b1')
    $var.b2 = $var.MainWindow.FindName('b2')
    $var.b3 = $var.MainWindow.FindName('b3')
    $var.b4 = $var.MainWindow.FindName('b4')
    $var.b5 = $var.MainWindow.FindName('b5')
    $var.tbl1 = $var.MainWindow.FindName('tbl1')

    $var.MainWindow.add_MouseLeftButtonDown({
        $var.lb1.UnselectAll()
        $var.lb2.UnselectAll()
    })
#############################################################
#############################################################
    #ListBox1 vorbereiten:
    function lb1_pre{
        $var.actprinter = @()
        $var.actprinter += $(Get-Printer).Name | Sort | where {$_ -match "\\"}
        $var.lb1.ItemsSource = $var.actprinter
    }
    lb1_pre

    $var.lb1.add_MouseLeftButtonDown({
        $var.lb1.UnselectAll()
    })

    $var.lb1.add_GotFocus({
        $var.lb2.UnselectAll()
    })

    $var.tb1.add_TextChanged({
        $var.search = @()
        $var.search += $($var.actprinter | ? { $_ -like "*$($var.tb1.Text)*" })
        $var.lb1.ItemsSource = @($var.search | Sort)
    })
############################################################# 
#############################################################
    #ListBox2 vorbereiten:
    function lb2_pre{
        # 1 Server:
        $a = $(Get-Printer -ComputerName $var.prints1)
        # 2 Server:
        #$a = $(Get-Printer -ComputerName $var.prints1) + $(Get-Printer -ComputerName $var.prints2)
        # 3 Server (neue Variable am Anfang vom Skript nicht vergessen):
        #$a = $(Get-Printer -ComputerName $var.prints1) + $(Get-Printer -ComputerName $var.prints2) + $(Get-Printer -ComputerName $var.prints3)
        # n Server: usw...
        $var.servprinter = @()
        $a | Sort | foreach {
            $var.servprinter += $("\\" + $_.ComputerName + "\" + $_.Name)
        }
        $var.lb2.ItemsSource = $var.servprinter | Sort
    }
    lb2_pre

    $var.lb2.add_MouseLeftButtonDown({
        $var.lb2.UnselectAll()
    })

    $var.lb2.add_GotFocus({
        $var.lb1.UnselectAll()
    })

    $var.tb2.add_TextChanged({
        $var.search = @()
        $var.search += $($var.servprinter | ? { $_ -like "*$($var.tb2.Text)*" })
        $var.lb2.ItemsSource = @($var.search | Sort)
    })
#############################################################
#############################################################
$var.b1.add_Click({
    if (!$var.lb2.SelectedItems){
        [System.Windows.Forms.MessageBox]::Show("Bitte rechts einen Drucker wählen","Fehlende Auswahl",0)
    }else{
        $var.lb2.SelectedItems | foreach {
            rundll32 printui.dll,PrintUIEntry /in /n "$($_)"
        }
    }
})

#############################################################
#############################################################
$var.b2.add_Click({
    if (!$var.lb1.SelectedItems){
        [System.Windows.Forms.MessageBox]::Show("Bitte links einen Drucker wählen","Fehlende Auswahl",0)
    }else{
        $var.lb1.SelectedItems | foreach {
            rundll32 printui.dll,PrintUIEntry /dn /n "$($_)"
        }
        
    }
})

#############################################################
#############################################################
$var.b3.add_Click({
    if ($var.lb1.SelectedItem){
        rundll32 printui.dll,PrintUIEntry /in /y /n "$($var.lb1.SelectedItem)"
    }else{[System.Windows.Forms.MessageBox]::Show("Bitte links einen Drucker wählen","Fehlende Auswahl",0)}
})

#############################################################
#############################################################
$var.b4.add_Click({
    control printers
})

$var.b5.add_Click({
    if ($var.lb1.SelectedItem){
        $Printers = Get-CimInstance -ClassName Win32_Printer
        $Printer = $Printers | where {$_.Name -eq $var.lb1.SelectedItem}
        Invoke-CimMethod -InputObject $Printer -MethodName PrintTestPage
    }else{[System.Windows.Forms.MessageBox]::Show("Bitte links einen Drucker wählen","Fehlende Auswahl",0)}
})
#############################################################
#############################################################
$var.tbl1.Text = 'Hilfe:
Über Hinzufügen wird bzw. werden die auf der rechten Seite markierten Drucker installiert.

Über Löschen wird der bzw. die markierten Drucker auf der linken Seite gelöscht.

Mit dem "Als Standard" Button wird der Markierte Drucker der Linken, bzw. rechten Seite als Standard gesetzt.
Ist der gewählte Drucker nicht installiert, wird dieser zusätzlich Installiert.'
#############################################################

#Anzeigen der GUI und offenhalten des Scriptes bis die GUI geschlossen wird.
    $var.MainWindow.ShowDialog() | Out-Null
}

$var.timer.add_Tick({
    function lb3_pre{
        $var.actprinter = @()
        $var.actprinter += $(Get-Printer).Name | Sort | where {$_ -match "\\"}
        $var.lb1.ItemsSource = $var.actprinter
    }
    if (!$($var.tb1.Text)){
        lb3_pre
    }
})
$var.timer.Interval = 500
$var.timer.Start()

mainwindow

$var.mainwindow.Close() | Out-Null
$var.timer.Stop()
$var.timer.Dispose()
