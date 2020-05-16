# rundll32-printui-gui-Druckverwaltung-Win10 
This is a script to Manage the Windows 10 Network Printer with an easy GUI. It is designed for your Users.

You only need to set the variables at the beginning to your printserver like "Printserver_01".
You don't need to use the FQDN. You can use more than one Printserver. If you use more, please follow this:

After line 7:  
$var.prints1 = "Printserver1.4styler.lokal"  
$var.prints2 = "Printserver2.4styler.lokal"  
$var.prints3 = "Printserver3"  
$var.prints4 = "Some_ohter_Server"  
  
below line 103:  
  
$a = $(Get-Printer -ComputerName $var.prints1)  
or:  
$a = $(Get-Printer -ComputerName $var.prints1) + $(Get-Printer -ComputerName $var.prints2)  
or:  
$a = $(Get-Printer -ComputerName $var.prints1) + $(Get-Printer -ComputerName $var.prints2) + $(Get-Printer -ComputerName $var.prints3)  
or:  
 $a = $(Get-Printer -ComputerName $var.prints1) + $(Get-Printer -ComputerName $var.prints2) + $(Get-Printer -ComputerName $var.prints3) + $(Get-Printer -ComputerName $var.prints4)  
 or....  
 and so on.  
  
  
I needed a solution for our WS 2016 TS Farm because the standard Win 2016 Printer GUI has a lot of bugs. With that as
an alternative Form it will work like a charm. I have also seen the same Printer Issuses on Windows 10.

I hope it will help you.
