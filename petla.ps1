$baseDir = 'C:\do III stopnia\badania\przeliczenia\' #katalg plikow pomiarowych i pliku docelowego

$x = 1 #numer pliku z pomiarami
$cz = 1 #numer czujnika z pliku z pomiarami
$pomiar = 'pomiar '#nazwa pliku z pomiarami string - poczatek
$czN = '_cz' #nazwa pliku z pomiarami string - koncowka

#Write-Host ($baseDir + $pomiar + "$x".PadLeft(3, '0') + $czN + "$cz".PadLeft(2, '0')) #- tutaj sprawdzalem jak zapisuje mi się nazwa pliku - jest ok

    for ($x = 1; $x -le 2 ; $x++) { # petla ktora otwiera kolejny pomiar czyli zmieni 001, 002 etc (na razie od 1 do 2)
 "$x = " + (1 + $x) 
for ($cz = 1; $cz -le 3 ; $cz++) { # petla ktora otwiera kolejny czujnik w danym pomiarze czyli zmieni cz01, cz02 i cz03 - to zawsze jest od 1 do 3
 "$cz = " + (1 + $cz) 

$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open($baseDir + $pomiar + "$x".PadLeft(3, '0') + $czN + "$cz".PadLeft(2, '0'))
$Worksheet = $Workbook.WorkSheets.item(“Arkusz1”)
$worksheet.activate()
$excel.Visible = $true
}
    }