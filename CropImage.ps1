Using Module "C:\Users\vital\Repositories\Utilities\ColorWriter.psm1"

<#
Script schneidet Bildausschnitte aus vorgegebenen Bildern aus.

Version:
    0.5:    Auflösung der Ausgabedatei auf 300dpi erhöht. Defaultwert ist/war 96dpi
    0.4:    Prozentuale Skalierung (Verkleinerung oder Vergrößerung) hinzugefügt.
    0.3:    Variable Größenangabe für die Output-Datei implementiert. Davor entsprach die Ausgabedatei immer dem Ausschnitt aus dem originalen Bild.
    0.2:    Bunte Ausgabe von Text und Parameter implementiert.
    0.1:    Erste Version von CropImage

Useage:
    . ".\CropImage.ps1"
    CropImage -s [source] -d [dest] -w [width] -h [height] => -x = 0, -y = 0
    CropImage -s [source] -d [dest] -x [X-coordinate] -y [Y-coordinate] -w [width] -h [height]
    CropImage -s [source] -d [dest] -x [X-coordinate] -y [Y-coordinate] -w [width] -h [height] -ow [output width] -oh [output height]
    CropImage -s [source] -d [dest] -p [percent]
    CropImage -s "C:\Users\Vitali\Pictures\Screnshots\Screenshot_2023-04-18-08-53-44-95.jpg" -d "C:\Users\Vitali\Pictures\Screnshots\Cuttings\Screenshot_2023-04-18-08-53-44-95-05.jpg" -x 0 -y 12464 -w 1440 -h 3216
#>

Function CropImage {
    Param(
        # Specifies the image source path. The source image should exist, otherwise the program will be terminated throw an error.
        [Parameter(Mandatory)]
        [Alias('source', 'src', 's')]
        [String] $strSourcePath,
        # Specifies the imgage destination path. The destination image shouldn't already exists, otherwise the program will be terminated throw an error.
        [Parameter(Mandatory)]
        [Alias('dest', 'dst', 'd')]
        [String] $strDestPath,
        # Specifies the start x-coordinate for the crop. X-coordinate should be within the size of the source file.
        [Alias('x')]
        [ValidateRange(0, 10000)]
        [Int] $intX,
        # Specifies the start y-coordinate for the crop. Y-coordinate should be within the size of the source file.
        [Alias('y')]
        [ValidateRange(0, 10000)]
        [Int] $intY,
        # Specifies the width for the crop. X-coordinate +  width should be within the size of the source file.
        [Alias('width', 'w')]
        [ValidateRange(1, 10000)]
        [Int] $intWidth,
        # Specifies the height for the crop. Y-coordinate + height should be within the size of the source file.
        [Alias('height', 'h')]
        [ValidateRange(1, 10000)]
        [Int] $intHeight,
        # Specifies the width of the output picture.
        [Alias('outputwidth', 'ow')]
        [ValidateRange(1, 10000)]
        [Int] $intDestWidth,
        # Specifies the height of the output picture.
        [Alias('outputheight', 'oh')]
        [ValidateRange(1, 10000)]
        [Int] $intDestHeight,
        # Specifies the percent of the output picture.
        [Alias('percent', 'p')]
        [ValidateRange(0.01, 10.0)]
        [double] $dblDestPercent,
        # Specifies the switch that, when activated, does not overwrite existing files.
        [Alias('save')]
        [Switch] $swiSaveMode
    )

    Begin{
        [String] $strTitle = "CropImage"
        [String] $strVersion = "v0.5"

        [ColorWriter]::Write(@("$strTitle ", "Blue", "$strVersion ", "Green", "started", "White"))

        # Check parameter
        If(-not (Test-Path $strSourcePath)){
            Throw "Source image doesn't exist: $strSourcePath"
        }
        If($swiSaveMode -and (Test-Path $strDestPath)){
            Throw "Destination image already exists: $strDestPath"
        }

        # Lade die System.Drawing-Assembly
        Add-Type -AssemblyName System.Drawing 
        # Add-Type -AssemblyName System.Drawing.Rectangle

        # Bitmap-Objekt aus der JPG-Datei erstellen
        $btmSourceBitmap = New-Object System.Drawing.Bitmap($strSourcePath)

        [Int] $intDestDPI = 300

        # Percent specified => Calculate $intDestWidth and intDestHeight
        If($dblDestPercent -ne 0){
            # Es ist sehr interessant, dass die obige Definition des maximal möglichen Wertes von 10000 auch hier eingehalten wird und ein Fehler geworfen wird.
            # Der Fehler lässt sich jedoch nicht abfangen (siehe unten).
            If($btmSourceBitmap.Width * $dblDestPercent -gt 10000){
                Throw "The percentage parameter $dblDestPercent exceeds the maximum allowed width value of 10000 to $($btmSourceBitmap.Width * $dblDestPercent)."
            }
            If($btmSourceBitmap.Height * $dblDestPercent -gt 10000){
                Throw "The percentage parameter $dblDestPercent exceeds the maximum allowed height value of 10000 to $($btmSourceBitmap.Height * $dblDestPercent)."
            }

            $intWidth = $btmSourceBitmap.Width
            $intHeight = $btmSourceBitmap.Height
            $intDestWidth = [Int] $btmSourceBitmap.Width * $dblDestPercent 
            $intDestHeight = [Int] $btmSourceBitmap.Height * $dblDestPercent
            <# 
                try {
                    $intDestWidth = [Int] $intWidth * $dblDestPercent 
                }
                catch [System.Management.Automation.ParameterBindingValidationException] {
                    Throw "The percentage parameter $dblDestPercent exceeds the maximum allowed width value of 10000 to $($intWidth * $dblDestPercent)."
                }
            #>
        }
        
        # Größenangabe für den Crop (Auschnitt) für die Input-Datei wurde weggelassen.
        # => Kalkuliere die Crop-Größe aus der Input-Datei
        If($intHeight -eq 0 -and $intWidth -eq 0){
            $intWidth = $btmSourceBitmap.Width
            $intHeight = $btmSourceBitmap.Height
        }
        ElseIf($intWidth -eq 0){
            $intWidth = [Int] $btmSourceBitmap.Width / $btmSourceBitmap.Height * $intHeight
        }
        ElseIf($intHeight -eq 0){
            $intHeight = [Int] $btmSourceBitmap.Width / $btmSourceBitmap.Height * $intWidth
        }

        <# Possible Combination: 
        $intDestWidth = 0 & $intDestHeight = 0 => Take the size of the crop
        $intDestWidth = w & $intDestHeight = 0 => Calculate $intDestHeight 
        $intDestWidth = 0 & $intDestHeight = h => Calculate $intDestWidth 
        $intDestWidth = w & $intDestHeight = h => Nothing will be taken
        #>

        # Größenangabe für die Output-Datei wurde weggelassen.
        If($intDestWidth -eq 0 -and $intDestHeight -eq 0){
            $intDestWidth = $intWidth
            $intDestHeight = $intHeight
        }
        ElseIf($intDestHeight -eq 0){
            $intDestHeight = [Int] $intHeight / $intWidth * $intDestWidth
        }
        ElseIf($intDestWidth -eq 0){
            $intDestWidth = [Int] $intWidth / $intHeight * $intDestHeight
        }

        # Parameter output
        $arrColoredString = @(
            "Source path: ", "White", "$strSourcePath`r`n", "Yellow",
            "Source picture parameters (width, height, h-dpi, v-dpi): ", "White", 
            "$($btmSourceBitmap.Width), $($btmSourceBitmap.Height), $($btmSourceBitmap.HorizontalResolution), $($btmSourceBitmap.VerticalResolution)`r`n", "Green",
            "Crop rectangle (x, y, w, h): ", "White", "$intX, $intY, $intWidth, $intHeight`r`n", "Green",
            "Destination path: ", "White", "$strDestPath`r`n", "Yellow",
            "Destination picture parameters (width, height, h-dpi, v-dpi): ", "White", 
            "$intDestWidth, $intDestHeight, $intDestDPI, $intDestDPI", "Green"
        )
        [ColorWriter]::Write($arrColoredString)

        # Interrupt execution
        # Throw "X-coordinate should be within the size of the source file: $intX >= $($btmSourceBitmap.Width)!"

        # Check parameter
        If($intX -ge $btmSourceBitmap.Width){
            Throw "X-coordinate should be within the size of the source file: $intX >= $($btmSourceBitmap.Width)!"
        }
        If($intY -ge $btmSourceBitmap.Height){
            Throw "Y-coordinate should be within the size of the source file: $intY >= $($btmSourceBitmap.Height)!"
        }
        If(($intX + $intWidth) -gt $btmSourceBitmap.Width){
            Throw "X-coordinate + width should be within the size of the source file: ($intX + $intWidth) >= $($btmSourceBitmap.Width)!"
        }
        If(($intY + $intHeight) -gt $btmSourceBitmap.Height){
            Throw "Y-coordinate + height should be within the size of the source file: ($intY + $intHeight) >= $($btmSourceBitmap.Height)!"
        }
    }
    Process{
       # Bitmap-Objekt mit der Zielgröße erstellen
        $btmDestBitmap = New-Object System.Drawing.Bitmap($intDestWidth, $intDestHeight)

        # Obwohl es keinen Speicherplatzunterschied zwischen 96dpi (default) und 300dpi gibt, setze ich es hier trotzdem auf 300dpi,
        # weil alle meine zugeschnittenen Bilder für die Buchführung eine Auflösung von 300dpi aufweisen.
        $btmDestBitmap.SetResolution($intDestDPI, $intDestDPI)

        # Grafikobjekt aus dem Bitmap-Objekt erstellen
        $grpDestGraphics = [System.Drawing.Graphics]::FromImage($btmDestBitmap)

        # Grafikobjekt so konfigurieren, dass es das Bild mit der Zielgröße zuschneidet
        $grpDestGraphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $grpDestGraphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $grpDestGraphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $grpDestGraphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality

        # Bild mit der Zielgröße auf das Grafikobjekt zeichnen
        $reaSourceRect = New-Object Drawing.Rectangle($intX, $intY, $intWidth, $intHeight)
        $reaDestRect = New-Object Drawing.Rectangle(0, 0, $intDestWidth, $intDestHeight)

        $grpDestGraphics.DrawImage($btmSourceBitmap, $reaDestRect, $reaSourceRect, [System.Drawing.GraphicsUnit]::Pixel)
     
        $btmDestBitmap.Save($strDestPath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    }
    End{
        
        [ColorWriter]::Write(@("Destination file created, dimension (width, heigh): ", "White", 
        "($($btmDestBitmap.Width), $($btmDestBitmap.Height))", "Green"))

        # Ziel-Grafikobjekt freigeben
        $grpDestGraphics.Dispose()

        # Ziel-Bitmap-Objekt freigeben
        $btmDestBitmap.Dispose()

        # Quell-Bitmap-Objekt freigeben
        $btmSourceBitmap.Dispose()

        [ColorWriter]::Write(@("$strTitle ", "Blue", "$strVersion ", "Green", "finished", "White"))
    }
}



## Benzin-Rechungen ##
# Benzin-Rechnungen sollten auf 936 x 2800 zugeschnitten werden:
[String] $strSourcePath = "C:\Users\vital\Documents\IMG_20260125_0003.jpg"
[String] $strDestPath = "C:\Users\vital\Documents\Finanzen\2025\04 Rohdaten\Sonstige Rechnungen\Scan 11.jpg"

CropImage -s $strSourcePath -d $strDestPath -x 0 -y 0 -w 936 -h 2800

#Parktickets / Parkrechnungen
# CropImage -s $strSourcePath -d $strDestPath -x 0 -y 0 -w 684 -h 2200


<#
# Generated with: C:\Users\vital\Progs\CropImage\Comman String Generator.xlsx
CropImage -s "C:\Users\vital\Documents\IMG_20260125_0005.jpg" -d "C:\Users\vital\Documents\Finanzen\2025\04 Rohdaten\Benzin Rechnungen\Scan 31.jpg" -x 0 -y 0 -w 936 -h 2800
CropImage -s "C:\Users\vital\Documents\IMG_20260125_0006.jpg" -d "C:\Users\vital\Documents\Finanzen\2025\04 Rohdaten\Benzin Rechnungen\Scan 32.jpg" -x 0 -y 0 -w 936 -h 2800
CropImage -s "C:\Users\vital\Documents\IMG_20260125_0007.jpg" -d "C:\Users\vital\Documents\Finanzen\2025\04 Rohdaten\Benzin Rechnungen\Scan 33.jpg" -x 0 -y 0 -w 936 -h 2800
CropImage -s "C:\Users\vital\Documents\IMG_20260125_0008.jpg" -d "C:\Users\vital\Documents\Finanzen\2025\04 Rohdaten\Benzin Rechnungen\Scan 34.jpg" -x 0 -y 0 -w 936 -h 2800
CropImage -s "C:\Users\vital\Documents\IMG_20260125_0009.jpg" -d "C:\Users\vital\Documents\Finanzen\2025\04 Rohdaten\Benzin Rechnungen\Scan 35.jpg" -x 0 -y 0 -w 936 -h 2800
CropImage -s "C:\Users\vital\Documents\IMG_20260125_0010.jpg" -d "C:\Users\vital\Documents\Finanzen\2025\04 Rohdaten\Benzin Rechnungen\Scan 36.jpg" -x 0 -y 0 -w 936 -h 2800
CropImage -s "C:\Users\vital\Documents\IMG_20260125_0011.jpg" -d "C:\Users\vital\Documents\Finanzen\2025\04 Rohdaten\Benzin Rechnungen\Scan 37.jpg" -x 0 -y 0 -w 936 -h 2800
CropImage -s "C:\Users\vital\Documents\IMG_20260125_0012.jpg" -d "C:\Users\vital\Documents\Finanzen\2025\04 Rohdaten\Benzin Rechnungen\Scan 38.jpg" -x 0 -y 0 -w 936 -h 2800
#>


# Überweisung
# CropImage -s "$strSourcePath\014.jpg" -d "$strDestPath\20230109 �berweisung an Frommer.jpg" -x 0 -y 0 -w 1800 -h 1250
# CropImage -s "$strSourcePath\015.jpg" -d "$strDestPath\20230109 �berweisung an Frommer.jpg" -x 0 -y 0 -w 1800 -h 1250

# Kontoauszug, GEZ
# CropImage -s $strSourcePath -d $strDestPath -x 0 -y 0 -w 2480 -h 1250

## Lange Screenshots ##
# Auflösung eines langen Screenshots: 1440 * 30150
# Auflösung eines normalen Screenshots: 1440 * 3216
#[String] $strSourcePath = "C:\Users\vital\Progs\CropImage\Screenshots\Screenshot_2023-04-18-07-46-22-67.jpg"
#[String] $strDestPath = "C:\Users\vital\Progs\CropImage\Screenshots\Cut\Test.jpg"
#CropImage -s $strSourcePath -d $strDestPath -x 0 -y 0 -w 1440 -h 3216

# Bewirtungsbeleg / Din A4-Scan vom alten Canon-Drucker
# CropImage -s "$strSourcePath\003.jpg" -d "$strDestPath\20240110 Zahlungsverpflichtung, unterschrieben.jpg" -x 0 -y 0 -w 2492 -h 3504

#[String] $strSourcePath = "C:\Users\vital\Music\Sophias Hörbücher\Was ist was\input\Bienen und Wespen - Im Gespräch mit einer Drohne.jpeg" 
#[String] $strDestPath = "C:\Users\vital\Music\Sophias Hörbücher\Was ist was\output\Bienen und Wespen - Im Gespräch mit einer Drohne.jpeg"
#CropImage -s $strSourcePath -d $strDestPath -p 0.35
