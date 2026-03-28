<#
Script gibt übergebenen Text-Array in übergebener Farben-Array aus.

Improvements:
    

Version:
    v0.1:   Erste Version von ColorWriter.
    v0.2:   Texte und Farben können jetzt in einem Array übergeben werden.
    v0.3:   Erste Version von WriteTable.

Useage:
    Using Module "C:\Users\vital\Progs\Utilities\ColorWriter.psm1"
    
    [ColorWriter]::Test()

    $strFilePath = "C:\Users\Vitali\Pictures\Meine Scans\Scan XX.jpg"
    [ColorWriter]::WriteOpenFile($strFilePath)
#>

class ColorWriter {
    <#
   .Description
   Function static [void] Write prints the content of the passed $arrString with the passed $arrColor colors.

   [String[]] $arrString: array of strings.
   [String[]] $arrColor: array of colors.
   #>
   static [void] Write([String[]] $arrString, [String[]] $arrColor){
       For ($i=0; $i -lt $arrString.length; $i++){
              Write-Host $arrString[$i] -ForegroundColor $arrColor[$i] -NoNewline
       }
       Write-Host
   }
   <#
   .Description
   Function static [void] Write prints the content of the passed $arrColoredString.

   [String[]] $arrColoredString: array of string followed by color.
   #>
   static [void] Write([String[]] $arrColoredString){
       If($arrColoredString.length % 2 -ne 0){
           Throw "Passed string array should contain text and color in alternation, array-length: $($arrColoredString.length)"
       }
       For ($i=0; $i -lt $arrColoredString.length; $i=$i+2){
              Write-Host $arrColoredString[$i] -ForegroundColor $arrColoredString[$i+1] -NoNewline
       }
       Write-Host
   }
   <#
   .Description
   Example method of how to use the static Writer methode.

   [String[]] $arrString: array of strings.
   [String[]] $arrColor: array of colors.
   #>

   static [void] WriteOpenFile($strFilePath){
       $arrPrintStrings = @(
           "Open the ",
           $(If($strFilePath.ToLower().EndsWith(".jpg")){"picture "}
           ElseIf($strFilePath.ToLower().EndsWith(".pdf")){"pdf "}
           Else{""}),
           "file: ",
           "$strFilePath")

       $arrPrintColors = @(
           "White",
           "Green",
           "White",
           "Yellow")

       [ColorWriter]::Write($arrPrintStrings, $arrPrintColors)
   }
    static [void] WriteTable([String[][]] $mtxString, [String] $strTitleColor, [String] $strEntryColor){

        # Überprüfung des Formats der Matrix
        # Alle enthaltenen Array, sollten die gleiche Länge besitzen
        [Int] $intCount = $mtxString[0].Count
        ForEach($arrString in $mtxString){
            If($intCount -ne $arrString.Count){
                Throw "ColorWriter.WriteTable(): Length of row arrays is not the same. $intCount vs $($arrString.Count)."
            }
        }

        # Bestimmung der maximalen Breite jeder Spalte, abh�ngig vom Eintrag
        [Int] $intLength = 0
        [Int[]] $arrLength = @()
        
        For ($i=0; $i -lt ($mtxString[0].length-1); $i++){
            For ($j=0; $j -lt $mtxString.length; $j++){
                If($mtxString[$j][$i].Length -gt $intLength){
                    $intLength = $mtxString[$j][$i].Length + 2
                }
            }
            $arrLength += $intLength
            $intLength = 0
        }

        # Ausgabe der Tabellen�berschrift 
        For ($i=0; $i -lt $mtxString[0].length; $i++){
            If($i -lt $mtxString[0].length-1){
                Write-Host $mtxString[0][$i](" " * ($arrLength[$i] - $mtxString[0][$i].Length)) -ForegroundColor $strTitleColor -NoNewline
            }
            Else{# Der letzte Eintrag hat keine Abst�nde, au�erdem hat er eine Newline
                Write-Host $mtxString[0][$i] -ForegroundColor $strTitleColor
            }
            
        } 

        # Ausgabe des Tabelleninhalts
        For ($i=1; $i -lt $mtxString.length; $i++){
            For ($j=0; $j -lt $mtxString[$i].length; $j++){
                If($j -lt $mtxString[0].length-1){
                    Write-Host $mtxString[$i][$j](" " * ($arrLength[$j] - $mtxString[$i][$j].Length)) -ForegroundColor $strEntryColor -NoNewline
                }
                Else{# Der letzte Eintrag hat keine Abst�nde, au�erdem hat er eine Newline
                    Write-Host $mtxString[$i][$j] -ForegroundColor $strEntryColor
                }
            }
        }
    }
    static [void] Test(){
        $strSourcePath = "C:\Users\Vitali\Pictures\Screnshots\Screenshot_2023-04-18-08-53-44-95.jpg"
            $arrColoredString = @(
            "-=: Input parameter :=-`r`n", "White",
            "Source path: ", "Red",
            "$strSourcePath `r`n", "Yellow"
        )

        [ColorWriter]::Write($arrColoredString)
    }
}