#Tells the user information
"This is going to PDF all your excel files in each of the folders and subfolders of your current directory.
Press any key to continue... Close if not"

#Require the user to press a key to continue
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
"Getting on your files..."

# Ask to redo files
"Do you want to re-export all the existing pdfs in your directory? Press y or Press n"
$key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

$redo = $key.character -eq "y"
if($redo){"OK I will redo those!"}
"Starting the process..."

#pwd means print working directory. This is a string!
$path = pwd

#See https://devblogs.microsoft.com/scripting/save-a-microsoft-excel-workbook-as-a-pdf-file-by-using-powershell/
#Cause idk really
$xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type]

# Finds the paths for all excel files within your current directory and sub directories with the extension .xlsm
$excelFiles = Get-ChildItem -Path $path -include *.xlsm -recurse

#This opens up excel
$objExcel = New-Object -ComObject excel.application

# Set the visible property to false to avoid a bunch of spreadsheets popping up and closing 
$objExcel.visible = $false

# initialize counters. counter counts how many pdfs that are saved. counterexists counts all the pdfs that already exist
$counter = 0

# for each excel file, check if the pdf of the file exists, if it doesnt, make a pdf of the daily summary page and save it in its folder. 
# If the pdf already exists skip.
foreach($wb in $excelFiles){

    # takes each excel file, and sets a variable with the full path, and one with just the extension (.xlsm). 
    # Then replace .xlsm with .pdf in the string path. And then get the file name only (-leaf).
    $fullpath = $wb.FullName
    $fileext = $wb.extension
    $pdfpath = $fullpath.Replace($fileext,".pdf")
    $basepdfpath = Split-Path -Path $pdfpath -leaf
    $pdfexists = Test-Path -Path $pdfpath
    $wedontwantredo = !($redo)

    if($pdfexists -and $wedontwantredo){
        "Skipping $basepdfpath"
        #continue means move to the next excel file a.k.a. "each" This will skipp the rest in this block, this goes back to the foreach.
        continue
    }

    $workbook = $objExcel.workbooks.open($wb.fullname, 3)
    $workbook.Saved = $true
    "saving $basepdfpath"
    $workbook.worksheets("Daily Summary").ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $pdfpath)
    $objExcel.Workbooks.close()
    # adds 1 to the counter if this runs
    $counter = $counter + 1
}
"Files Saved: $counter"

$objExcel.Quit()
