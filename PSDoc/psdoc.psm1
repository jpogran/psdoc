function Get-Hashtable {
    # Path to PSD1 file to evaluate
    param (
        [parameter(
            Position = 0,
            ValueFromPipelineByPropertyName,
            Mandatory
        )]
        [ValidateNotNullOrEmpty()]
        [Alias('FullName')]
        [string]
        $Path
    )
    process
    {
        Write-Verbose "Loading data from $Path."
        Invoke-Expression "DATA { $(Get-Content -Raw -Path $Path) }"
    }
<#
Copied from PowerShellOrg DSC module https://github.com/PowerShellOrg/DSC/blob/master/Tooling/DscDevelopment/DscDevelopment.psm1
#>
}

Function Start-PSDoc
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Path,
        [Parameter()]
        $OutputPath = (Join-Path $Path "output")
    )

    # convention
    $configPath = Join-Path $Path "config"
    $configs = Get-ChildItem -Path $configPath -Filter *.psd1

    $configs | % {
        $files    = @();
        $hash     = Get-Hashtable $_.FullName
        $name     = $hash["Name"]
        $title    = $hash["Title"]
        $author   = $hash["Author"]
        $subject  = $hash["Subject"]
        $abstract = $hash["Abstract"]
        $company = $hash["Company"]
        $email = $hash["Email"]

        $header = (Get-ChildItem (Join-Path $Path $hash["Header"]))
        Write-Verbose $header
        $files  += $header

        $hash["Content"] | %{
            $value = $_
            Write-Verbose $value

            $resolved = Join-Path $Path $value
            $file     = Get-ChildItem $resolved

            Write-Verbose $file.FullName
            $files    += $file
        }

        if($hash.ContainsKey("Template")){
            $template = (Get-ChildItem (Join-Path $Path $hash["Template"])).FullName; Write-Verbose $template
        }

        $hash["formats"] | % {
            $format = $_
            $outfile = "$(Join-Path (Resolve-Path $OutputPath) $name).$($_)"
            Write-Verbose $outfile

            $rawFile = [IO.Path]::GetTempFileName()
            Write-Verbose $rawFile

            $files | %{
                $content += Get-Content -Path $_.FullName -Raw
                $content += "`n`n"
            }

            Out-File -FilePath $rawfile -InputObject $content -Encoding ascii -Force

            switch($format){
                "pdf"{
                    pandoc -s -S $rawfile --toc -N --variable mainfont=Georgia --variable sansfont=Arial --variable monofont="Bitstream Vera Sans Mono" --variable fontsize=12pt --variable version=1.9 -o $outfile
                }
                "html"{
                    pandoc $rawfile -s --highlight-style=zenburn --toc -N -t $format -o $outfile
                }
                "docx"{
                    Write-Warning "Do not Open any Word documents while this process runs"
                    pandoc -s -S $rawfile --toc --normalize -N -t $format -o $outfile --reference-docx $template
                    $word = New-WordDocument -Open $outfile `
                        -Visible $false `
                        -CompanyName $company `
                        -DocTitle $title `
                        -DocSubject $subject `
                        -DocAbstract $abstract `
                        -DocUserName $author `
                        -DocEmail $email
                    $word.NewTOC(1)
                    $word.NewCoverPage()
                    #$word.NewFooter("This document is intended for IPsoft personnel, as a master document of all available information for IPwin and IPwin Modules")
                    $word.Save()
                    $word.CloseDocument()
                }
                default{
                    pandoc -s -S $rawfile --toc --normalize -N -t $format -o $outfile
                }

            }

            $content = $null
        }
    }
<#
.EXAMPLE
pushd C:\Users\james\data\code\personal\psdoc
Import-Module .\PSDoc -Verbose
popd
cd C:\users\james\data\code\ipsoft\ipwin
Start-PSDoc -Path .\ipwindocs -Verbose

#>
}

function New-WordDocument
{
    [CmdletBinding()]
    param (
        [Parameter(HelpMessage='Make the document visible (or not).')]
        [bool]
        $Visible = $true,
        [Parameter(HelpMessage='Company name for cover page.')]
        [string]
        $CompanyName='Contoso Inc.',
        [Parameter(HelpMessage='Document title for cover page.')]
        [string]
        $DocTitle = 'Your Report',
        [Parameter(HelpMessage='Document subject for cover page.')]
        [string]
        $DocSubject = 'A great Word report.',
        [Parameter(HelpMessage='Document subject for cover page.')]
        [string]
        $DocAbstract = 'A great Word report.',
        [Parameter(HelpMessage='User name for cover page.')]
        [string]
        $DocUserName = $env:username,
        [Parameter(HelpMessage='Email for cover page.')]
        [string]
        $DocEmail = "$($env:username)@$($env:Domain)",
        $Open
    )
    try
    {
        $WordApp = New-Object -ComObject 'Word.Application'
        $WordVersion = [int]$WordApp.Version
        switch ($WordVersion) {
            15 {
                write-verbose 'Running Microsoft Word 2013'
                $WordProduct = 'Word 2013'
            }
            14 {
                write-verbose 'Running Microsoft Word 2010'
                $WordProduct = 'Word 2010'
            }
            12 {
                write-verbose 'Running Microsoft Word 2007'
                $WordProduct = 'Word 2007'
            }
            11 {
                write-verbose 'Running Microsoft Word 2003'
                $WordProduct = 'Word 2003'
            }
        }

        # Create a new blank document to work with and make the Word application visible.
        if($Open)
        {
            $WordDoc = $WordApp.Documents.Open($Open)
        }else{
            $WordDoc = $WordApp.Documents.Add()
        }
        $WordApp.Visible = $Visible

        # Store the old culture for later restoration.
        $OldCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture

        # May speed things up when creating larger docs
        # $SpellCheckSetting = $WordApp.Options.CheckSpellingAsYouType
        $GrammarCheckSetting = $WordApp.Options.CheckGrammarAsYouType
        $WordApp.Options.CheckSpellingAsYouType = $False
        $WordApp.Options.CheckGrammarAsYouType = $False

        # Set base culture
        ([System.Threading.Thread]::CurrentThread.CurrentCulture = 'en-US') | Out-Null

        $WordProps =
        @{
            'CompanyName'         = $CompanyName
            'Title'               = $DocTitle
            'Abstract'            = $DocAbstract
            'Subject'             = $DocSubject
            'Username'            = $DocUserName
            'Email'               = $DocEmail
            'Application'         = $WordApp
            'Document'            = $WordDoc
            'Selection'           = $WordApp.Selection
            'OldCulture'          = $OldCulture
            'SpellCheckSetting'   = $SpellCheckSetting
            'GrammarCheckSetting' = $GrammarCheckSetting
            'WordVersion'         = $WordVersion
            'WordProduct'         = $WordProduct
            'TableOfContents'     = $null
            'Saved'               = $false
        }
        $NewDoc = New-Object -TypeName PsObject -Property $WordProps
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewLine -Value {
            param (
                [Parameter( HelpMessage='Number of lines to instert.')]
                [int]
                $lines = 1
            )
            for ($index = 0; $index -lt $lines; $index++) {
                ($this.Selection).TypeParagraph()
            }
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name Save -Value {
            try
            {
                $this.Document.Save()
                $this.Saved = $true
            }
            catch
            {
                Write-Warning "Report was unable to be saved"
                $this.Saved = $false
            }
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name SaveAs -Value {
            param (
                [Parameter( HelpMessage='Report file name.')]
                [string]
                $WordDocFileName = 'report.docx'
            )
            try
            {
                $this.Document.SaveAs([ref]$WordDocFileName)
                $this.Saved = $true
            }
            catch
            {
                Write-Warning "Report was unable to be saved as $WordDocFileName"
                $this.Saved = $false
            }
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewText -Value {
            param (
                [Parameter( HelpMessage='Text to instert.')]
                [string]
                $text = ''
            )
            ($this.Selection).TypeText($text)
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewPageBreak -Value {
            ($this.Selection).InsertNewPage()
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name MoveToEnd -Value {
            ($this.Selection).Start = (($this.Selection).StoryLength - 1)
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewCoverPage -Value {
            param (
                [Parameter( HelpMessage='Coverpage Template.')]
                [string]
                $CoverPage = 'Facet'
            )
            # Go back to the beginning of the document
            $this.Selection.GoTo(1, 2, $null, 1) | Out-Null
            [bool]$CoverPagesExist = $False
            [bool]$BuildingBlocksExist = $False

            $this.Application.Templates.LoadBuildingBlocks()
            if ($this.WordVersion -eq 12) # Word 2007
            {
                $BuildingBlocks = $this.Application.Templates |
                    Where {$_.name -eq 'Building Blocks.dotx'}
            }
            else
            {
                $BuildingBlocks = $this.Application.Templates |
                    Where {$_.name -eq 'Built-In Building Blocks.dotx'}
            }

            Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
            $part = $Null

            if ($BuildingBlocks -ne $Null)
            {
                $BuildingBlocksExist = $True

                try
                {
                    Write-Verbose 'Setting Coverpage'
                    $part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
                }
                catch
                {
                    $part = $Null
                }

                if ($part -ne $Null)
                {
                    $CoverPagesExist = $True
                }
            }

            if ($CoverPagesExist)
            {
                Write-Verbose "$(Get-Date): Set Cover Page Properties"
                $this.SetDocProp($this.document.BuiltInDocumentProperties, 'Company', $this.CompanyName)
                $this.SetDocProp($this.document.BuiltInDocumentProperties, 'Title', $this.Title)
                $this.SetDocProp($this.document.BuiltInDocumentProperties, 'Subject', $this.Subject)
                $this.SetDocProp($this.document.BuiltInDocumentProperties, 'Author', $this.Username)
                #$this.SetDocProp($this.document.BuiltInDocumentProperties, 'Email', $this.Email)

                #Get the Coverpage XML part
                $cp = $this.Document.CustomXMLParts | where {$_.NamespaceURI -match "coverPageProps$"}

                #get the abstract XML part
                $ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}
                [string]$abstract = "$($this.Abstract) for $($this.CompanyName)"
                $ab.Text = $abstract

                $ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
                [string]$abstract = (Get-Date -Format d).ToString()
                $ab.Text = $abstract

                $part.Insert($this.Selection.Range,$True) | out-null
                $this.Selection.InsertNewPage()
            }
            else
            {
                $this.NewLine(5)
                $this.Selection.Style = "Title"
                $this.Selection.ParagraphFormat.Alignment = "wdAlignParagraphCenter"
                $this.Selection.TypeText($this.Title)
                $this.NewLine()
                $this.Selection.ParagraphFormat.Alignment = "wdAlignParagraphCenter"
                $this.Selection.Font.Size = 24
                $this.Selection.TypeText($this.Subject)
                $this.NewLine()
                $this.Selection.ParagraphFormat.Alignment = "wdAlignParagraphCenter"
                $this.Selection.Font.Size = 18
                $this.Selection.TypeText("Date: $(get-date)")
                $this.NewPageBreak()
            }
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewBlankPage -Value {
            param (
                [Parameter(HelpMessage='Cover page sub-title')]
                [int]
                $NumberOfPages
            )
            for ($i = 0; $i -lt $NumberOfPages; $i++){
                $this.Selection.Font.Size = 11
                $this.Selection.ParagraphFormat.Alignment = "wdAlignParagraphLeft"
                $this.NewPageBreak()
            }
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewTable -Value {
            param (
                [Parameter(HelpMessage='Rows')]
                [int]
                $NumRows=1,
                [Parameter(HelpMessage='Columns')]
                [int]
                $NumCols=1,
                [Parameter(HelpMessage='Include first row as header')]
                [bool]
                $HeaderRow = $true
            )
            $NewTable = $this.Document.Tables.Add($this.Selection.Range, $NumRows, $NumCols)
            $NewTable.AllowAutofit = $true
            $NewTable.AutoFitBehavior(2)
            $NewTable.AllowPageBreaks = $false
            $NewTable.Style = "Grid Table 4 - Accent 1"
            $NewTable.ApplyStyleHeadingRows = $HeaderRow
            return $NewTable
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewTableFromArray -Value {
            param (
                [Parameter(Mandatory=$true,
                           HelpMessage='Array of objects.')]
                $ObjArray,
                [Parameter(HelpMessage='Include first row as header')]
                [bool]
                $HeaderRow = $true
            )
            $AllObjects = @()
            $AllObjects += $ObjArray
            if ($HeaderRow)
            {
                $TableToInsert = ($AllObjects |
                                    ConvertTo-Csv -NoTypeInformation |
                                        Out-String) -replace '"',''
            }
            else
            {
                $TableToInsert = ($AllObjects |
                                    ConvertTo-Csv -NoTypeInformation |
                                        Select -Skip 1 |
                                            Out-String) -replace '"',''
            }
            $Range = $this.Selection.Range
            $Range.Text = "$TableToInsert"
            $Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByCommas
            $NewTable = $Range.ConvertToTable($Separator)
            $NewTable.AutoFormat([Microsoft.Office.Interop.Word.WdTableFormat]::wdTableFormatElegant)
            $NewTable.AllowAutofit = $true
            $NewTable.AutoFitBehavior(2)
            $NewTable.AllowPageBreaks = $false
            $NewTable.Style = "Grid Table 4 - Accent 1"
            $NewTable.ApplyStyleHeadingRows = $true
            return $NewTable
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewBookmark -Value {
            param (
                [Parameter(Mandatory=$true,
                           HelpMessage='A bookmark name')]
                [string]
                $BookmarkName
            )
            $this:Document.Bookmarks.Add($BookmarkName,$this.Selection)
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name SetDocProp -Value {
            #jeff hicks
            Param(
                [object]
                $Properties,
                [string]
                $Name,
                [string]
                $Value
            )
            #get the property object
            $prop = $properties | ForEach {
                $propname=$_.GetType().InvokeMember("Name","GetProperty",$Null,$_,$Null)
                If($propname -eq $Name)
                {
                    Return $_
                }
            }

            #set the value
            $Prop.GetType().InvokeMember("Value","SetProperty",$Null,$prop,$Value)
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewHeading -Value {
            param(
                [string]
                $Label = '',
                [string]
                $Style = 'Heading 1'
            )
            $this.Selection.Style = $Style
            $this.Selection.TypeText($Label)
            $this.Selection.TypeParagraph()
            $this.Selection.Style = "Normal"
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name NewTOC -Value {
            param (
                [Parameter(Mandatory=$true,
                           HelpMessage='A number to instert your table of contents into.')]
                [int]
                $PageNumber = 2,
                [string]
                $TOCHeading = 'Table of Contents',
                [string]
                $TOCHeaderStyle = 'Heading 1'
            )
            # Go back to the beginning of page two.
            $this.Selection.GoTo(1, 2, $null, $PageNumber) | Out-Null
            $this.NewHeading($TOCHeading,$TOCHeaderStyle)

            # Create Table of Contents for document.
            # Set Range to beginning of document to insert the Table of Contents.
            $TOCRange = $this.Selection.Range
            $useHeadingStyles = $true
            $upperHeadingLevel = 1 # <-- Heading1 or Title
            $lowerHeadingLevel = 2 # <-- Heading2 or Subtitle
            $useFields = $false
            $tableID = $null
            $rightAlignPageNumbers = $true
            $includePageNumbers = $true

            # to include any other style set in the document add them here
            $addedStyles = $null
            $useHyperlinks = $true
            $hidePageNumbersInWeb = $true
            $useOutlineLevels = $true

            # Insert Table of Contents
            $TableOfContents = $this.Document.TablesOfContents.Add($TocRange, $useHeadingStyles,
                               $upperHeadingLevel, $lowerHeadingLevel, $useFields, $tableID,
                               $rightAlignPageNumbers, $includePageNumbers, $addedStyles,
                               $useHyperlinks, $hidePageNumbersInWeb, $useOutlineLevels)
            $TableOfContents.TabLeader = 0
            $this.TableOfContents = $TableOfContents
            $this.MoveToEnd()
        }
        $NewDoc | Add-Member -MemberType ScriptMethod -Name CloseDocument -Value {
            try
            {
                # $WordObject.Application.Options.CheckSpellingAsYouType = $WordObject.SpellCheckSetting
                $this.Application.Options.CheckGrammarAsYouType = $this.GrammarCheckSetting
                $this.Document.Save()
                $this.Application.Quit()
                [System.Threading.Thread]::CurrentThread.CurrentCulture = $this.OldCulture


                # Truly release the com object, otherwise it will linger like a bad ghost
                [system.Runtime.InteropServices.marshal]::ReleaseComObject($this.Application) | Out-Null

                # Perform garbage collection
                [gc]::collect()
                [gc]::WaitForPendingFinalizers()
            }
            catch
            {
                Write-Warning 'There was an issue closing the word document.'
                Write-Warning ('Close-WordDocument: {0}' -f $_.Exception.Message)
            }
        }
        Return $NewDoc
    }
    catch
    {
        Write-Error 'There was an issue instantiating the new word document, is MS word installed?'
        Write-Error ('New-WordDocument: {0}' -f $_.Exception.Message)
        Throw "New-WordDocument: Problems creating new word document"
    }
}

Export-ModuleMember -Function Start-PSDoc