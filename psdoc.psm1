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
}

Function Compile-Files
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
        $files  = @();        
        $hash    = Get-Hashtable $_.FullName
        $name    = $hash["Name"]
        $content = $hash["Content"].GetEnumerator() | Sort Name

        $content.GetEnumerator() | %{
            $value = $_.Value
            Write-Verbose $value

            $resolved = Join-Path $Path $value
            $file     = Get-ChildItem $resolved
            
            Write-Verbose $file.FullName
            $files    += $file
        }

        $hash["formats"] | % {
            $outfile = "$(Join-Path (Resolve-Path $OutputPath) $name)$($_)"
            
            Write-Verbose $outfile
            $files | cat | pandoc -s -S --toc -o $outfile
        }
    }
<#
.EXAMPLE
    Compile-Files -Path ./ipwin-docs -OutputPath ./ipwin-docs/output -Verbose
#>
}