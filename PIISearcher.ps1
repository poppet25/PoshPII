using namespace System.Xml

try{
    Add-Type -Path ./lib/DocumentFormat.OpenXml.dll
    Add-Type -Path ./lib/Xceed.Words.NET.dll
}catch{
    Write-Host $_
}

function Find-Excel
{
    Param(
        [Parameter(ValueFromPipeline=$true)][String]$File
    )
    $spreadsheetDocument = [SpreadsheetDocument]::Open($File, $false)
    $workbookPart = $spreadsheetDocument.WorkbookPart

    $reader = [OpenXmlReader]::Create($workbookPart.SharedStringTablePart)

    while ($reader.Read())
    {
        if ($reader.ElementType -eq [SharedStringItem])
        {
            ($reader.LoadCurrentElement()).InnerText | Find-SSN
        }
    }
}

function Find-Word
{
    Param(
        [Parameter(ValueFromPipeline=$true)][String]$File
    )
    try{
        $document = [WordProcessingDocument]::Open($File, $false)
        $text = $document.MainDocumentPart.Document.Body.InnerText

    }catch{

    }


}

function unpack-docx
{
    Param(
        [String]$File
    )
    $doc = new-object -type system.xml.xmldocument
    $Package = [System.IO.Packaging.Package]::Open($File)
    $documentType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    $OfficeDocRel = $Package.GetRelationshipsByType($documentType)
    $documentPart = $Package.GetPart([System.IO.Packaging.PackUriHelper]::ResolvePartUri("/", $OfficeDocRel.TargetUri))
    $doc.load($documentPart.GetStream())

    $mgr = [XmlNamespaceManager]::new($doc.NameTable)
    $mgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

    $doc.SelectNodes("/descendant::w:t", $mgr) | Start-RSJob -ScriptBlock {
        Param(
            [Parameter(ValueFromPipeline=$true)][psobject]$node
        )

        return $node.InnerText
    }

    while(Get-RSJob)
    {
        $jobs = Get-RSJob | Where-Object { $_.State -in "Completed","Stoped","Finished"}
        $jobs | Receive-RSJob
        $jobs | Remove-RSJob
    }

    $Package.Close()
}

#Find-Word -File "./docs/HubertDean.docx"
unpack-docx -File "./docs/HubertDean.docx"