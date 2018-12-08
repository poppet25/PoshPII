
$data = 0..1000 | %{ [Array]@($_, "Name$_","FileName$_") }
$chunkSize = 1000
$chunks = $data.Count/$chunkSize
$skip = 0
$sb = @'
    <style type="text/css">
        body{
            margin-top:20px;
        }
    </style>
    <script type="text/javascript">
    $(document).ready(function(){
        $('body').addClass('container-fluid')
        $('table').dataTable({
            "columns": [
                { "data": "Name" },
                { "data": "Value" },
                { "data": "FileType" }
            ]
        }); 
    });
    </script>
'@
$css = @(
    '<link rel="stylesheet" type="text/css" href="js/DataTables/datatables.min.css">',
    '<link rel="stylesheet" type="text/css" href="js/bootstrap-4.1.3-dist/css/bootstrap.min.css">'
    )
$js = @(
    '<script src="js/jquery-3.3.1.min.js" ></script>',
    '<script src="js/DataTables/datatables.min.js" ></script>',
    '<script src="js/bootstrap-4.1.3-dist/js/bootstrap.min.js" ></script>',
    $sb
    )
$body = @'
    <div class="card card-body bg-secondary text-center">
        <h1>PII Searcher</h1>
    </div>
'@
$meta = @{
    'Content-Type' = "text/html"
}

for($idx = 1; $idx -lt $chunks; $idx++)
{
    Start-RSJob -Name "job$idx" -ScriptBlock {
        $Using:data | Select-Object -skip $Using:skip -First $Using:chunkSize | %{ New-Object -TypeName psobject -Property @{ID=$_; Name="ID$_"; Value=$_}}
    }
    $skip = $idx*$chunkSize
}

while(Get-RSJob)
{
    Write-Host "." -NoNewline
    Start-Sleep -Milliseconds 100

    Get-RSJob | `
        Where-Object { $_.State -in "Completed"} | `
        Receive-RSJob  | `
        ConvertTo-Html -Meta $meta -Head $css -PostContent $js -PreContent $body -Charset utf8 |  `
        Out-File -FilePath "file.htm"

    Get-RSJob | Where-Object { $_.State -in "Completed"} | Remove-RSJob
}