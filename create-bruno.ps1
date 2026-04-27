<#	
  .Synopsis
    PowerShell script to create a Bruno collection folder from a Workspace ONE UEM server.
  .NOTES
	  Created:   	    January, 2026
	  Created by:	    Phil Helmling
	  Organization:   Omnissa LLC
    Filename:       create-bruno.ps1
    GitHub:         https://github.com/helmlingp/create_wsone_api_bruno_collection
    
  .DESCRIPTION
    PowerShell script to create a Bruno collection folder from a Workspace ONE UEM server.
    Steps:
    1. Prompts for server and OAuth details
    2. Fetches OAuth2 token (client_credentials)
    3. Gets server version from /api/system/info
    4. Creates folder named with the version
    5. Writes bruno.json and collection.bru with placeholders replaced
    6. Downloads the API docs endpoints into the folder
    7. Replaces `post { auth: ... }` occurrences with `post { auth: inherit }` in all .bru files

  .REQUIREMENTS
    PowerShell 5.1 or later
    Workspace ONE UEM server with OAuth2 client credentials
    A Global Environment configured within Bruno with the following variables:
      Oauth_token_URL: https://UEM_OAUTH_TOKEN_URL e.g. https://uat.uemauth.workspaceone.com/connect/token
      Oauth_clientID: YOUR_CLIENT_ID
      Oauth_client_secret: YOUR_CLIENT_SECRET
      YOUR_API_SERVER: e.g. apiXXX.awmdm.com or https://apiXXX.awmdm.com
      aw-tenant-code: optional

    see https://docs.omnissa.com/bundle/WorkspaceONE-UEM-Console-BasicsVSaaS/page/UsingUEMFunctionalityWithRESTAPI.html for more information on setting up OAuth clients in Workspace ONE UEM.
  .EXAMPLE
    Provide connection parameters on command line
    powershell.exe -ep bypass -file .\create-bruno.ps1 -clientId YOUR_CLIENT_ID -clientSecret YOUR_CLIENT_SECRET -tokenUrl Oauth_token_URL -Server YOUR_API_SERVER

    Prompt for connection parameters 
    powershell.exe -ep bypass -file .\create-bruno.ps1

#>
param (
    [Parameter(Mandatory=$false)]
    [string]$clientId=$script:clientId,
    [Parameter(Mandatory=$false)]
    [string]$clientSecret=$script:clientSecret,
    [Parameter(Mandatory=$false)]
    [string]$tokenUrl=$script:tokenUrl,
    [Parameter(Mandatory=$false)]
    [string]$Server=$script:Server
)
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#-----------------------------------------------------------[Conversion helpers]-----------------------------------------------------------

function Sanitize-Name([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return 'unnamed' }
    $name = $s -replace '[\\/:*?"<>|]', '-'
    $name = $name -replace '\s+', ' '
    return $name.Trim()
}

function Escape-Braces([string]$s) {
    if ($null -eq $s) { return '' }
    return $s.Replace('{', '{{').Replace('}', '}}')
}

function Get-ShortSummary([string]$summary) {
    if ([string]::IsNullOrWhiteSpace($summary)) { return '' }
    $clean = ($summary -replace '\r?\n', ' ').Trim()
    $idx = $clean.IndexOf('.')
    if ($idx -gt -1) { return $clean.Substring(0, $idx).Trim() }
    return $clean
}

function Get-SchemaRefs($schema) {
    $refs = [System.Collections.Generic.HashSet[string]]::new()
    if ($null -eq $schema) { return @() }
    function Walk($node, $set) {
        if ($null -eq $node) { return }
        if ($node -is [System.Management.Automation.PSCustomObject]) {
            foreach ($prop in $node.PSObject.Properties) {
                if ($prop.Name -eq '$ref' -and $prop.Value) { [void]$set.Add([string]$prop.Value) }
                Walk $prop.Value $set
            }
            return
        }
        if ($node -is [System.Collections.IDictionary]) {
            foreach ($k in $node.Keys) {
                $v = $node[$k]
                if ($k -eq '$ref' -and $v) { [void]$set.Add([string]$v) }
                Walk $v $set
            }
            return
        }
        if ($node -is [System.Collections.IEnumerable] -and -not ($node -is [string])) {
            foreach ($item in $node) { Walk $item $set }
        }
    }
    Walk $schema $refs
    return @($refs)
}

function Get-Props($obj) {
    if ($null -eq $obj) { return @() }
    if ($obj -is [System.Collections.IDictionary]) {
        return @($obj.GetEnumerator() | ForEach-Object {
            [PSCustomObject]@{ Name = [string]$_.Key; Value = $_.Value }
        })
    }
    return @($obj.PSObject.Properties | ForEach-Object {
        [PSCustomObject]@{ Name = [string]$_.Name; Value = $_.Value }
    })
}

function Resolve-Ref([string]$ref, $api) {
    if (-not $ref.StartsWith('#/')) { return $null }
    $parts = $ref.TrimStart('#/') -split '/'
    $node = $api
    foreach ($part in $parts) {
        if ($node -is [System.Collections.IDictionary] -and $node.ContainsKey($part)) {
            $node = $node[$part]
        } else { return $null }
    }
    return $node
}

function Render-Schema($schema, $api, $indent, [System.Collections.Generic.HashSet[string]]$seen) {
    $lines = @()
    if ($null -eq $schema) { return $lines }
    if ($indent -gt 3) { return @("$indent(max depth reached)") }
    $pad = '  ' * $indent

    if ($schema -is [System.Collections.IDictionary] -and $schema.ContainsKey('$ref')) {
        $refStr = [string]$schema['$ref']
        if ($seen.Contains($refStr)) { $lines += "${pad}(circular ref: $refStr)"; return $lines }
        [void]$seen.Add($refStr)
        $resolved = Resolve-Ref $refStr $api
        if ($null -eq $resolved) { $lines += "${pad}(unresolved: $refStr)"; [void]$seen.Remove($refStr); return $lines }
        $schemaName = $refStr -replace '^#/components/schemas/', ''
        $desc = if ($resolved.ContainsKey('description')) { ' - ' + [string]$resolved['description'] } else { '' }
        $lines += "${pad}[$schemaName]$desc"
        $lines += Render-Schema $resolved $api $indent $seen
        [void]$seen.Remove($refStr)
        return $lines
    }

    if ($schema -is [System.Collections.IDictionary] -and $schema.ContainsKey('type') -and $schema['type'] -eq 'array') {
        if ($schema.ContainsKey('items')) {
            $lines += "${pad}(array of:)"
            $lines += Render-Schema $schema['items'] $api ($indent + 1) $seen
        }
        return $lines
    }

    if ($schema -is [System.Collections.IDictionary] -and $schema.ContainsKey('properties')) {
        $required = @()
        if ($schema.ContainsKey('required')) { $required = @($schema['required']) }
        foreach ($prop in (Get-Props $schema['properties'])) {
            $pName = $prop.Name
            $pSchema = $prop.Value
            $pType = ''; $pDesc = ''; $pEnum = ''; $pFmt = ''
            $pReq = if ($required -contains $pName) { ' (required)' } else { '' }
            if ($pSchema -is [System.Collections.IDictionary]) {
                if ($pSchema.ContainsKey('type'))        { $pType = [string]$pSchema['type'] }
                if ($pSchema.ContainsKey('format'))      { $pFmt  = '/' + [string]$pSchema['format'] }
                if ($pSchema.ContainsKey('description')) { $pDesc = ' - ' + [string]$pSchema['description'] }
                if ($pSchema.ContainsKey('enum')) {
                    $vals = @($pSchema['enum']) -join ', '
                    $pEnum = " [enum: $vals]"
                }
            }
            if ($pSchema -is [System.Collections.IDictionary] -and $pSchema.ContainsKey('$ref')) {
                $lines += "${pad}${pName}${pReq}:"
                $lines += Render-Schema $pSchema $api ($indent + 1) $seen
            } elseif ($pSchema -is [System.Collections.IDictionary] -and $pSchema.ContainsKey('type') -and $pSchema['type'] -eq 'object' -and $pSchema.ContainsKey('properties')) {
                $lines += "${pad}${pName}${pReq}: object"
                $lines += Render-Schema $pSchema $api ($indent + 1) $seen
            } elseif ($pSchema -is [System.Collections.IDictionary] -and $pSchema.ContainsKey('type') -and $pSchema['type'] -eq 'array') {
                $lines += "${pad}${pName}${pReq}: array${pDesc}"
                $lines += Render-Schema $pSchema $api ($indent + 1) $seen
            } else {
                $lines += "${pad}${pName}${pReq}: ${pType}${pFmt}${pEnum}${pDesc}"
            }
        }
    }
    return $lines
}

function Get-SchemaDocLines([string]$refStr, $api) {
    $lines = [System.Collections.Generic.List[string]]::new()
    $schemaName = $refStr -replace '^#/components/schemas/', ''
    $resolved = Resolve-Ref $refStr $api
    if ($null -eq $resolved) { $lines.Add("  Schema: $schemaName (unresolved)"); return $lines }
    $desc = if ($resolved -is [System.Collections.IDictionary] -and $resolved.ContainsKey('description')) {
        ' - ' + [string]$resolved['description'] } else { '' }
    $lines.Add("  Schema: $schemaName$desc")
    $seen = [System.Collections.Generic.HashSet[string]]::new()
    [void]$seen.Add($refStr)
    foreach ($l in (Render-Schema $resolved $api 1 $seen)) { $lines.Add("  $l") }
    return $lines
}

function Convert-ToMarkdownListLines($lines) {
    $md = [System.Collections.Generic.List[string]]::new()
    foreach ($line in @($lines)) {
        if ($null -eq $line) { continue }
        $raw = [string]$line
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }
        $trim = $raw.Trim()
        $leading = $raw.Length - $raw.TrimStart().Length
        $level = [int][Math]::Floor($leading / 2)
        if ($level -lt 0) { $level = 0 }
        $md.Add(('  ' * $level) + '- ' + $trim)
    }
    return $md
}

function Get-SchemaLabel($schema) {
    if ($null -eq $schema) { return 'unknown' }
    if ($schema -is [System.Collections.IDictionary]) {
        if ($schema.ContainsKey('$ref')) {
            return ([string]$schema['$ref'] -replace '^#/components/schemas/', '')
        }
        if ($schema.ContainsKey('type')) {
            $t = [string]$schema['type']
            if ($t -eq 'array' -and $schema.ContainsKey('items')) {
                return "array[$(Get-SchemaLabel $schema['items'])]"
            }
            if ($schema.ContainsKey('format')) {
                return "$t/$([string]$schema['format'])"
            }
            return $t
        }
    }
    return 'object'
}

function Resolve-Parameter($parameter, $api) {
    if ($parameter -is [System.Collections.IDictionary] -and $parameter.ContainsKey('$ref')) {
        $resolved = Resolve-Ref ([string]$parameter['$ref']) $api
        if ($null -ne $resolved) { return $resolved }
    }
    return $parameter
}

function Get-ParameterDocLines($parameters, $api) {
    $lines = @()
    if ($null -eq $parameters) { return $lines }

    foreach ($pRaw in @($parameters)) {
        $p = Resolve-Parameter $pRaw $api
        if ($null -eq $p) { continue }

        $name = if ($p.name) { [string]$p.name } else { 'unnamed' }
        $loc = if ($p.in) { [string]$p.in } else { 'unknown' }
        $required = if ($p.required -eq $true) { 'required' } else { 'optional' }
        $desc = if ($p.description) { ' - ' + (Escape-Braces ([string]$p.description)) } else { '' }

        $typeLabel = ''
        if ($p.schema) {
            $typeLabel = Get-SchemaLabel $p.schema
        } elseif ($p.content) {
            $media = @(Get-Props $p.content | ForEach-Object { [string]$_.Name })
            if ($media.Count -gt 0) {
                $typeLabel = 'content: ' + ($media -join ', ')
            }
        }

        if ($typeLabel) {
            $lines += "  $name ($loc, $required): $typeLabel$desc"
        } else {
            $lines += "  $name ($loc, $required)$desc"
        }
    }

    return $lines
}

function Get-ResponseDocLines($responses, $api) {
    $lines = @()
    if ($null -eq $responses) { return $lines }

    $respProps = Get-Props $responses
    $sorted = $respProps | Sort-Object {
        if ($_.Name -match '^\d+$') { [int]$_.Name } else { 9999 }
    }, Name

    foreach ($resp in $sorted) {
        $status = [string]$resp.Name
        $r = $resp.Value
        $desc = if ($r.description) { Escape-Braces ([string]$r.description) } else { '' }
        if ($desc) {
            $lines += "  ${status}: $desc"
        } else {
            $lines += "  $status"
        }

        if ($r.content) {
            foreach ($ct in (Get-Props $r.content)) {
                $schemaLabel = ''
                if ($ct.Value.schema) { $schemaLabel = Get-SchemaLabel $ct.Value.schema }
                if ($schemaLabel) {
                    $lines += "    - $($ct.Name): $schemaLabel"
                } else {
                    $lines += "    - $($ct.Name)"
                }
            }
        }
    }

    return $lines
}

function Get-ParamValue($parameter) {
    if ($null -eq $parameter) { return '' }
    if ($parameter.example) { return [string]$parameter.example }
    if ($parameter.schema -and $parameter.schema.example) { return [string]$parameter.schema.example }
    return ''
}

function Get-ReasonPhrase([string]$statusCode) {
    if ($statusCode -notmatch '^\d+$') { return $statusCode }
    try {
        return [string]([System.Net.HttpStatusCode]([int]$statusCode))
    } catch {
        return $statusCode
    }
}

function Quote-BruString([string]$value) {
    if ($null -eq $value) { $value = '' }
    $safe = $value -replace '\\', '\\\\'
    $safe = $safe -replace '"', '\\"'
    $safe = $safe -replace '\r?\n', ' '
    return '"' + $safe + '"'
}

function Get-PrimitiveExample([string]$type, [string]$format) {
    if ($type -eq 'integer') { return 0 }
    if ($type -eq 'number') { return 0 }
    if ($type -eq 'boolean') { return $false }
    if ($type -eq 'string') {
        if ($format -eq 'uuid') { return '00000000-0000-0000-0000-000000000000' }
        if ($format -eq 'date-time') { return '2026-01-01T00:00:00Z' }
        if ($format -eq 'date') { return '2026-01-01' }
        return 'Text value'
    }
    return $null
}

function Build-SchemaExample($schema, $api, [System.Collections.Generic.HashSet[string]]$seen, [int]$depth) {
    if ($null -eq $schema) { return $null }
    if ($depth -gt 6) { return $null }

    if ($schema -is [System.Collections.IDictionary]) {
        if ($schema.ContainsKey('example')) {
            return $schema['example']
        }

        if ($schema.ContainsKey('$ref')) {
            $refStr = [string]$schema['$ref']
            if ($seen.Contains($refStr)) { return $null }
            [void]$seen.Add($refStr)
            $resolved = Resolve-Ref $refStr $api
            $result = Build-SchemaExample $resolved $api $seen ($depth + 1)
            [void]$seen.Remove($refStr)
            return $result
        }

        if ($schema.ContainsKey('enum')) {
            $vals = @($schema['enum'])
            if ($vals.Count -gt 0) { return $vals[0] }
        }

        if ($schema.ContainsKey('type')) {
            $t = [string]$schema['type']
            if ($t -eq 'object') {
                $obj = [ordered]@{}
                if ($schema.ContainsKey('properties')) {
                    foreach ($prop in (Get-Props $schema['properties'])) {
                        $obj[[string]$prop.Name] = Build-SchemaExample $prop.Value $api $seen ($depth + 1)
                    }
                } elseif ($schema.ContainsKey('additionalProperties')) {
                    $obj['key'] = Build-SchemaExample $schema['additionalProperties'] $api $seen ($depth + 1)
                }
                return $obj
            }
            if ($t -eq 'array') {
                if ($schema.ContainsKey('items')) {
                    return @(Build-SchemaExample $schema['items'] $api $seen ($depth + 1))
                }
                return @()
            }
            return Get-PrimitiveExample $t ([string]$schema['format'])
        }

        if ($schema.ContainsKey('properties')) {
            $obj = [ordered]@{}
            foreach ($prop in (Get-Props $schema['properties'])) {
                $obj[[string]$prop.Name] = Build-SchemaExample $prop.Value $api $seen ($depth + 1)
            }
            return $obj
        }
    }

    return $null
}

function Get-SchemaExampleObject($schema, $api, [System.Collections.Generic.HashSet[string]]$seen) {
    if ($null -eq $schema) { return $null }
    if ($schema -is [System.Collections.IDictionary]) {
        if ($schema.ContainsKey('example')) { return $schema['example'] }
        if ($schema.ContainsKey('$ref')) {
            $refStr = [string]$schema['$ref']
            if ($seen.Contains($refStr)) { return $null }
            [void]$seen.Add($refStr)
            $resolved = Resolve-Ref $refStr $api
            $result = Get-SchemaExampleObject $resolved $api $seen
            [void]$seen.Remove($refStr)
            return $result
        }
    }
    return $null
}

function Get-ContentExampleText($contentObj, $api) {
    if ($null -eq $contentObj) { return '' }

    # Prefer schema model examples (including resolved $ref schemas).
    if ($contentObj.schema) {
        $seen = [System.Collections.Generic.HashSet[string]]::new()
        $schemaExample = Get-SchemaExampleObject $contentObj.schema $api $seen
        if ($null -ne $schemaExample) {
            if ($schemaExample -is [string]) { return [string]$schemaExample }
            return ($schemaExample | ConvertTo-Json -Depth 100)
        }
    }

    if ($contentObj.example) {
        if ($contentObj.example -is [string]) { return [string]$contentObj.example }
        return ($contentObj.example | ConvertTo-Json -Depth 100)
    }

    if ($contentObj.examples) {
        $firstEx = Get-Props $contentObj.examples | Select-Object -First 1
        if ($firstEx -and $firstEx.Value.value) {
            if ($firstEx.Value.value -is [string]) { return [string]$firstEx.Value.value }
            return ($firstEx.Value.value | ConvertTo-Json -Depth 100)
        }
    }

    if ($contentObj.schema) {
        $seen = [System.Collections.Generic.HashSet[string]]::new()
        $generated = Build-SchemaExample $contentObj.schema $api $seen 0
        if ($null -ne $generated) {
            if ($generated -is [string]) { return [string]$generated }
            return ($generated | ConvertTo-Json -Depth 100)
        }
    }

    return ''
}

function Build-RequestLines([string]$url, [string]$method, [string]$bodyMode, [bool]$hasBody, [string[]]$bodyJsonLines, $pathParamLines, $queryParamLines, $headersBlock) {
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('  request: {')
    $lines.Add("    url: $url")
    $lines.Add("    method: $method")
    if ($hasBody) { $lines.Add("    mode: $bodyMode") } else { $lines.Add('    mode: none') }
    if ($pathParamLines.Count -gt 0) {
        $lines.Add('    params:path: {')
        foreach ($line in $pathParamLines) { $lines.Add("    $line") }
        $lines.Add('    }')
    }
    if ($queryParamLines.Count -gt 0) {
        $lines.Add('    params:query: {')
        foreach ($line in $queryParamLines) { $lines.Add("    $line") }
        $lines.Add('    }')
    }
    if ($headersBlock.Count -gt 0) {
        $lines.Add('    ')
        $lines.Add('    headers: {')
        foreach ($h in $headersBlock) { $lines.Add("    $h") }
        $lines.Add('    }')
    }
    if ($hasBody) {
        $lines.Add('    ')
        if ($bodyMode -eq 'json') {
            $lines.Add('    body:json: {')
            foreach ($line in $bodyJsonLines) { $lines.Add("      $line") }
            $lines.Add('    }')
        } else {
            $lines.Add('    body:text: {')
            $lines.Add("      '''")
            foreach ($line in $bodyJsonLines) { $lines.Add("      $line") }
            $lines.Add("      '''")
            $lines.Add('    }')
        }
    }
    $lines.Add('  }')
    return $lines
}

function Get-ResponseExampleBlocks($responses, $api, [string]$url, [string]$method, [string]$bodyMode, [bool]$hasBody, [string]$bodyJson, $pathParamLines, $queryParamLines, $headersBlock) {
    $blocks = @()
    if (-not $responses) { return $blocks }

    $requestAcceptIsJson = $false
    foreach ($h in @($headersBlock)) {
        if ([string]$h -match '^\s*Accept:\s*application/json') { $requestAcceptIsJson = $true; break }
    }

    $bodyJsonLines = [string[]]@($bodyJson -split "`r?`n")

    foreach ($resp in (Get-Props $responses)) {
        $status    = [string]$resp.Name
        $respValue = $resp.Value
        $name      = if ($respValue.description) { ([string]$respValue.description).Trim() } else { "Status $status" }
        if (-not $name) { $name = "Status $status" }
        $reason    = Get-ReasonPhrase $status

        $contentProps = if ($respValue.content) { @(Get-Props $respValue.content) } else { @() }

        if ($contentProps.Count -eq 0) {
            $example = [System.Collections.Generic.List[string]]::new()
            $example.Add('example {')
            $example.Add("  name: " + (Quote-BruString $name))
            $example.Add('  ')
            $example.AddRange([string[]](Build-RequestLines $url $method $bodyMode $hasBody $bodyJsonLines $pathParamLines $queryParamLines $headersBlock))
            $example.Add('  ')
            $example.Add('  response: {')
            $example.Add('    status: {')
            $example.Add("      code: $status")
            $example.Add("      text: $reason")
            $example.Add('    }')
            $example.Add('  }')
            $example.Add('}')
            $blocks += ,($example -join "`n")
            continue
        }

        foreach ($ct in $contentProps) {
            $mediaType = [string]$ct.Name
            if ($requestAcceptIsJson -and $mediaType -match '^application/xml') { continue }
            $exampleText = Get-ContentExampleText $ct.Value $api
            $respMode    = if ($mediaType -like '*json*') { 'json' } else { 'text' }
            if (-not $exampleText) { $exampleText = if ($respMode -eq 'json') { '{}' } else { '' } }
            $exampleTextLines = [string[]]@($exampleText -split "`r?`n")

            $example = [System.Collections.Generic.List[string]]::new()
            $example.Add('example {')
            $example.Add("  name: " + (Quote-BruString $name))
            $example.Add('  ')
            $example.AddRange([string[]](Build-RequestLines $url $method $bodyMode $hasBody $bodyJsonLines $pathParamLines $queryParamLines $headersBlock))
            $example.Add('  ')
            $example.Add('  response: {')
            $example.Add('    headers: {')
            $example.Add("      Content-Type: $mediaType")
            $example.Add('    }')
            $example.Add('  ')
            $example.Add('    status: {')
            $example.Add("      code: $status")
            $example.Add("      text: $reason")
            $example.Add('    }')
            $example.Add('  ')
            $example.Add('    body: {')
            $example.Add("      type: $respMode")
            $example.Add("      content: '''")
            foreach ($line in $exampleTextLines) { $example.Add("      $line") }
            $example.Add("      '''")
            $example.Add('    }')
            $example.Add('  }')
            $example.Add('}')
            $blocks += ,($example -join "`n")
        }
    }

    return $blocks
}

# Convert a single downloaded OpenAPI JSON file into Bruno folder/request files under $outRoot.
function Convert-OpenApiToBruno {
    param(
        [string]$srcFile,   # path to the downloaded OpenAPI JSON file
        [string]$outRoot,   # output directory (the API subfolder)
        [string]$apiName,   # human name, e.g. "MAM API V1"
        [int]   $apiSeq,    # sequence number within the parent collection
        [string]$apiDocs,   # short description for the folder.bru docs block
        [string]$urlPrefix  # URL segment inserted between {{baseUrl}} and the path, e.g. "/mam"
    )

    Write-Host "  Parsing $srcFile ..." -ForegroundColor Gray
    $json = Get-Content -Raw -Path $srcFile
    $api  = $json | ConvertFrom-Json -AsHashtable -Depth 100

    # folder.bru — child entry in the versioned root collection
    $folderBruPath = Join-Path $outRoot 'folder.bru'
    if (-not (Test-Path $folderBruPath)) {
        $folderBruContent = "meta {`n  name: $apiName`n  seq: $apiSeq`n}`n`nauth {`n  mode: inherit`n}`n`ndocs {`n  $apiDocs`n`n  Contact Support:`n   Name: Workspace ONE UEM`n}"
        [System.IO.File]::WriteAllText($folderBruPath, $folderBruContent, [System.Text.UTF8Encoding]::new($false))
    }

    # Optional standalone collection files so the API folder can be opened on its own
    $brunoJsonPath = Join-Path $outRoot 'bruno.json'
    if (-not (Test-Path $brunoJsonPath)) {
        "{`n  `"version`": `"1`",`n  `"name`": `"$apiName`",`n  `"type`": `"collection`",`n  `"ignore`": [`"node_modules`", `".git`"]`n}" | Set-Content -Path $brunoJsonPath -Encoding UTF8
    }

    $collectionBruPath = Join-Path $outRoot 'collection.bru'
    if (-not (Test-Path $collectionBruPath)) {
        "meta {`n  name: $apiName`n}`n`nauth {`n  mode: inherit`n}" | Set-Content -Path $collectionBruPath -Encoding UTF8
    }

    $pathProps  = Get-Props $api.paths
    $seqByTag   = @{}
    $total      = 0

    foreach ($pathProp in $pathProps) {
        $pathKey  = [string]$pathProp.Name
        $pathItem = $pathProp.Value
        $ops = Get-Props $pathItem | Where-Object { $_.Name -in @('get','post','put','patch','delete','head','options') }

        foreach ($opProp in $ops) {
            $method = $opProp.Name.ToUpperInvariant()
            $op     = $opProp.Value

            $tag        = if ($op.tags -and $op.tags.Count -gt 0) { [string]$op.tags[0] } else { 'General' }
            $tagDirName = Sanitize-Name $tag
            $tagDir     = Join-Path $outRoot $tagDirName
            if (-not (Test-Path $tagDir)) { New-Item -ItemType Directory -Path $tagDir | Out-Null }

            $folderBru = Join-Path $tagDir 'folder.bru'
            if (-not (Test-Path $folderBru)) {
                "meta {`n  name: $tag`n}`n`nauth {`n  mode: inherit`n}" | Set-Content -Path $folderBru -Encoding UTF8
            }

            if (-not $seqByTag.ContainsKey($tagDirName)) { $seqByTag[$tagDirName] = 1 }
            $seq = [int]$seqByTag[$tagDirName]
            $seqByTag[$tagDirName] = $seq + 1

            $opName     = if ($op.summary) {
                $shortSummary = Get-ShortSummary ([string]$op.summary)
                if ($shortSummary) { $shortSummary } else { [string]$op.summary }
            } elseif ($op.operationId) { [string]$op.operationId } else { "$method $pathKey" }
            $safeOpName = Sanitize-Name $opName
            $bruPath    = Join-Path $tagDir ($safeOpName + '.bru')

            $urlPath = [regex]::Replace($pathKey, '\{([^}]+)\}', ':$1')
            $url     = "{{baseUrl}}$urlPrefix$urlPath"

            $contentType = ''
            if ($op.requestBody -and $op.requestBody.content) {
                $ctProps = Get-Props $op.requestBody.content
                if ($ctProps.Count -gt 0) { $contentType = [string]$ctProps[0].Name }
            }

            $accept = ''
            if ($op.responses) {
                $success = Get-Props $op.responses | Where-Object { $_.Name -match '^2\d\d$' } | Select-Object -First 1
                if ($success -and $success.Value.content) {
                    $aProps = Get-Props $success.Value.content
                    if ($aProps.Count -gt 0) { $accept = [string]$aProps[0].Name }
                }
            }
            if (-not $accept) { $accept = 'application/json' }

            $reqRefs = @()
            if ($op.requestBody -and $op.requestBody.content) {
                foreach ($ct in (Get-Props $op.requestBody.content)) {
                    if ($ct.Value.schema) { $reqRefs += Get-SchemaRefs $ct.Value.schema }
                }
                $reqRefs = $reqRefs | Where-Object { $_ } | Sort-Object -Unique
            }

            $bodyJson = ''
            if ($op.requestBody -and $op.requestBody.content) {
                $ctProps = Get-Props $op.requestBody.content
                if ($ctProps.Count -gt 0) {
                    $firstContent = $ctProps[0].Value
                    $bodyJson = Get-ContentExampleText $firstContent $api
                }
            }
            if (-not $bodyJson) { $bodyJson = '{}' }
            $bodyJsonLines = [string[]]@($bodyJson -split "`r?`n")

            $allParams = @()
            if ($pathItem.parameters) { $allParams += @($pathItem.parameters) }
            if ($op.parameters) { $allParams += @($op.parameters) }

            # De-duplicate by "in:name" with operation-level params taking precedence.
            $paramMap = @{}
            foreach ($pRaw in $allParams) {
                $p = Resolve-Parameter $pRaw $api
                if ($null -eq $p) { continue }
                $keyName = if ($p.name) { [string]$p.name } else { 'unnamed' }
                $keyLoc = if ($p.in) { [string]$p.in } else { 'unknown' }
                $paramMap["$keyLoc`:$keyName"] = $pRaw
            }
            $mergedParams = @($paramMap.Values)

            $pathParamLines = @()
            $queryParamLines = @()
            foreach ($pRaw in $mergedParams) {
                $p = Resolve-Parameter $pRaw $api
                if ($null -eq $p) { continue }
                $pName = if ($p.name) { [string]$p.name } else { '' }
                if (-not $pName) { continue }
                $pValue = Get-ParamValue $p
                if ($p.in -eq 'path') {
                    $pathParamLines += "  ${pName}: $pValue"
                } elseif ($p.in -eq 'query') {
                    $queryParamLines += "  ${pName}: $pValue"
                }
            }

            $docLines = @()
            if ($op.description) {
                $descText = ([string]$op.description) -replace '(?m)^(\s*)[\*\-]\s+', '$1'
                $docLines += (Escape-Braces $descText)
            }
            if ($reqRefs.Count -gt 0) {
                $docLines += ''
                $docLines += '### Request Body'
                foreach ($r in $reqRefs) {
                    $docLines += Convert-ToMarkdownListLines (Get-SchemaDocLines $r $api)
                }
            }

            $methodBlock  = $method.ToLowerInvariant()
            $headersBlock = @()
            if ($accept)      { $headersBlock += "  Accept: $accept" }
            $headersText = ($headersBlock -join "`n")

            $hasBody  = $op.requestBody -ne $null
            $bodyMode = if ($hasBody) {
                if ($contentType -like 'application/json*') { 'json' } else { 'text' }
            } else {
                'none'
            }
            $exampleBlocks = Get-ResponseExampleBlocks -responses $op.responses -api $api -url $url -method $method -bodyMode $bodyMode -hasBody $hasBody -bodyJson $bodyJson -pathParamLines $pathParamLines -queryParamLines $queryParamLines -headersBlock $headersBlock

            $bru = [System.Collections.Generic.List[string]]::new()
            $bru.Add('meta {')
            $bru.Add("  name: $opName")
            $bru.Add('  type: http')
            $bru.Add("  seq: $seq")
            $bru.Add('}')
            $bru.Add('')
            $bru.Add("$methodBlock {")
            $bru.Add("  url: $url")
            $bru.Add("  body: $bodyMode")
            $bru.Add('  auth: inherit')
            $bru.Add('}')
            $bru.Add('')
            if ($pathParamLines.Count -gt 0) {
                $bru.Add('params:path {')
                $bru.AddRange([string[]]$pathParamLines)
                $bru.Add('}')
                $bru.Add('')
            }
            if ($queryParamLines.Count -gt 0) {
                $bru.Add('params:query {')
                $bru.AddRange([string[]]$queryParamLines)
                $bru.Add('}')
                $bru.Add('')
            }
            if ($headersText) {
                $bru.Add('headers {')
                $bru.Add($headersText)
                $bru.Add('}')
                $bru.Add('')
            }
            if ($hasBody) {
                if ($bodyMode -eq 'json') {
                    $bru.Add('body:json {')
                    foreach ($line in $bodyJsonLines) { $bru.Add("  $line") }
                    $bru.Add('}')
                } else {
                    $bru.Add('body:text {')
                    $bru.Add('  ')
                    $bru.Add('}')
                }
                $bru.Add('')
            }
            $bru.Add('settings {')
            $bru.Add('  encodeUrl: true')
            $bru.Add('}')
            $bru.Add('')
            $bru.Add('docs {')
            foreach ($line in ($docLines -join "`n") -split "`r?`n") { $bru.Add("  $line") }
            $bru.Add('}')

            foreach ($exampleBlock in $exampleBlocks) {
                $bru.Add('')
                foreach ($line in ($exampleBlock -split "`r?`n")) { $bru.Add($line) }
            }

            [System.IO.File]::WriteAllText($bruPath, ($bru -join "`n"), [System.Text.UTF8Encoding]::new($false))
            $total++
        }
    }

    Write-Host "  Created $total Bruno request files under: $outRoot" -ForegroundColor Green
}

#-----------------------------------------------------------[Server / download helpers]-----------------------------------------------------------

function setupServerAuth {
    if ([string]::IsNullOrEmpty($script:clientId) -or [string]::IsNullOrEmpty($script:clientSecret) -or
        [string]::IsNullOrEmpty($script:tokenUrl)  -or [string]::IsNullOrEmpty($script:Server)) {
        $script:Server       = Read-Host -Prompt 'Enter the Workspace ONE UEM Server Name'
        $script:clientId     = Read-Host -Prompt 'Enter the OAuth Client ID'
        $script:clientSecret = Read-Host -Prompt 'Enter the OAuth Client Secret'
        $script:tokenUrl     = Read-Host -Prompt 'Enter the OAuth Token URL'
    }
    Write-Host "Requesting OAuth2 token..." -ForegroundColor Yellow
    try {
        $tokenBody     = @{ grant_type = 'client_credentials'; client_id = $script:clientId; client_secret = $script:clientSecret }
        $tokenResponse = Invoke-RestMethod -Method Post -Uri $script:tokenUrl -Body $tokenBody -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
    } catch {
        Write-Error "Failed to fetch token from $($script:tokenUrl): $_"
        exit 1
    }
    $accessToken = $null
    if ($tokenResponse -is [System.Management.Automation.PSCustomObject]) {
        if ($tokenResponse.access_token) { $accessToken = $tokenResponse.access_token }
        elseif ($tokenResponse.accessToken) { $accessToken = $tokenResponse.accessToken }
    }
    if (-not $accessToken) {
        Write-Error "No access token found in token response."
        exit 1
    }
    return $accessToken
}

function getServerVersion([string]$accessToken, [string]$serverBase) {
    Write-Host "Querying $serverBase/api/system/info ..." -ForegroundColor Yellow
    try {
        $headers = @{ Authorization = "Bearer $accessToken"; Accept = 'application/json'; Version = 1 }
        $infoRaw = Invoke-WebRequest -Uri "$serverBase/api/system/info" -Headers $headers -Method Get -ErrorAction Stop
        $infoObj = $infoRaw | ConvertFrom-Json -ErrorAction Stop
    } catch {
        Write-Error "Failed to get system info: $_"
        exit 1
    }
    $serverVersion = $infoObj.ProductVersion
    if (-not $serverVersion) {
        Write-Error "Could not determine server version from /api/system/info."
        exit 1
    }
    Write-Host "Server version: $serverVersion" -ForegroundColor Green
    return $serverVersion
}

# Downloads each endpoint's OpenAPI JSON into <collectionRoot>/<ep.Name>/<apiFile>.
# Returns the list of endpoint hashtables that downloaded successfully.
function downloadApiDocs([string]$serverBase, [array]$endpoints, [string]$collectionRoot, [string]$accessToken) {
    $headers    = @{ Authorization = "Bearer $accessToken" }
    $downloaded = @()
    foreach ($ep in $endpoints) {
        $apiFile  = $ep.Path.Split('/')[-1]            # e.g. "mamv1"
        $apiDir   = Join-Path $collectionRoot $ep.Name  # e.g. "<version>/MAM API V1"
        if (-not (Test-Path $apiDir)) { New-Item -ItemType Directory -Path $apiDir | Out-Null }
        $filePath = Join-Path $apiDir $apiFile           # e.g. "<version>/MAM API V1/mamv1"
        $url      = "$serverBase$($ep.Path)"
        Write-Host "Downloading $url -> $filePath" -ForegroundColor Gray
        try {
            Invoke-WebRequest -Uri $url -Headers $headers -OutFile $filePath -ErrorAction Stop
            $downloaded += $ep
        } catch {
            Write-Warning "Failed to download $($ep.Name): $_"
        }
    }
    return $downloaded
}

#-----------------------------------------------------------[Main]-----------------------------------------------------------

Write-Host "Workspace ONE UEM -> Bruno collection builder" -ForegroundColor Cyan
$current_path = $PSScriptRoot

$accessToken = setupServerAuth

# Normalise server URL
if ($script:Server -match '^https?://') {
    [string]$serverBase = $script:Server.TrimEnd('/')
} else {
    [string]$serverBase = "https://$($script:Server.TrimEnd('/'))"
}
Write-Host "Server: $serverBase" -ForegroundColor Green

$serverVersion = getServerVersion -accessToken $accessToken -serverBase $serverBase

# ---- Create versioned collection root ----
$collectionRoot = Join-Path $current_path $serverVersion
Write-Host "Creating collection folder: $collectionRoot" -ForegroundColor Yellow
if (-not (Test-Path $collectionRoot)) { New-Item -ItemType Directory -Path $collectionRoot | Out-Null }

$brunoJson = "{`n  `"version`": `"1`",`n  `"name`": `"$serverVersion`",`n  `"type`": `"collection`",`n  `"ignore`": [`"node_modules`", `".git`"]`n}"

$collectionBru = @"
meta {
  name: $serverVersion
}

auth {
  mode: oauth2
}

auth:oauth2 {
  grant_type: client_credentials
  access_token_url: {{Oauth_token_URL}}
  refresh_token_url: 
  client_id: {{Oauth_clientID}}
  client_secret: {{Oauth_client_secret}}
  scope: 
  credentials_placement: basic_auth_header
  credentials_id: 
  token_placement: header
  token_header_prefix: 
  auto_fetch_token: true
  auto_refresh_token: false
}

vars:pre-request {
  baseUrl: https://{{YOUR_API_SERVER}}/api
}

docs {
  Please refer to https://developer.omnissa.com/workspace-one-uem-apis/ for information on usage.
}
"@

Set-Content -Path (Join-Path $collectionRoot 'bruno.json')     -Value $brunoJson     -Encoding UTF8
Set-Content -Path (Join-Path $collectionRoot 'collection.bru') -Value $collectionBru -Encoding UTF8
Write-Host "Root collection files written." -ForegroundColor Green

# ---- API endpoints: UrlPrefix is inserted between {{baseUrl}} and the OpenAPI operation path ----
$endpoints = @(
  @{ Path = '/api/help/Docs/mamv1';    Name = 'MAM API V1';    Seq = 1;  UrlPrefix = '/mam';    Docs = 'Workspace ONE UEM MAM REST API V1' },
  @{ Path = '/api/help/Docs/mamv2';    Name = 'MAM API V2';    Seq = 2;  UrlPrefix = '/mam';    Docs = 'Workspace ONE UEM MAM REST API V2' },
  @{ Path = '/api/help/Docs/mcmv1';    Name = 'MCM API V1';    Seq = 3;  UrlPrefix = '/mcm';    Docs = 'Workspace ONE UEM MCM REST API V1' },
  @{ Path = '/api/help/Docs/mdmv1';    Name = 'MDM API V1';    Seq = 4;  UrlPrefix = '/mdm';    Docs = 'Workspace ONE UEM MDM REST API V1' },
  @{ Path = '/api/help/Docs/mdmv2';    Name = 'MDM API V2';    Seq = 5;  UrlPrefix = '/mdm';    Docs = 'Workspace ONE UEM MDM REST API V2' },
  @{ Path = '/api/help/Docs/mdmv3';    Name = 'MDM API V3';    Seq = 6;  UrlPrefix = '/mdm';    Docs = 'Workspace ONE UEM MDM REST API V3' },
  @{ Path = '/api/help/Docs/mdmv4';    Name = 'MDM API V4';    Seq = 7;  UrlPrefix = '/mdm';    Docs = 'Workspace ONE UEM MDM REST API V4' },
  @{ Path = '/api/help/Docs/memv1';    Name = 'MEM API V1';    Seq = 8;  UrlPrefix = '/mem';    Docs = 'Workspace ONE UEM MEM REST API V1' },
  @{ Path = '/api/help/Docs/systemv1'; Name = 'System API V1'; Seq = 9;  UrlPrefix = '/system'; Docs = 'Workspace ONE UEM System REST API V1' },
  @{ Path = '/api/help/Docs/systemv2'; Name = 'System API V2'; Seq = 10; UrlPrefix = '/system'; Docs = 'Workspace ONE UEM System REST API V2' }
)

# ---- Download ----
Write-Host "`nDownloading API docs..." -ForegroundColor Yellow
$downloaded = downloadApiDocs -serverBase $serverBase -endpoints $endpoints -collectionRoot $collectionRoot -accessToken $accessToken

# ---- Convert each downloaded doc to Bruno format ----
Write-Host "`nConverting API docs to Bruno format..." -ForegroundColor Yellow
foreach ($ep in $downloaded) {
    $apiFile = $ep.Path.Split('/')[-1]
    $apiDir  = Join-Path $collectionRoot $ep.Name
    $srcFile = Join-Path $apiDir $apiFile
    Write-Host "Converting $($ep.Name) ..." -ForegroundColor Yellow
    Convert-OpenApiToBruno `
        -srcFile   $srcFile `
        -outRoot   $apiDir `
        -apiName   $ep.Name `
        -apiSeq    $ep.Seq `
        -apiDocs   $ep.Docs `
        -urlPrefix $ep.UrlPrefix
}

Write-Host "`nDone. Collection created at: $collectionRoot" -ForegroundColor Cyan
