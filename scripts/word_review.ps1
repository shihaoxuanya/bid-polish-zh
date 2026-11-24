#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Original,
    [Parameter(Mandatory = $true)][string]$Tracked,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Author,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [string]$ExpectedAccepted,
    [ValidateRange(0, 2147483647)][int]$ExpectedRevisionCount
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-PackageXml($entry) {
    $stream = $entry.Open()
    try {
        $xml = New-Object System.Xml.XmlDocument
        $xml.XmlResolver = $null
        $settings = New-Object System.Xml.XmlReaderSettings
        $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
        $reader = [System.Xml.XmlReader]::Create($stream, $settings)
        try { $xml.Load($reader) } finally { $reader.Dispose() }
        return ,$xml
    } finally { $stream.Dispose() }
}

function Resolve-WordInput([string]$fileName) {
    $file = Get-Item -LiteralPath $fileName
    if ($file.PSIsContainer -or $file.Extension -ine '.docx') {
        throw "Expected a local .docx file: $fileName"
    }
    $archive = [System.IO.Compression.ZipFile]::OpenRead($file.FullName)
    try {
        if (-not $archive.GetEntry('word/document.xml')) { throw 'Missing Word document part.' }
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -match '(^|/)(vbaProject\.bin|activeX/|webextensions/)') {
                throw "Interactive content is unsuitable for unattended review: $($entry.FullName)"
            }
            if ($entry.FullName -match '\.rels$') {
                $xml = Read-PackageXml $entry
                foreach ($rel in $xml.DocumentElement.ChildNodes) {
                    if ($rel.GetAttribute('TargetMode') -ieq 'External' -and
                        $rel.GetAttribute('Type') -notmatch '/hyperlink$') {
                        throw "External content link in $($entry.FullName). Word review was not started."
                    }
                }
            } elseif ($entry.FullName -match '^word/[^/]+\.xml$') {
                $xml = Read-PackageXml $entry
                if ($xml.SelectSingleNode("//*[local-name()='altChunk']")) { throw 'Embedded document import requires manual review.' }
                # Field commands can span multiple runs. A field boundary separates adjacent commands.
                $codes = New-Object System.Text.StringBuilder
                foreach ($node in $xml.SelectNodes("//*[local-name()='instrText' or local-name()='fldChar' or local-name()='fldSimple']")) {
                    if ($node.LocalName -eq 'instrText') { [void]$codes.Append($node.InnerText) }
                    elseif ($node.LocalName -eq 'fldSimple') {
                        [void]$codes.AppendLine()
                        [void]$codes.AppendLine($node.GetAttribute('instr', $node.NamespaceURI))
                    } else { [void]$codes.AppendLine() }
                }
                if ($codes.ToString() -match '(?im)^\s*(DDEAUTO|DDE|LINK|INCLUDETEXT|INCLUDEPICTURE|DATABASE|RD)\s') {
                    throw "External-data field in $($entry.FullName). Word review was not started."
                }
            }
        }
    } finally { $archive.Dispose() }
    return $file.FullName
}

function Open-WordFile($app, [string]$fileName, [bool]$readOnly) {
    $doc = $app.Documents.Open($fileName, $false, $readOnly, $false)
    if (-not $readOnly -and $doc.ReadOnly) { throw "Could not open the QA copy for editing: $fileName" }
    # Render Word's final view, leaving all original revisions and comments in the file.
    $doc.Windows.Item(1).View.Type = 3
    $doc.Windows.Item(1).View.RevisionsView = 0
    $doc.Windows.Item(1).View.ShowRevisionsAndComments = $false
    return ,$doc
}

function Get-TextHash([string]$value) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($value)
        return [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '')
    } finally { $sha.Dispose() }
}

function Get-DocumentHash([string]$fileName) {
    $stream = [IO.File]::OpenRead($fileName)
    try {
        $sha = [Security.Cryptography.SHA256]::Create()
        try { return [BitConverter]::ToString($sha.ComputeHash($stream)) }
        finally { $sha.Dispose() }
    } finally { $stream.Dispose() }
}

function Get-WordState($doc, [string]$batchAuthor) {
    $stories = New-Object 'System.Collections.Generic.List[string]'
    $seenStories = @{}
    $revisions = New-Object 'System.Collections.Generic.List[string]'
    $comments = New-Object 'System.Collections.Generic.List[string]'
    $batchCount = 0
    $batchComments = 0
    foreach ($first in $doc.StoryRanges) {
        $story = $first
        $instance = 0
        while ($null -ne $story) {
            $view = $story.Duplicate
            # Compare Word's field instructions instead of cached page/date results.
            $view.TextRetrievalMode.IncludeFieldCodes = $true
            $view.TextRetrievalMode.IncludeHiddenText = $true
            $storyType = [int]$view.StoryType
            $key = '{0}:{1}' -f $storyType, $instance
            $storyText = ([string]$view.Text).Replace([string][char]5, '')
            $digest = Get-TextHash $storyText
            if ($seenStories.ContainsKey($key)) {
                # Word can enumerate the very same linked header/footer twice after layout.
                if ($seenStories[$key] -cne $digest) { throw "Conflicting Word story at $key" }
                $instance++
                $story = $story.NextStoryRange
                continue
            }
            $seenStories[$key] = $digest
            if ($storyType -ne 4 -and $storyType -le 11) {
                # Comments are compared separately, including each original anchored quotation.
                # Repagination in Word can materialize unused, empty header/footer stories
                # and duplicate built-in note separators (types 12+). Neither is document prose.
                if ($storyType -lt 6 -or $storyText.TrimEnd([char]13).Length -gt 0) {
                    $stories.Add(('{0}:{1}' -f $key, $digest))
                }
            }
            foreach ($rev in $story.Revisions) {
                if ($rev.Author -ceq $batchAuthor) { $batchCount++ }
                else {
                    $revisions.Add(('{0}|{1}|{2}|{3}|{4}' -f $story.StoryType, $rev.Author,
                        [int]$rev.Type, $rev.Date.Ticks, (Get-TextHash ([string]$rev.Range.Text))))
                }
            }
            $instance++
            $story = $story.NextStoryRange
        }
    }
    foreach ($comment in $doc.Comments) {
        if ($comment.Author -ceq $batchAuthor) { $batchComments++; continue }
        $comments.Add(('{0}|{1}|{2}|{3}' -f $comment.Author, $comment.Date.Ticks,
            (Get-TextHash ([string]$comment.Range.Text)), (Get-TextHash ([string]$comment.Scope.Text))))
    }
    $stories.Sort([StringComparer]::Ordinal)
    $revisions.Sort([StringComparer]::Ordinal)
    $comments.Sort([StringComparer]::Ordinal)
    return [ordered]@{
        story_hash = Get-TextHash ($stories -join "`n")
        revision_hash = Get-TextHash ($revisions -join "`n")
        comment_hash = Get-TextHash ($comments -join "`n")
        other_revisions = $revisions.Count
        comments = $comments.Count
        batch_revisions = $batchCount
        batch_comments = $batchComments
    }
}

function Resolve-Batch($doc, [string]$batchAuthor, [bool]$accept) {
    $doc.TrackRevisions = $false
    $count = 0
    foreach ($first in $doc.StoryRanges) {
        $story = $first
        while ($null -ne $story) {
            for ($i = $story.Revisions.Count; $i -ge 1; $i--) {
                if ($i -gt $story.Revisions.Count) { continue }
                $rev = $story.Revisions.Item($i)
                if ($rev.Author -cne $batchAuthor) { continue }
                if ($accept) { $rev.Accept() } else { $rev.Reject() }
                $count++
            }
            $story = $story.NextStoryRange
        }
    }
    return $count
}

function Get-Page($range) {
    $start = $range.Duplicate
    $start.Collapse(1)
    return [ordered]@{
        physical_page = [int]$start.Information(3)
        adjusted_page = [int]$start.Information(1)
        section = [int]$start.Information(2)
    }
}

function Get-Excerpt([string]$value) {
    $line = ($value -replace '[\x00-\x1f]+', ' ').Trim()
    if ($line.Length -gt 160) { return $line.Substring(0, 160) }
    return $line
}

function Get-PageMap($doc) {
    $doc.Repaginate()
    $headings = New-Object 'System.Collections.Generic.List[object]'
    $sections = New-Object 'System.Collections.Generic.List[object]'
    $pointers = New-Object 'System.Collections.Generic.List[object]'
    $tocs = New-Object 'System.Collections.Generic.List[object]'
    foreach ($section in $doc.Sections) { $sections.Add((Get-Page $section.Range)) }
    foreach ($toc in $doc.TablesOfContents) {
        $range = $toc.Range
        $point = Get-Page $range
        $point.start = $range.Start
        $point.end = $range.End
        $point.cached_text = [string]$range.Text
        $tocs.Add($point)
    }
    $paragraphs = $doc.Paragraphs
    $paragraphCount = $paragraphs.Count
    for ($i = 1; $i -le $paragraphCount; $i++) {
        $paragraph = $paragraphs.Item($i)
        $range = $paragraph.Range
        $level = [int]$paragraph.OutlineLevel
        $text = ([string]$range.Text).Trim()
        if ($level -ge 1 -and $level -le 9) {
            $point = Get-Page $range
            $point.text = Get-Excerpt $text
            $point.level = $level
            $point.paragraph = $i
            $headings.Add($point)
        }
        # Literal references, manually typed TOC lines and page-index column headers.
        if ($text -match '(?i)\u7b2c\s*[0-9\u4e00-\u9fff-]{1,10}\s*\u9875|\b[Pp]\.?\s*\d+\b|[.\u2026\u00b7\t]{2,}\s*\d+\s*$|\u9875[\u7801\u6b21]|\u54cd\u5e94\u7d22\u5f15|\u8bc4\u5206\u7d22\u5f15') {
            $inToc = $false
            foreach ($tocPoint in $tocs) {
                if ($range.Start -ge $tocPoint.start -and $range.Start -lt $tocPoint.end) { $inToc = $true; break }
            }
            if (-not $inToc) {
                $point = Get-Page $range
                $point.paragraph = $i
                $point.text = Get-Excerpt $text
                $point.in_table = [bool]$range.Information(12)
                $pointers.Add($point)
            }
        }
    }
    return [ordered]@{ pages = [int]$doc.ComputeStatistics(2); sections = $sections; headings = $headings
        automatic_tocs = $tocs; page_reference_candidates = $pointers }
}

function Add-Check([string]$name, [bool]$ok) {
    $report.checks[$name] = $ok
    if (-not $ok) { throw "Verification failed: $name" }
}

$created = $false
$word = $null
$originalDoc = $null
$trackedDoc = $null
$expectedDoc = $null
$acceptedDoc = $null
$rejectedDoc = $null
$oldOptions = @{}
$report = [ordered]@{ status = 'failed'; checks = [ordered]@{}; states = [ordered]@{}; layout = [ordered]@{} }
$failure = $null
try {
    $originalPath = Resolve-WordInput $Original
    $trackedPath = Resolve-WordInput $Tracked
    if ($originalPath -eq $trackedPath) { throw 'Original and Tracked must be distinct.' }
    if ($ExpectedAccepted) { $expectedPath = Resolve-WordInput $ExpectedAccepted }
    $qaPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
    if (Test-Path -LiteralPath $qaPath) { throw "Use a new QA directory: $qaPath" }
    [void](New-Item -Path $qaPath -ItemType Directory -ErrorAction Stop)
    $created = $true
    $acceptedPath = Join-Path $qaPath 'accepted.docx'
    $rejectedPath = Join-Path $qaPath 'rejected.docx'
    $originalHash = Get-DocumentHash $originalPath
    $trackedHash = Get-DocumentHash $trackedPath
    [IO.File]::Copy($trackedPath, $acceptedPath, $false)
    [IO.File]::Copy($trackedPath, $rejectedPath, $false)
    (Get-Item -LiteralPath $acceptedPath).IsReadOnly = $false
    (Get-Item -LiteralPath $rejectedPath).IsReadOnly = $false
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    foreach ($key in 'UpdateLinksAtOpen','UpdateLinksAtPrint','UpdateFieldsAtPrint') {
        $oldOptions[$key] = $word.Options.$key
        $word.Options.$key = $false
    }
    $originalDoc = Open-WordFile $word $originalPath $true
    $source = Get-WordState $originalDoc $Author
    $report.states.original = $source
    Add-Check 'unique_batch_author_absent_in_source' ($source.batch_revisions -eq 0 -and $source.batch_comments -eq 0)
    $report.layout.original = Get-PageMap $originalDoc
    # Pass a plain CLR string, not a provider-decorated PowerShell path object, to Word COM.
    [string]$originalPdf = [IO.Path]::Combine($qaPath, 'original.pdf')
    $originalDoc.ExportAsFixedFormat($originalPdf, 17)
    $originalDoc.Close(0)
    $originalDoc = $null
    $trackedDoc = Open-WordFile $word $trackedPath $true
    $trackedState = Get-WordState $trackedDoc $Author
    $report.states.tracked = $trackedState
    if ($PSBoundParameters.ContainsKey('ExpectedRevisionCount')) {
        Add-Check 'expected_word_revision_count' ($trackedState.batch_revisions -eq $ExpectedRevisionCount)
    }
    Add-Check 'prior_revisions_and_comments_retained_in_tracked' (
        $source.revision_hash -eq $trackedState.revision_hash -and $source.comment_hash -eq $trackedState.comment_hash)
    $trackedDoc.Close(0)
    $trackedDoc = $null
    $rejectedDoc = Open-WordFile $word $rejectedPath $false
    $report.rejections = Resolve-Batch $rejectedDoc $Author $false
    $rejectedDoc.Save()
    $rejectedDoc.Close(0)
    $rejectedDoc = Open-WordFile $word $rejectedPath $true
    $rejected = Get-WordState $rejectedDoc $Author
    $report.states.rejected = $rejected
    Add-Check 'reject_removes_batch_and_restores_source_text' (
        $rejected.batch_revisions -eq 0 -and $rejected.story_hash -eq $source.story_hash)
    Add-Check 'prior_revisions_and_comments_retained_after_reject' (
        $source.revision_hash -eq $rejected.revision_hash -and $source.comment_hash -eq $rejected.comment_hash)
    $rejectedDoc.Close(0)
    $rejectedDoc = $null
    $acceptedDoc = Open-WordFile $word $acceptedPath $false
    $report.acceptances = Resolve-Batch $acceptedDoc $Author $true
    $acceptedDoc.Save()
    $acceptedDoc.Close(0)
    $acceptedDoc = Open-WordFile $word $acceptedPath $false
    $accepted = Get-WordState $acceptedDoc $Author
    $report.states.accepted_before_toc = $accepted
    Add-Check 'accept_removes_batch' ($accepted.batch_revisions -eq 0)
    Add-Check 'accepted_text_matches_tracked_final_view' ($accepted.story_hash -eq $trackedState.story_hash)
    Add-Check 'prior_revisions_and_comments_retained_after_accept' (
        $source.revision_hash -eq $accepted.revision_hash -and $source.comment_hash -eq $accepted.comment_hash)
    $report.expected_accepted_compared = [bool]$ExpectedAccepted
    if ($ExpectedAccepted) {
        $expectedDoc = Open-WordFile $word $expectedPath $true
        $expected = Get-WordState $expectedDoc $Author
        $report.states.expected = $expected
        Add-Check 'accepted_text_matches_independent_expected' ($accepted.story_hash -eq $expected.story_hash)
        $expectedDoc.Close(0)
        $expectedDoc = $null
    }
    $report.layout.accepted_before_toc = Get-PageMap $acceptedDoc
    foreach ($toc in $acceptedDoc.TablesOfContents) { $toc.Update() }
    $acceptedDoc.Repaginate()
    foreach ($toc in $acceptedDoc.TablesOfContents) { $toc.UpdatePageNumbers() }
    $acceptedDoc.Save()
    $report.layout.accepted_after_toc = Get-PageMap $acceptedDoc
    $refreshed = Get-WordState $acceptedDoc $Author
    Add-Check 'prior_revisions_and_comments_retained_after_toc_refresh' (
        $source.revision_hash -eq $refreshed.revision_hash -and $source.comment_hash -eq $refreshed.comment_hash)
    [string]$acceptedPdf = [IO.Path]::Combine($qaPath, 'accepted.pdf')
    $acceptedDoc.ExportAsFixedFormat($acceptedPdf, 17)
    $acceptedDoc.Close(0)
    $acceptedDoc = $null
    Add-Check 'source_and_tracked_files_unchanged' (
        (Get-DocumentHash $originalPath) -eq $originalHash -and
        (Get-DocumentHash $trackedPath) -eq $trackedHash)
    $report.status = 'passed'
} catch {
    $failure = $_.Exception.Message
    $report.error = $failure
    $report.error_location = $_.ScriptStackTrace
} finally {
    foreach ($doc in @($expectedDoc, $acceptedDoc, $rejectedDoc, $trackedDoc, $originalDoc)) {
        if ($null -ne $doc) { try { $doc.Close(0) } catch {} }
    }
    if ($null -ne $word) {
        foreach ($key in $oldOptions.Keys) { try { $word.Options.$key = $oldOptions[$key] } catch {} }
        try { $word.Quit() } catch {
            if (-not $failure) {
                $failure = "Could not close the review Word instance: $($_.Exception.Message)"
                $report.error = $failure
                $report.status = 'failed'
            }
        }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
    if ($created) {
        $reportPath = Join-Path $qaPath 'review.json'
        [IO.File]::WriteAllText($reportPath, ($report | ConvertTo-Json -Depth 10), (New-Object Text.UTF8Encoding($false)))
        Write-Output $reportPath
    }
}
if ($failure) { [Console]::Error.WriteLine($failure); exit 1 }
