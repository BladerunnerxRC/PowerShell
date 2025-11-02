@{
    IncludeDefaultRules = $true

    ExcludeRules = @(
        'PSAvoidUsingWriteHost' # We intentionally use Write-Host for an interactive, colorized UI
    )

    Rules = @{
        PSUseConsistentIndentation = @{
            Enable = $true
            IndentationSize = 4
            PipelineIndentation = 'IncreaseIndentationAfterEveryPipeline'
            Kind = 'space'
        }
        PSUseConsistentWhitespace = @{ Enable = $true }
        PSPlaceOpenBrace           = @{ Enable = $true; OnSameLine = $true }
        PSPlaceCloseBrace          = @{ Enable = $true; NoEmptyLineBefore = $true }
        PSUseConsistentQuoteMarks  = @{ Enable = $true; QuotePreference = 'Single' }
    }
}