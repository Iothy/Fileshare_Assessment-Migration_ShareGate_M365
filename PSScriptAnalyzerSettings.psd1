@{
    Severity = @('Error', 'Warning')
    ExcludeRules = @(
        'PSAvoidUsingWriteHost',
        'PSReviewUnusedParameter',
        'PSUseShouldProcessForStateChangingFunctions'
    )
    Rules = @{
        PSUseConsistentIndentation = @{
            Enable = $true
            Kind = 'space'
            IndentationSize = 4
        }
        PSUseConsistentWhitespace = @{
            Enable = $true
            CheckOpenBrace = $true
            CheckOpenParen = $true
            CheckSeparator = $true
            CheckPipe = $true
            CheckOperator = $true
            CheckInnerBrace = $true
        }
    }
}
