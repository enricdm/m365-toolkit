<#
.SYNOPSIS
  Pester tests for Platforms\_Shared\Modules\M365.Common.psm1.

.DESCRIPTION
  The module exists because of one bug: an empty Graph collection being read
  as a single object and injected as a phantom record, which once disabled a
  safety interlock on a bulk delete. The first test in this file pins that
  behaviour down forever; the rest cover paging, the item cap and throttling.

  Invoke-MgGraphRequest is stubbed - no tenant, no network, no credentials.

.NOTES
  Run: Invoke-Pester -Path Tests   (Pester 5+)
#>

BeforeAll {
    # A stub must exist before the module resolves the command name.
    function global:Invoke-MgGraphRequest { param($Method, $Uri, $ErrorAction) }

    $modulePath = Join-Path $PSScriptRoot '..\Platforms\_Shared\Modules\M365.Common.psm1'
    Import-Module $modulePath -Force
}

AfterAll {
    Remove-Module M365.Common -Force -ErrorAction SilentlyContinue
    Remove-Item function:Invoke-MgGraphRequest -ErrorAction SilentlyContinue
}

Describe 'Invoke-GraphPaged' {

    Context 'the phantom-record regression' {

        It 'returns an EMPTY list for an empty collection, not the raw response' {
            Mock Invoke-MgGraphRequest { @{ value = @() } } -ModuleName M365.Common

            $result = Invoke-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/things'

            $result.Count | Should -Be 0
        }

        It 'still returns an empty list when the empty page carries other keys' {
            Mock Invoke-MgGraphRequest {
                @{ value = @(); '@odata.context' = 'https://graph.microsoft.com/v1.0/$metadata#things' }
            } -ModuleName M365.Common

            (Invoke-GraphPaged -Uri 'https://x/things').Count | Should -Be 0
        }

        It 'treats a response without a value key as a single object' {
            Mock Invoke-MgGraphRequest { @{ id = 'abc'; displayName = 'One Thing' } } -ModuleName M365.Common

            $result = Invoke-GraphPaged -Uri 'https://x/things/abc'

            $result.Count | Should -Be 1
            $result[0].id | Should -Be 'abc'
        }

        It 'skips null entries inside a collection' {
            Mock Invoke-MgGraphRequest { @{ value = @(@{ id = '1' }, $null, @{ id = '2' }) } } -ModuleName M365.Common

            (Invoke-GraphPaged -Uri 'https://x/things').Count | Should -Be 2
        }
    }

    Context 'paging' {

        It 'follows @odata.nextLink until it runs out' {
            $script:page = 0
            Mock Invoke-MgGraphRequest {
                $script:page++
                if ($script:page -lt 3) {
                    @{ value = @(@{ id = "p$($script:page)" }); '@odata.nextLink' = "https://x/things?page=$($script:page + 1)" }
                } else {
                    @{ value = @(@{ id = "p$($script:page)" }) }
                }
            } -ModuleName M365.Common

            $result = Invoke-GraphPaged -Uri 'https://x/things'

            $result.Count | Should -Be 3
            $result[2].id | Should -Be 'p3'
            Should -Invoke Invoke-MgGraphRequest -ModuleName M365.Common -Times 3 -Exactly
        }

        It 'requests the URI it was given, then the nextLink verbatim' {
            $script:seen = [System.Collections.Generic.List[string]]::new()
            Mock Invoke-MgGraphRequest {
                $script:seen.Add($Uri)
                if ($script:seen.Count -eq 1) { @{ value = @(); '@odata.nextLink' = 'https://x/things?$skiptoken=abc' } }
                else { @{ value = @() } }
            } -ModuleName M365.Common

            $null = Invoke-GraphPaged -Uri 'https://x/things?$top=999'

            $script:seen[0] | Should -Be 'https://x/things?$top=999'
            $script:seen[1] | Should -Be 'https://x/things?$skiptoken=abc'
        }
    }

    Context 'MaxItems' {

        It 'trims the result to exactly MaxItems even mid-page' {
            Mock Invoke-MgGraphRequest {
                @{ value = @(1..10 | ForEach-Object { @{ id = "$_" } }) }
            } -ModuleName M365.Common

            (Invoke-GraphPaged -Uri 'https://x/things' -MaxItems 7).Count | Should -Be 7
        }

        It 'stops fetching pages once the cap is reached' {
            Mock Invoke-MgGraphRequest {
                @{ value = @(@{ id = 'a' }, @{ id = 'b' }); '@odata.nextLink' = 'https://x/more' }
            } -ModuleName M365.Common

            $result = Invoke-GraphPaged -Uri 'https://x/things' -MaxItems 2

            $result.Count | Should -Be 2
            Should -Invoke Invoke-MgGraphRequest -ModuleName M365.Common -Times 1 -Exactly
        }
    }

    Context 'throttling' {

        It 'retries after a throttle error and succeeds' {
            $script:calls = 0
            Mock Invoke-MgGraphRequest {
                $script:calls++
                if ($script:calls -eq 1) { throw 'Too Many Requests: request throttled' }
                @{ value = @(@{ id = 'ok' }) }
            } -ModuleName M365.Common
            Mock Start-Sleep {} -ModuleName M365.Common

            $result = Invoke-GraphPaged -Uri 'https://x/things' -WarningAction SilentlyContinue

            $result.Count | Should -Be 1
            Should -Invoke Invoke-MgGraphRequest -ModuleName M365.Common -Times 2 -Exactly
        }

        It 'does not retry a non-throttle error' {
            Mock Invoke-MgGraphRequest { throw 'Authorization_RequestDenied' } -ModuleName M365.Common
            Mock Start-Sleep {} -ModuleName M365.Common

            { Invoke-GraphPaged -Uri 'https://x/things' } | Should -Throw '*Authorization_RequestDenied*'
            Should -Invoke Invoke-MgGraphRequest -ModuleName M365.Common -Times 1 -Exactly
        }

        It 'does not mistake a 429 inside a GUID for throttling' {
            Mock Invoke-MgGraphRequest { throw 'Resource 5b1c4290-0000-0000-0000-000000000000 not found' } -ModuleName M365.Common
            Mock Start-Sleep {} -ModuleName M365.Common

            { Invoke-GraphPaged -Uri 'https://x/things' } | Should -Throw '*not found*'
            Should -Invoke Invoke-MgGraphRequest -ModuleName M365.Common -Times 1 -Exactly
        }
    }
}
