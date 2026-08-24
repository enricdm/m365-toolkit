@{
    # ===========================================================================
    #  SAMPLE domain -> country mapping data.
    #
    #  Every domain below is a placeholder (contoso.*). Replace them with your
    #  own e-mail domains before running anything that consumes this file, or
    #  point the consuming script at your own copy with -DomainMapPath /
    #  --domain-map-path.
    #
    #  Consumed by:
    #    Platforms/EntraID/Applications/Update-AppContact.ps1
    #    Platforms/EntraID/Identity/Get-MissingUsageLocation.ps1
    #
    #  Format note: keep this file to plain hashtables and string arrays with
    #  single-quoted values, and no expressions. That subset can be parsed by a
    #  non-PowerShell reader without evaluating the file, which is worth
    #  preserving if this mapping is ever consumed from another language.
    # ===========================================================================

    # ---------------------------------------------------------------------------
    # DomainToCountry - e-mail domain -> ISO 3166-1 alpha-2 country code.
    # Use this wherever the value has to be a real ISO code, e.g. stamping
    # usageLocation or comparing against sign-in country.
    # ---------------------------------------------------------------------------
    DomainToCountry  = @{
        'contoso.de'    = 'DE'
        'contoso.es'    = 'ES'
        'contoso.it'    = 'IT'
        'contoso.fr'    = 'FR'
        'contoso.pt'    = 'PT'
        'contoso.be'    = 'BE'
        'contoso.ch'    = 'CH'
        'contoso.cz'    = 'CZ'
        'contoso.sk'    = 'SK'
        'contoso.hu'    = 'HU'
        'contoso.ro'    = 'RO'
        'contoso.se'    = 'SE'
        'contoso.fi'    = 'FI'
        'contoso.dk'    = 'DK'
        'contoso.ee'    = 'EE'
        'contoso.lt'    = 'LT'
        'contoso.lu'    = 'LU'
        'contoso.co.uk' = 'GB'
        'contoso.com.br' = 'BR'
        'contoso.mx'    = 'MX'
        'contoso.cl'    = 'CL'
        'contoso.pe'    = 'PE'
        'contoso.co'    = 'CO'
        'contoso.ec'    = 'EC'
        'contoso.ph'    = 'PH'
        'contoso.ae'    = 'AE'
        'contoso.ng'    = 'NG'
        'contoso.gh'    = 'GH'
    }

    # ---------------------------------------------------------------------------
    # CentralDomains - domains shared by more than one country. A user on one of
    # these cannot be placed by domain alone; the consumer falls back to sign-in
    # geography or usageLocation. Subdomains of these entries count as central.
    # ---------------------------------------------------------------------------
    CentralDomains   = @(
        'contoso.com'
        'contoso.onmicrosoft.com'
        'contoso.local'
    )

    # ---------------------------------------------------------------------------
    # CountryToDomains - internal naming-convention code -> preferred e-mail
    # domains, most likely first. Used to guess a person's mailbox address from
    # their name plus the country their account naming implies.
    #
    # NOTE: the keys here are *naming-convention* codes, not ISO codes. They are
    # whatever two-letter prefix your app/group naming standard uses (this sample
    # keeps the common 'UK' instead of ISO 'GB', and adds HQ/GL for central teams).
    # ---------------------------------------------------------------------------
    CountryToDomains = @{
        'HQ' = @('contoso.com')
        'GL' = @('contoso.com')
        'DE' = @('contoso.com', 'contoso.de')
        'CH' = @('contoso.com', 'contoso.ch')
        'CZ' = @('contoso.com', 'contoso.cz')
        'PH' = @('contoso.com')
        'ES' = @('contoso.es', 'contoso.com')
        'IT' = @('contoso.it', 'contoso.com')
        'EE' = @('contoso.ee', 'contoso.com')
        'SE' = @('contoso.se', 'contoso.com')
        'FI' = @('contoso.fi', 'contoso.com')
        'DK' = @('contoso.dk', 'contoso.com')
        'FR' = @('contoso.fr', 'contoso.com')
        'PT' = @('contoso.pt', 'contoso.com')
        'BE' = @('contoso.be', 'contoso.com')
        'BR' = @('contoso.com.br', 'contoso.com')
        'UK' = @('contoso.co.uk', 'contoso.com')
        'CO' = @('contoso.co', 'contoso.com')
        'EC' = @('contoso.ec', 'contoso.com')
    }

    # ---------------------------------------------------------------------------
    # TldToCountry - public-suffix fragment of an app's homepage URL -> naming
    # -convention code (same key space as CountryToDomains). Weakest signal;
    # only used when nothing better is available.
    # ---------------------------------------------------------------------------
    TldToCountry     = @{
        'es'    = 'ES'
        'it'    = 'IT'
        'de'    = 'DE'
        'fr'    = 'FR'
        'fi'    = 'FI'
        'se'    = 'SE'
        'pt'    = 'PT'
        'ee'    = 'EE'
        'lt'    = 'LT'
        'ec'    = 'EC'
        'ph'    = 'PH'
        'dk'    = 'DK'
        'be'    = 'BE'
        'ch'    = 'CH'
        'cz'    = 'CZ'
        'co.uk' = 'UK'
    }

    # ---------------------------------------------------------------------------
    # TeamDomainToCode - domain of a shared/team mailbox -> naming-convention
    # code. Same key space as CountryToDomains, so a team alias discovered in the
    # data can be attributed to the country team that owns it. The central domain
    # maps to the global team rather than to a country.
    # ---------------------------------------------------------------------------
    TeamDomainToCode = @{
        'contoso.es'    = 'ES'
        'contoso.it'    = 'IT'
        'contoso.de'    = 'DE'
        'contoso.ee'    = 'EE'
        'contoso.se'    = 'SE'
        'contoso.fi'    = 'FI'
        'contoso.fr'    = 'FR'
        'contoso.pt'    = 'PT'
        'contoso.dk'    = 'DK'
        'contoso.co.uk' = 'UK'
        'contoso.com'   = 'GL'
    }
}
