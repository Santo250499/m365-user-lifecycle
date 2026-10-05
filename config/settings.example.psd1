@{
    # Sample placeholders. Copy this file to config/settings.psd1 (gitignored)
    # and replace the ids there. Do not put a client secret in this file.
    TenantId              = '00000000-0000-0000-0000-000000000000'
    ClientId              = '00000000-0000-0000-0000-000000000000'
    CertificateThumbprint = ''
    AuthMode              = 'Interactive'
    UsageLocation         = 'AU'
    LogDirectory          = 'logs'
    Scopes                = @(
        'User.ReadWrite.All'
        'GroupMember.ReadWrite.All'
        'Organization.Read.All'
        'LicenseAssignment.ReadWrite.All'
        'User.RevokeSessions.All'
    )
}
