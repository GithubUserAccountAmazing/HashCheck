# Changelog.md

## [1.1.0] - 2023-04-11
### Changed
- Changed the method of calculating the SHA256 hash of the file and the certificate from certUtil.exe to Get-FileHash cmdlet
- This improves the security and reliability of the script
- No changes to the functionality or logic of the script
- Changed formatting, added comments.

## [1.0.0] - 2022-07-07
### Added
- Initial release of the script
- The script checks the integrity of an Excel file and a certificate using SHA256 hash
- If the hashes match, it imports the certificate to the TrustedPublisher store and opens the Excel file
- If the hashes do not match, it displays a warning message and aborts
- The script uses a splash image and a hidden PowerShell window to run
