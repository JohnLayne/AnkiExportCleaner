# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Office Add-ins integration with custom "Anki Tools" ribbon
- Yeoman generator-office project structure
- Node.js backend API server for Python integration
- Four functional ribbon buttons (Import, Export, Validate, Help)
- Comprehensive documentation with proper attributions
- Git-friendly project structure with proper .gitignore
- GUID preservation for Anki card relationship maintenance
- File naming with -CLEANED suffix for exported files

### Changed
- Modernized from VBA-only to Office Add-ins primary approach
- Updated README.md with Office Add-ins setup instructions
- Improved project structure with dual-approach support

### Fixed
- File explosion issue with proper dependency exclusion
- Directory naming (spaces to hyphens for git compatibility)
- Documentation accuracy and completeness

## [1.0.0] - Legacy VBA Approach

### Added
- Initial VBA-based Excel integration
- HTML cleaning and UTF-8 encoding support
- Quick Access Toolbar and custom ribbon VBA implementations
- Croatian diacritics support and testing
- Basic Python script with Excel export functionality
