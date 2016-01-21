# Change Log
All notable changes to this project will be documented in this file.

## [1.0.3c] - TBA
### Fixed
- Fixed Supplier Parts Master: Box Specs/M3 typecast error
- Cleanup TTC Parts: WEST Field check
- **BUGFIX #5** - Fixed false positive in Customer Contract Details: Customer Contract check

## [1.0.3b] - 2016-01-19
### Added
- Added extra logic in Supplier Contract: WH Code check to see if WH Code already registered before
- Moved pre-ARS and 'Check All Master Sheets?' to option flags
⋅⋅* Run program on command line with '-h' or '--help' flag to see list of options
- **HotFix**: Recompiled into .exe

### Fixed
- Fixed false positive in Module Group Code check for IN codes
- **BUGFIX #4** - Rewrote Module Group: Shipping Frequency check to be more accurate
⋅⋅* Rewrote all Shipping Frequency checks
- **BUGFIX #2** - Fixed false positives in TTC Parts: Material Tax Class
- Added GM check for MOD West fields
- Added Gross Weight check for MOD Customer Parts: Next_SPQ

## [1.0.3] - 2016-01-13
### Added
- Added ability to reference .xlsx for Global Master check
- Added ability to reference post-update GM in Results Folder (only .xlsx)
- Added extra logic to discern between S500 and non-S500 if there is discrepancy in GM data
- Added check if 'NEW/MOD' fields have values aside from 'NEW' and 'MOD' (includes whitespace)
- Added Supplier Code and Exp Country reference for Customer Parts Master: Exp Back No. check
- Added additional information for NEW entries that are already registered in System
- Added 'Null' to list of Exp Back No. in Customer Parts Master: Exp Back No. check to skip check
- Added extra logic in Container Group: WH Code check to see if WH Code already registered before
- Added extra check for MOD Customer/Supplier Parts - Raise error if part has already been discontinued
- Added extra logic in Container Group: Container Type check to see if Container Type already registered before
- Added extra logic in Module Group: WH Code check to see if WH Code already registered before
- Added extra logic in Part Master: WEST Fields to consider multiple Imp/Exp Country scenario

### Fixed
- Fixed error in Customer Parts: Imp HS Code logic to show PASS for <same as exp>
- Strip whitespace from 'NEW/MOD' fields for more accurate classification of rows
- Fixed error in Customer Contract Details: Module Group that caused program to crash occasionally
- Fixed error that causes non-ASCII filenames to crash program
- Fixed false positives in Parts Master: WEST Fields (Exp) for C1 parts.
- Fixed false negative in Customer/Supplier Parts Master: Parts Master WEST check for WESt-optional TW parts
- Fixed error in Customer Contract: Cross Dock Flag referencing the wrong cell, causing program to crash
- Fixed error in TTC Contract: WEST Exp Sales No./Imp Purchase No. that caused program to crash
- Fixed error in Customer Contract Details: TTC Contract check for PK parts that caused program to crash
- Fixed false positives in TTC Contract: Mid E-Signature Flag
- Fixed errors in all Shipping Route checks, causing first row to not be referenced.
- Fixed errors in Compulsory Field check that caused program to crash
- Fixed syntax error in Module Group Master: Module Type check that caused program to crash
- Rewrote Customer Contract Details: TTC Contract check for Module Groups
- Fixed error in Container Group: Source Port/Destination Port check
- Fixed false positive in Supplier Parts Master: TTC Parts No. check for part in Customer Contract Details

## [1.0.2] - 2015-12-22
### Added
- Added additional checkpoint for Customer Contract Details: Customer Parts Name to tally with Customer Parts Master
- Added 'Press any key to exit' to prevent auto-close of command line
- Added additional checkpoint for Customer Parts Master: Paired Parts (Show error for if TTC P/N Paired Parts different OL from Customer Part)
- Added additional checkpoint for Customer Parts Master: Part No. (Show error if no WEST Field registered for Imp WEST Customer in Parts Master)
- Added additional checkpoint for Supplier Parts Master: Part No. (Show error if no WEST Field registered for Exp WEST Supplier in Parts Master)
- Added additional Suppliers (TH-TBJ1 and TH-TBSJ2) that require Back No. check in Supplier Parts Master

### Changed
- Rewrote logic for Customer Contract Details: Supplier Contract to be more robust
- Rewrote logic for Inner Packing BOM: Material No. (Now only prompts WARNING if Material No. is completely new)
- Rewrote logic for Inner Packing BOM: Sequence No. (Properly detects duplicate sequence no.)
- Rewrote logic for Parts Master: WEST Fields (Properly accounts for multiple rows in Parts Master)

### Fixed
- Fixed error in get MOD backup row logic that caused program to crash
- Rewrote Supplier Contract: WEST Fields due to it not properly detecting errors

## [1.0.1a] - 2015-12-15
### Fixed
- Fixed error in Customer Parts: Paired Parts that caused program to crash
- Fixed false positives in Customer Parts: Imp Country Code for 'IN, I1, I2, I3' region
- Fixed false positives in Customer Parts: WEST Invoice No. for 'I2' region

## [1.0.1] - 2015-12-15
### Changed
- Both .xls and .XLS files are detected in '1) Submit'

### Removed
- Removed command line output for non-english parts that caused program to crash

### Fixed
- Fixed logic that caused some masters to be incorrectly recognised as not ALL MOD
- Fixed error in Module Group Master check that causes Module Type Master to not be loaded properly
- Fixed error in Parts Master: Exp HS Code that caused program to crash
- Fixed error in Supplier Contract: West Fields that caused program to crash
- Fixed error in Inner Packing BOM: SPQ that caused program to crash
- Fixed error in Customer Parts: Next_SPQ that caused program to crash

## [1.0.0 Stable] - 2015-12-15
### Added
- First build of MCT
