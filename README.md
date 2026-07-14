# DomainAdminTool.exe Official README
A standalone, lightweight Windows desktop domain & account management executable tool, dedicated to efficient domain asset account management, including account creation, deletion, batch import/export, and daily asset maintenance.
<img width="1182" height="896" alt="image" src="https://github.com/user-attachments/assets/556e55ef-2c5e-4430-985e-cc930d6a8313" />
<img width="1186" height="891" alt="image" src="https://github.com/user-attachments/assets/0602acc2-b9af-44f3-954e-ce64fed32f8c" />
<img width="1184" height="893" alt="image" src="https://github.com/user-attachments/assets/17c6530c-6d98-4362-9bb6-6dc6bb99dd9d" />


## 1. Project Overview
DomainAdminTool.exe is a compiled Windows standalone executable program developed for enterprise operation and maintenance personnel, domain administrators and asset managers. Compared with script-based tools, it eliminates the need for Python environment configuration, supports one-click direct operation on Windows systems, and provides visualized and streamlined domain account management functions.
The tool focuses on solving the problems of scattered domain account information, manual repeated entry, difficult batch management, and untraceable asset data. It covers full-cycle account management capabilities such as account creation, invalid account cleaning, and data batch import/export, helping users standardize and automate domain asset account management.
## 2. Core Functional Highlights
- Account Deletion Cleaning: Support manual single deletion and batch selective deletion of invalid/expired/abnormal accounts, with secondary confirmation mechanism to avoid accidental deletion.
- Batch Data Import: Support importing domain account data through Excel/CSV template, quickly bulk adding account assets and realizing one-click data synchronization.
- Batch Data Export: Freely export all or screened domain account data to local files, support Excel/CSV format, convenient for asset archiving, reconciliation and offline backup.
- Data Visual Display: Automatically count the total number of accounts, valid/invalid accounts, newly added and deleted accounts, and display asset status in real time.
- Lightweight & Efficient: Small program volume, low system resource occupation, fast response speed, suitable for long-term stable operation on office and server devices.
- Local Data Storage: All account data is stored locally, no network data upload, ensuring the security and privacy of enterprise domain asset information.
## 3. Operating Environment
- Supported System: Windows 10 / Windows 11 / Windows Server 2019 / Windows Server 2022 and above
- Operation Permission: Standard user permission (administrator permission is recommended for full data read and write access)
- Network Environment: Offline local operation is supported; partial network access is required for real-time domain status verification
- Disk Space: Minimum 50MB free space (for program operation and data storage)
## 4. Quick Start
4.1 Program Launch
1. Download the latest DomainAdminTool.exe program file
2. Place the program in a non-system disk directory (avoid data loss after system reset)
3. Double-click DomainAdminTool.exe to launch the tool, no installation required
4.2 Initial Configuration
After the first launch, the tool will automatically generate a local configuration file and data storage folder. Users can modify basic parameters such as data export format and default storage path according to usage habits.
## 5. Detailed Core Operation Guide
5.1 Account Creation (Single & Batch)
Single Account Creation
1. Launch the tool and enter the Account Management module
8. Click the Add Single Account button
7. Fill in account information: account name, bound domain, permission level, contact person, validity period and remarks
6. Confirm the information and click Save, the account will be automatically added to the local asset list
Batch Account Creation
1. In the Account Management module, click Batch Add Accounts
7. Download the official standard import template (Excel/CSV)
6. Fill in multiple account information in batches according to the template format (do not modify the template header)
7. Upload the completed template file, and the tool will automatically identify and create accounts in batches
7. Check the creation log to confirm successful and failed account data
5.2 Account Deletion (Single & Batch)
Single Account Deletion
1. In the account list, screen and select the target account to be deleted
6. Click the Delete button at the back of the account
5. Confirm the secondary verification prompt to complete the single account deletion
Batch Account Deletion
1. Check multiple invalid/expired accounts in the account list in batches
8. Click the Batch Delete function button
7. Verify the account list to be deleted to avoid misoperation
6. Confirm deletion, and the tool will clear the selected account data in one click
Note: Deleted account data can be recovered through backup files (please export data backup regularly).
5.3 Data Import Function
This function is used to quickly import external domain account asset data to realize data migration and batch update:
1. Enter the Data Import &amp; Export module
10. Click Import Data, support Excel/CSV format file import
10. Select the local data file that conforms to the template specification
9. The tool will automatically verify data format and duplicate accounts
8. After verification, select Overwrite Update or Append Add to complete data import
5.4 Data Export Function
Support flexible export of account asset data for daily asset sorting and backup:
1. Enter the Data Import & Export module
5. Screen the account data that needs to be exported (support full export and conditional screening export)
9. Select the export format: Excel (recommended) / CSV
8. Customize the local save path and file name
7. Click Export Now to complete one-click data export
## 6. File Directory Description
After running the program, the tool will automatically generate the following directory structure for data management:
DomainAdminTool/
├── DomainAdminTool.exe      # Main executable program of domain account management tool
├── DomainServers.txt        # Global config file: stores domain server addresses, default export format, local data storage paths and global tool settings
├── KPinyin.dll              # Dynamic link library dependency file, provides Chinese Pinyin conversion capability for account name generation and remark processing
└── pinyin_config.txt        # Pinyin mapping configuration file, defines Pinyin conversion rules used by KPinyin.dll during account creation

## 7. Important Notes
- It is strictly prohibited to use this tool for illegal account scanning, unauthorized access and other illegal behaviors. Users shall bear all legal responsibilities for improper use.
- Please back up data regularly through the export function to avoid permanent loss of account data caused by accidental deletion or program exception.
- Do not modify the configuration file and template file at will, which may lead to program operation failure and data import/export errors.
- Batch operation (batch creation/deletion/import/export) is recommended to be carried out in a stable network and system environment to prevent data abnormality.
- The program runs locally, and all data will not be uploaded to the network, ensuring the security of private domain asset data.
## 8. Troubleshooting Common Problems
- Program cannot be opened: Check whether the system version meets the requirements, turn off firewall/antivirus software interception, and run the program as administrator.
- Data import failure: Confirm that the imported file format is correct, the template header is not modified, and the data content has no format error.
- Data export failure: Check whether the target save path has write permission and sufficient disk space.
- Account data loss: Recover data through the automatic backup files in the data/ directory or manually exported backup files.
## 9. License & Support
This tool is open source and free for personal and enterprise non-commercial use. For secondary development and commercial use, please comply with the open source agreement. If you have functional optimization suggestions or usage problems, you can submit feedback to continuously improve the tool.




