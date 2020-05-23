# Windows Inventory Discovery

Windows Inventory Discovery, is a script that generates inventory documents for every computer it has been executed on. It was developed to automate the creation of inventory documents regarding 160+ computers of an internal network.

## Getting Started

After cloning the repo, execute the `main.ps1` script file.

### Prerequisites

Make sure you have installed all of the following prerequisites on your machine:

* Windows 7 or Windows 8/8.1 or Windows 10
* Microsoft Word 2003 and above
* PowerShell 2.0 and above

### How it works

The script collects information about the machine it has been executed on such as:
* CPU
* RAM
* Hard Disk
* GPU
* Operating System
* Type (Desktop or Laptop)
* Manufacturer
* Computer Name

After every successful execution it saves the inventory document in the `reports` folder. The name of the file is in the format of `{computername}.docx`. It is also given an ID number which is saved in the `list.csv` file.

For every action such as start of the execution, document creation, finish e.t.c. a log entry is created in the `app_log.log` file.

### Domain Execution
It is advised to create a policy in the Active Directory Domain Controller to execute the script after user login. In case the script has been executed before, the execution stops right before collecting the information, so a new report is not generated.

## Built With

* [PowerShell](https://docs.microsoft.com/en-us/powershell/) - A task automation and configuration management framework from Microsoft, consisting of a command-line shell and associated scripting language. 

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details
