
# Azure Advisor Excel Report Generator

<img src="https://img.shields.io/badge/Azure%20CLI%20-v2.53.1-blue?style=flat-square">   <img src="https://img.shields.io/badge/VSCode%20-v1.84.2-purple?style=flat-square">   <img src="https://img.shields.io/badge/Python%20-v3.10.12-darkblue?style=flat-square">

This Repository is a Python script that generates an Excel workbook report from your Azure Advisor recommendations.

## Example report

Example I - Overview sheet (png)
![azure-advisor-excel-reporter-overview-example.png](/examples/azure-advisor-excel-reporter-overview-example.png)

Example II - Security recommendations (png)
![azure-advisor-excel-reporter-sec-example.png](/examples/azure-advisor-excel-reporter-sec-example.png)

Example III - Full example workbook (xlsx)
![azure-advisor-excel-reporter-example.xlsx](/examples/azure-advisor-excel-reporter-example.xlsx)

## Prerequisites

Before you begin, ensure you have met the following requirements:

- You own an **Azure subscription**.
- Latest version of **Azure CLI**.
- A **Python** environment with the following libraries installed:
  - pandas
  - azure-identity
  - azure.mgmt.advisor
  - dotenv
  - datetime
  - jinja2 (Version 3.1.2+)
- jq 1.6 (JSON processor to prettify things up )

_Tested with Python3.10.12 & AzCli2.53.1_

## Getting Started

This project is a Python script that generates an Excel workbook report from your Azure Advisor recommendations.

To use Azure Advisor Excel Report Generator, follow these steps:

### 1. Clone this repository to your local machine.

```bash
git clone git@github.com:zakarel/Azure-Advisor-Excel-Reporter.git
```

### 2. Configure & Run the pre-script.sh that will execute the following:

 * Creation of a service principle with the name: "sp-Azure-Advisor-Excel-Reporter" which will be vaild for 1 year (you can renew the password after that time)
 * Creation of an environment hidden file (.env)
 * Extraction of the service principle authentication credentials and appending them into the environment file.
 * Assigning the subscription reader role to the service principle

- Changing the _<SubscriptionId>_ to the desired Subscription ID

- Executing the script with:
```bash
bash pre-script.sh
```

### 3. Install the required Python libraries using pip:

```bash
pip install pandas azure-identity azure.mgmt.advisor python-dotenv datetime jinja2
```

### 4. Install the latest jq on the client:

_Debian/Ubuntu/Kali based_
```bash
sudo apt-get install jq
```

_Fedora based_
```bash
sudo yum install jq
```

_Homebrew(Mac)_
```bash
brew install jq
```

### 5. Running the Script:

```bash
python3 azure-advisor-excel-reporter.py
```


_note: If you are getting a Permission denied when trying to execute any of the script first make sure you have granted it the execute permission_
```bash
chmod +x pre-script.sh
chmod +x azure-advisor-excel-reporter.py
```

## Authors

* **Tzahi Ariel** - *Initial work* - [zakarel](https://github.com/zakarel)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details

