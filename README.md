
# Azure Advisor Excel Report Generator

<img src="https://img.shields.io/badge/Azure%20CLI%20-v2.53.1-blue?style=flat-square">   <img src="https://img.shields.io/badge/VSCode%20-v1.84.2-purple?style=flat-square">   <img src="https://img.shields.io/badge/Python%20-v3.10.12-darkblue?style=flat-square">

## Getting Started

This project is a Python script that generates an Excel workbook report from your Azure Advisor recommendations.

## Example report

![4snetHub-3vnetSpoke-Arch.png](/examples/4snetHub-3vnetSpoke-Arch.png)

## Prerequisites

Before you begin, ensure you have met the following requirements:

- An **Azure subscription**.
- Latest version of **Azure CLI**.
- A **Python** environment with the following libraries installed:
  - pandas
  - azure-identity
  - azure.mgmt.advisor
  - dotenv
  - datetime
  - jinja2 (Version 3.1.2+)

_Tested with Python3.10.12 & AzCli2.53.1_

## Using Azure Advisor Excel Report Generator

To use Azure Advisor Excel Report Generator, follow these steps:

1. Clone this repository to your local machine.

```bash
git clone git@github.com:zakarel/Azure-Advisor-Excel-Reporter.git
```

2. Run the bash script in your client

```bash
bash pre-script.sh
```
This will create a .env file in the current directory that will hold credentials to authenticate to azure and give it the reader role

3. Install the required Python libraries using pip:

```bash
   pip install pandas azure-identity azure.mgmt.advisor python-dotenv datetime jinja2
```