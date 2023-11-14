#!/bin/bash -x

# Login to Azure
az login

# Set your subscription
az account set --subscription <SubscriptionId>

# Create a service principal and captures the output
sp=$(az ad sp create-for-rbac --name 'sp-Azure-Advisor-Excel-Reporter' --years 1)

# Extract the values from the service principal output
clientId=$(echo $sp | jq -r .appId)
clientSecret=$(echo $sp | jq -r .password)
tenantId=$(echo $sp | jq -r .tenant)
subscriptionId=<SubscriptionId>

# Write these values into a .env file
echo "AZURE_CLIENT_ID=$clientId" > .env
echo "AZURE_TENANT_ID=$tenantId" >> .env
echo "AZURE_CLIENT_SECRET=$clientSecret" >> .env
echo "AZURE_SUBSCRIPTION_ID=$subscriptionId" >> .env

# Assign the service principal to the subscription Reader role
az role assignment create --assignee $clientId --role Reader