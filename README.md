# Invoke-NodeManage

This is a small-ish script that installs Chocolatey, PSWindowsUpdate, and then manages updates and installed apps based
on a directives.json file.

It's much cheaper/less confusing/more devops than using WorkspaceOne or Automox.

## Prerequisites

There's a bunch of pre-requisites to make this work:
* Some sort of AD domain
* Some sort of domain service account to run the local scheduled task
* A GPO to deploy script, directives.json, and set a scheduled task
* If you're rocking a private (internal) choco feed, you'll want to install Chocolatey.server
* If you're splitting your custom choco packages into code and binary installer, you'll want a separate http fileserver (I used nginx on Ubuntu) since choco can't pull from a SMB share

## Invoke-NodeManage.ps1

The script has a few features:
* Logging to eventlog
* Installs choco and pswindowsupdate to the local machine
* Managed and unmanaged choco installs. Managed means pinned to prevent drift, like if you're maintaining a validated system
* Reboots
* Can pull all machines from an OU
* Very fast ADSISearcher function
* Default 'no-network' mode, will still attempt updates and reboots
* Sharding large deployments, so you're not rebooting/updating a billion nodes at the same time
* Uses a JSON file to determine actions/policies to follow

## Directives.json

The directives.json file contains all the policies you want for either an OU full of machines, or an individual machine. Individual machine policy will override the OU policy, so that special snowflake BMS server in the basement get the special treatment it deserves.

It is also a monolithic file, so it can get very large. It's not the best design, but it was easier for helpdesk and other sysadmins to deal with a single file vs. individual files.

## Set-NodeManageTask

Since deploying a scheduled task via GPO has been hobbled due to security reasons, we deploy an XML file and import it into schtasks. A bit hacky, but it does work.