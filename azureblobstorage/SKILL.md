---
name: azure-blob-uploader
description: Uploads a local file to Azure Blob Storage and returns a SAS read link.
user-invocable: true
metadata:
  requires:
    env:
      - OPENCLAW_AZURE_STORAGE_CONNECTION_STRING
    bins:
      - python3
---

# Azure Blob Uploader Skill

When the user asks to "upload this file to Azure" or "give me a SAS link":

## Prerequisites
- `python3` must be available on `PATH`.
- The Python package `azure-storage-blob` must be installed.
- Ensure the `OPENCLAW_AZURE_STORAGE_CONNECTION_STRING` environment variable is available.

## 1. Input Validation
- Identify the target file in the current workspace.


## 2. Execution Workflow
Use the `python3` tool to execute a script that performs these steps:
1. Connect to Azure Blob Storage using the connection string.
2. Upload the specified file to a container named `openclaw-files` (create it if it doesn't exist).
3. Generate a Shared Access Signature (SAS) token with `read` permissions valid for 24 hours.
4. Construct the full URL: `https://<account>.blob.core.windows.net/<container>/<blob>?<sas_token>`.

## 3. Output
- Return the final SAS URL to the user in the channel terminal.
- Confirm the file has been successfully uploaded.

## 4. Failure Handling
- If the required python package is missing, instruct the user to install it using `pip install azure-storage-blob`.
- If the upload fails, show the specific error from the Azure SDK.
- If the file is missing, ask the user to provide the file first.
